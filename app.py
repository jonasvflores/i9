import os
import pandas as pd
import re
import calendar
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import logging

# Configuração de logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("debug.log"),
        logging.StreamHandler()
    ]
)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SECRET_KEY'] = 'supersecretkey'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Limite de 16MB para uploads

# Criar diretório uploads se não existir
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    try:
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    except Exception as e:
        logging.error(f"Erro ao criar diretório uploads: {e}")
        raise

# Lista para armazenar os arquivos enviados e os meses associados
files_data = []

# Meses disponíveis para seleção
MONTHS = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
          "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

@app.route('/', methods=['GET', 'POST'])
def index():
    global files_data
    if request.method == 'POST':
        # Verificar se arquivos foram enviados
        if 'files[]' not in request.files:
            flash('Nenhum arquivo enviado.')
            return redirect(request.url)

        files = request.files.getlist('files[]')
        csv_month = request.form.get('csv_month')
        excel_month = request.form.get('excel_month')
        excel_year = request.form.get('excel_year', '2025')

        if not files:
            flash('Nenhum arquivo selecionado.')
            return redirect(request.url)

        for file in files:
            if file and file.filename:
                try:
                    filename = secure_filename(file.filename)
                    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    file.save(file_path)
                    file_type = "CSV" if filename.endswith('.csv') else "Excel"
                    if file_type == "CSV":
                        month_name = csv_month if csv_month in MONTHS else "Janeiro"
                        files_data.append((file_path, file_type, None, None, month_name))
                    else:
                        try:
                            month = int(excel_month)
                            if not 1 <= month <= 12:
                                raise ValueError
                            year = int(excel_year) if excel_year.isdigit() else 2025
                            month_name = MONTHS[month - 1]
                            files_data.append((file_path, file_type, month, year, month_name))
                        except ValueError:
                            flash('Mês ou ano inválido para o arquivo Excel.')
                            os.remove(file_path) if os.path.exists(file_path) else None
                            return redirect(request.url)
                except Exception as e:
                    flash(f"Erro ao processar o arquivo {filename}: {str(e)}")
                    os.remove(file_path) if os.path.exists(file_path) else None
                    return redirect(request.url)

        return render_template('index.html', files=files_data, months=MONTHS)

    files_data = []  # Limpar a lista de arquivos ao carregar a página
    return render_template('index.html', files=files_data, months=MONTHS)

@app.route('/remove/<int:index>')
def remove_file(index):
    global files_data
    if 0 <= index < len(files_data):
        file_path = files_data[index][0]
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception as e:
                logging.error(f"Erro ao remover arquivo {file_path}: {e}")
        files_data.pop(index)
    return redirect(url_for('index'))

@app.route('/combine', methods=['POST'])
def combine_files():
    global files_data
    if not files_data:
        flash('Nenhum arquivo selecionado para combinar.')
        return redirect(url_for('index'))

    all_data = []
    expected_columns = ["Dia_Semana", "Data", "Hora", "Aluno", "Professor", "Observações", "CRM_Esteira"]

    for file_path, file_type, month, year, month_name in files_data:
        try:
            if file_type == "CSV":
                df = pd.read_csv(file_path, encoding="utf-8")
                if not all(col in df.columns for col in expected_columns):
                    flash(f"O arquivo CSV {os.path.basename(file_path)} não contém as colunas esperadas.")
                    return redirect(url_for('index'))
                df["Mês"] = month_name
                all_data.append(df)
            else:
                xls = pd.ExcelFile(file_path)
                excel_data = []
                columns = ["HORA", "ALUNO", "PROFESSOR", "OBSERVAÇÕES / BIO", "CRM / ESTEIRA"]

                for sheet_name in xls.sheet_names:
                    day_of_month = extract_day_from_sheet_name(sheet_name)
                    if day_of_month is None:
                        continue

                    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                    header_row = None
                    for i, row in df.iterrows():
                        if row.astype(str).str.contains("HORA", case=False, na=False).any():
                            header_row = i
                            break

                    if header_row is None:
                        continue

                    df.columns = df.iloc[header_row].str.strip()
                    df = df.iloc[header_row + 1:].reset_index(drop=True)

                    available_columns = [col for col in df.columns if isinstance(col, str)]
                    matched_columns = {}
                    for expected_col in columns:
                        for col in available_columns:
                            if expected_col.lower() in col.lower():
                                matched_columns[expected_col] = col
                                break

                    if not all(col in matched_columns for col in columns):
                        continue

                    df = df.rename(columns={matched_columns[col]: col for col in matched_columns})
                    df = df[columns]
                    df["Data"] = f"{int(day_of_month):02d}/{month:02d}/{year}"
                    df["Dia_Semana"] = df["Data"].apply(get_weekday_name)

                    professor_columns = [col for col in df.columns if col.startswith("PROFESSOR")]
                    if len(professor_columns) > 1:
                        df["PROFESSOR"] = df[professor_columns].apply(
                            lambda x: ", ".join([str(val).strip() for val in x if pd.notna(val) and str(val).strip() != ""]), axis=1
                        )
                        df = df.drop(columns=[col for col in professor_columns if col != "PROFESSOR"])

                    df["HORA"] = df["HORA"].ffill()
                    df["HORA"] = df["HORA"].apply(convert_time)
                    df = df.dropna(subset=["ALUNO", "PROFESSOR", "OBSERVAÇÕES / BIO", "CRM / ESTEIRA"], how="all")
                    df = df[["Dia_Semana", "Data", "HORA", "ALUNO", "PROFESSOR", "OBSERVAÇÕES / BIO", "CRM / ESTEIRA"]]
                    df["Mês"] = month_name
                    excel_data.append(df)

                if excel_data:
                    excel_df = pd.concat(excel_data, ignore_index=True)
                    excel_df.columns = ["Dia_Semana", "Data", "Hora", "Aluno", "Professor", "Observações", "CRM_Esteira", "Mês"]
                    all_data.append(excel_df)

        except Exception as e:
            flash(f"Erro ao processar o arquivo {os.path.basename(file_path)}: {str(e)}")
            return redirect(url_for('index'))

    if not all_data:
        flash("Nenhum dado válido encontrado nos arquivos selecionados.")
        return redirect(url_for('index'))

    final_df = pd.concat(all_data, ignore_index=True)
    final_df = final_df[["Dia_Semana", "Data", "Hora", "Aluno", "Professor", "Observações", "CRM_Esteira", "Mês"]]

    # Ordenar por Mês, Data e Hora
    month_order = MONTHS
    final_df["Mês"] = pd.Categorical(final_df["Mês"], categories=month_order, ordered=True)
    final_df = final_df.sort_values(by=["Mês", "Data", "Hora"])

    output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'combined_output.csv')
    final_df.to_csv(output_path, index=False, encoding="utf-8")

    # Limpar arquivos temporários após o processamento
    for file_path, _, _, _, _ in files_data:
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception as e:
                logging.error(f"Erro ao remover arquivo temporário {file_path}: {e}")

    files_data = []  # Limpar a lista após o download
    return send_file(output_path, as_attachment=True, download_name='combined_output.csv')

def extract_day_from_sheet_name(sheet_name):
    try:
        day_number = sheet_name.split()[-1]
        return int(day_number)
    except (IndexError, ValueError):
        return None

def convert_time(time_value):
    try:
        if pd.isna(time_value) or str(time_value).strip() == "" or str(time_value).lower() == "nan":
            return ""
        time_str = str(time_value).strip()
        if not time_str:
            return ""
        if re.match(r"^\d{1,2}:\d{2}$", time_str):
            hours, minutes = map(int, time_str.split(":"))
            if 0 <= hours <= 23 and 0 <= minutes <= 59:
                return f"{hours:02d}:{minutes:02d}"
            else:
                logging.warning(f"Valor de hora fora do intervalo (string): {time_str}")
                return ""
        if isinstance(time_value, (int, float)):
            hours = int(float(time_value) * 24)
            minutes = int((float(time_value) * 24 - hours) * 60)
            if 0 <= hours <= 23 and 0 <= minutes <= 59:
                return f"{hours:02d}:{minutes:02d}"
            else:
                logging.warning(f"Valor de hora fora do intervalo (float): {time_value}")
                return ""
        try:
            time_float = float(time_str)
            hours = int(time_float * 24)
            minutes = int((time_float * 24 - hours) * 60)
            if 0 <= hours <= 23 and 0 <= minutes <= 59:
                return f"{hours:02d}:{minutes:02d}"
            else:
                logging.warning(f"Valor de hora fora do intervalo (float from string): {time_float}")
                return ""
        except ValueError:
            pass
        logging.warning(f"Formato de hora não reconhecido: {time_str} (tipo: {type(time_value)})")
        return str(time_value)
    except Exception as e:
        logging.error(f"Erro ao converter hora: {time_value} (tipo: {type(time_value)}), Erro: {str(e)}")
        return str(time_value)

def get_weekday_name(date_str):
    try:
        dia, mes, ano = map(int, date_str.split('/'))
        dia_semana = calendar.weekday(ano, mes, dia)
        nomes_dias = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
        return nomes_dias[dia_semana]
    except Exception:
        return ""

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
