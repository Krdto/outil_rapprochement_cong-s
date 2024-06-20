from flask import Flask, request, redirect, url_for, send_from_directory, render_template, flash
import os
import pandas as pd
import xlsxwriter
from datetime import timedelta

app = Flask(__name__)
app.secret_key = "infokey"
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def compare_dates_multiple_sheets(reference_file, comparison_file, output_file):
    try:
        print(f"Comparaison des dates entre {reference_file} et {comparison_file} dans {output_file}")
        
        df1 = pd.read_excel(comparison_file, sheet_name=None)
        df2 = pd.read_excel(reference_file)
        df2['Début'] = pd.to_datetime(df2['Début'], format='%Y-%m-%d', errors='coerce').dt.date
        df2['Fin'] = pd.to_datetime(df2['Fin'], format='%Y-%m-%d', errors='coerce').dt.date
        
        if df2['Début'].isnull().any():
            raise ValueError("Certaines dates de début sont mal formées dans le fichier de comparaison.")
        if df2['Fin'].isnull().any():
            raise ValueError("Certaines dates de fin sont mal formées dans le fichier de comparaison.")
        
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
        
        for sheet_name, df1_sheet in df1.items():
            df1_sheet['Date à contrôler'] = pd.to_datetime(df1_sheet['Date à contrôler'], format='%d/%m/%Y', errors='coerce').dt.date
            if df1_sheet['Date à contrôler'].isnull().any():
                raise ValueError(f"Certaines dates à contrôler sont mal formées dans la feuille {sheet_name} du fichier de référence.")
            results_df = pd.DataFrame(columns=['Matricule', 'Date à contrôler', 'Result'])
            for index, row1 in df1_sheet.iterrows():
                matricule = row1['Matricule']
                date_a_controler = row1['Date à contrôler']
                df2_filtered = df2[df2['Matricule'] == matricule]
                result = "La date ne correspond pas"
                for _, row2 in df2_filtered.iterrows():
                    current_date = row2['Début']
                    while current_date <= row2['Fin']:
                        if date_a_controler == current_date:
                            result = row2['Libellé']
                            break
                        current_date += timedelta(days=1)
                    if result != "La date ne correspond pas":
                        break
                results_df.loc[len(results_df)] = [matricule, date_a_controler, result]
            results_df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.set_column('B:B', 18)
        writer.close()
    except Exception as e:
        print(f"Une erreur s'est produite : {e}")

def compare_dates(reference_file, comparison_file, output_file):
    try:
        print(f"Comparaison des dates entre {reference_file} et {comparison_file} dans {output_file}")
        df1 = pd.read_excel(comparison_file)
        df2 = pd.read_excel(reference_file)
        required_columns_df1 = {'Matricule', 'Date de début', 'Date de fin'}
        required_columns_df2 = {'Matricule', 'Début', 'Fin', 'Libellé'}
        if not required_columns_df1.issubset(df1.columns) or not required_columns_df2.issubset(df2.columns):
            raise ValueError("Les fichiers ne contiennent pas les colonnes recherchées.")
        df1['Date de début'] = pd.to_datetime(df1['Date de début'], format='%d/%m/%Y', errors='coerce').dt.date
        df1['Date de fin'] = pd.to_datetime(df1['Date de fin'], format='%d/%m/%Y', errors='coerce').dt.date
        df2['Début'] = pd.to_datetime(df2['Début'], format='%Y-%m-%d', errors='coerce').dt.date
        df2['Fin'] = pd.to_datetime(df2['Fin'], format='%Y-%m-%d', errors='coerce').dt.date
        if df1['Date de début'].isnull().any():
            raise ValueError("Certaines dates de début sont mal formées dans le fichier de référence.")
        if df1['Date de fin'].isnull().any():
            raise ValueError("Certaines dates de fin sont mal formées dans le fichier de référence.")
        if df2['Début'].isnull().any():
            raise ValueError("Certaines dates de début sont mal formées dans le fichier de comparaison.")
        if df2['Fin'].isnull().any():
            raise ValueError("Certaines dates de fin sont mal formées dans le fichier de comparaison.")
        expanded_df1 = []
        for index, row in df1.iterrows():
            current_date = row['Date de début']
            end_date = row['Date de fin']
            while current_date <= end_date:
                expanded_df1.append({'Matricule': row['Matricule'], 'Date': current_date})
                current_date += timedelta(days=1)
        expanded_df1 = pd.DataFrame(expanded_df1)
        expanded_df2 = []
        for index, row in df2.iterrows():
            current_date = row['Début']
            end_date = row['Fin']
            while current_date <= end_date:
                expanded_df2.append({'Matricule': row['Matricule'], 'Date': current_date, 'Libellé': row['Libellé']})
                current_date += timedelta(days=1)
        expanded_df2 = pd.DataFrame(expanded_df2)
        merged_df = pd.merge(expanded_df1, expanded_df2, on=['Matricule', 'Date'], how='left')
        merged_df['Libellé'] = merged_df['Libellé'].fillna('La date ne correspond pas')
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
        merged_df.to_excel(writer, sheet_name='Results', index=False)
        worksheet = writer.sheets['Results']
        worksheet.set_column('A:D', 18)
        writer.close()
    except Exception as e:
        print(f"Une erreur s'est produite : {e}")

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    print("Upload route accessed")
    if request.method == 'POST':
        print("POST request received")
        if 'file1' not in request.files or 'file2' not in request.files or 'file3' not in request.files:
            flash('No file part', 'danger')
            print("Missing file part")
            return redirect(request.url)
        file1 = request.files['file1']
        file2 = request.files['file2']
        file3 = request.files['file3']
        
        if file1.filename == '' or file2.filename == '' or file3.filename == '':
            flash('No selected file', 'danger')
            print("No selected file")
            return redirect(request.url)
        if file1 and allowed_file(file1.filename) and file2 and allowed_file(file2.filename) and file3 and allowed_file(file3.filename):
            filename1 = os.path.join(app.config['UPLOAD_FOLDER'], file1.filename)
            filename2 = os.path.join(app.config['UPLOAD_FOLDER'], file2.filename)
            filename3 = os.path.join(app.config['UPLOAD_FOLDER'], file3.filename)
            file1.save(filename1)
            file2.save(filename2)
            file3.save(filename3)
            
            print(f"Files saved: {filename1}, {filename2}, {filename3}")
            
            output_file1 = os.path.join(app.config['UPLOAD_FOLDER'], f'output_{file2.filename}.xlsx')
            output_file2 = os.path.join(app.config['UPLOAD_FOLDER'], f'output_{file3.filename}.xlsx')
            
            compare_dates_multiple_sheets(filename1, filename2, output_file1)
            compare_dates_multiple_sheets(filename1, filename3, output_file2)
            
            flash('Les fichiers ont été comparés avec succès!', 'success')
            print("Comparison successful")
            
            # Delete uploaded files after processing
            os.remove(filename1)
            os.remove(filename2)
            os.remove(filename3)
            
            return redirect(url_for('download_file', filename=f'output_{file2.filename}.xlsx'))

    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)
    except FileNotFoundError:
        flash('File not found', 'danger')
        app.logger.error(f"File '{filename}' not found in directory '{app.config['UPLOAD_FOLDER']}'")
        return redirect(url_for('upload_files'))
    except Exception as e:
        flash(f"Error downloading file: {e}", 'danger')
        app.logger.error(f"Error downloading file '{filename}': {e}")
        return redirect(url_for('upload_files'))


if __name__ == "__main__":
    app.run(debug=True)