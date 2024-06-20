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

# Crée le dossier d'upload s'il n'existe pas
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def allowed_file(filename):
    """
    Vérifie si un fichier a une extension autorisée.

    Args:
        filename (str): Le nom du fichier.

    Returns:
        bool: True si le fichier a une extension autorisée, False sinon.
    """
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def compare_dates_multiple_sheets(reference_file, comparison_file, output_file):
    """
    Compare les dates entre un fichier de référence et un fichier de comparaison contenant plusieurs feuilles,
    et écrit les résultats dans un fichier de sortie.

    Args:
        reference_file (str): Chemin du fichier de référence.
        comparison_file (str): Chemin du fichier de comparaison avec plusieurs feuilles.
        output_file (str): Chemin du fichier de sortie pour les résultats.
    """
    try:
        print(f"Comparaison des dates entre {reference_file} et {comparison_file} dans {output_file}")

        # Lecture des fichiers Excel
        df1 = pd.read_excel(comparison_file, sheet_name=None)
        df2 = pd.read_excel(reference_file)
        
        # Standardise les noms de colonnes du dataframe de référence
        df2.columns = [col.strip().lower() for col in df2.columns]
        df2 = df2.rename(columns={
            'début': 'start_date',
            'fin': 'end_date',
            'matricule': 'matricule',
            'libellé': 'label'
        })
        
        # Conversion des colonnes de dates
        df2['start_date'] = pd.to_datetime(df2['start_date'], format='%Y-%m-%d', errors='coerce').dt.date
        df2['end_date'] = pd.to_datetime(df2['end_date'], format='%Y-%m-%d', errors='coerce').dt.date
        
        # Vérification de la validité des dates
        if df2['start_date'].isnull().any():
            raise ValueError("Certaines dates de début sont mal formées dans le fichier de référence.")
        if df2['end_date'].isnull().any():
            raise ValueError("Certaines dates de fin sont mal formées dans le fichier de référence.")

        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

        # Traitement de chaque feuille de comparaison
        for sheet_name, df1_sheet in df1.items():
            df1_sheet.columns = [col.strip().lower() for col in df1_sheet.columns]
            if 'date à contrôler' in df1_sheet.columns:
                df1_sheet = df1_sheet.rename(columns={'date à contrôler': 'control_date', 'matricule': 'matricule'})
            else:
                raise ValueError(f"La colonne 'Date à contrôler' est manquante dans la feuille {sheet_name} du fichier de comparaison.")
            
            # Conversion des colonnes de dates
            df1_sheet['control_date'] = pd.to_datetime(df1_sheet['control_date'], format='%d/%m/%Y', errors='coerce').dt.date
            if df1_sheet['control_date'].isnull().any():
                raise ValueError(f"Certaines dates à contrôler sont mal formées dans la feuille {sheet_name} du fichier de comparaison.")
            
            results_df = pd.DataFrame(columns=['Matricule', 'Date à contrôler', 'Result'])
            for index, row1 in df1_sheet.iterrows():
                matricule = row1['matricule']
                control_date = row1['control_date']
                df2_filtered = df2[df2['matricule'] == matricule]
                result = "La date ne correspond pas"
                for _, row2 in df2_filtered.iterrows():
                    current_date = row2['start_date']
                    while current_date <= row2['end_date']:
                        if control_date == current_date:
                            result = row2['label']
                            break
                        current_date += timedelta(days=1)
                    if result != "La date ne correspond pas":
                        break
                results_df.loc[len(results_df)] = [matricule, control_date, result]
            results_df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.set_column('B:B', 18)
        writer.close()
    except Exception as e:
        print(f"Une erreur s'est produite : {e}")

def compare_dates(reference_file, comparison_file, output_file):
    """
    Compare les dates entre un fichier de référence et un fichier de comparaison,
    et écrit les résultats dans un fichier de sortie.

    Args:
        reference_file (str): Chemin du fichier de référence.
        comparison_file (str): Chemin du fichier de comparaison.
        output_file (str): Chemin du fichier de sortie pour les résultats.
    """
    try:
        print(f"Comparaison des dates entre {reference_file} et {comparison_file} dans {output_file}")
        df1 = pd.read_excel(comparison_file)
        df2 = pd.read_excel(reference_file)
        
        # Standardise les noms de colonnes des dataframes
        df1.columns = [col.strip().lower() for col in df1.columns]
        df1 = df1.rename(columns={
            'date de début': 'start_date',
            'date de fin': 'end_date',
            'matricule': 'matricule'
        })
        df2.columns = [col.strip().lower() for col in df2.columns]
        df2 = df2.rename(columns={
            'début': 'start_date',
            'fin': 'end_date',
            'matricule': 'matricule',
            'libellé': 'label'
        })
        
        # Vérification des colonnes nécessaires
        required_columns_df1 = {'matricule', 'start_date', 'end_date'}
        required_columns_df2 = {'matricule', 'start_date', 'end_date', 'label'}
        
        if not required_columns_df1.issubset(df1.columns) or not required_columns_df2.issubset(df2.columns):
            raise ValueError("Les fichiers ne contiennent pas les colonnes recherchées.")
        
        # Conversion des colonnes de dates
        df1['start_date'] = pd.to_datetime(df1['start_date'], format='%d/%m/%Y', errors='coerce').dt.date
        df1['end_date'] = pd.to_datetime(df1['end_date'], format='%d/%m/%Y', errors='coerce').dt.date
        df2['start_date'] = pd.to_datetime(df2['start_date'], format='%Y-%m-%d', errors='coerce').dt.date
        df2['end_date'] = pd.to_datetime(df2['end_date'], format='%Y-%m-%d', errors='coerce').dt.date
        
        # Vérification de la validité des dates
        if df1['start_date'].isnull().any():
            raise ValueError("Certaines dates de début sont mal formées dans le fichier de comparaison.")
        if df1['end_date'].isnull().any():
            raise ValueError("Certaines dates de fin sont mal formées dans le fichier de comparaison.")
        if df2['start_date'].isnull().any():
            raise ValueError("Certaines dates de début sont mal formées dans le fichier de référence.")
        if df2['end_date'].isnull().any():
            raise ValueError("Certaines dates de fin sont mal formées dans le fichier de référence.")
        
        # Expansion des plages de dates pour les deux dataframes
        expanded_df1 = []
        for index, row in df1.iterrows():
            current_date = row['start_date']
            end_date = row['end_date']
            while current_date <= end_date:
                expanded_df1.append({'Matricule': row['matricule'], 'Date': current_date})
                current_date += timedelta(days=1)
        expanded_df1 = pd.DataFrame(expanded_df1)
        
        expanded_df2 = []
        for index, row in df2.iterrows():
            current_date = row['start_date']
            end_date = row['end_date']
            while current_date <= end_date:
                expanded_df2.append({'Matricule': row['matricule'], 'Date': current_date, 'Libellé': row['label']})
                current_date += timedelta(days=1)
        expanded_df2 = pd.DataFrame(expanded_df2)
        
        # Fusion des deux dataframes sur les colonnes Matricule et Date
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
    """
    Route principale pour le téléchargement et le traitement des fichiers.
    Gère l'upload des fichiers et lance les comparaisons de dates.
    """
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
            
            # Comparaison des dates pour les fichiers téléchargés
            compare_dates_multiple_sheets(filename1, filename2, output_file1)
            compare_dates(filename1, filename3, output_file2)
            
            flash('Les fichiers ont été comparés avec succès!', 'success')
            print("Comparaison réussie.")
            
            # Suppression des fichiers uploadés après traitement
            os.remove(filename1)
            os.remove(filename2)
            os.remove(filename3)
            
            return redirect(url_for('download_file', filename=f'output_{file2.filename}.xlsx'))

    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    """
    Route pour le téléchargement des fichiers générés.

    Args:
        filename (str): Le nom du fichier à télécharger.

    Returns:
        Response: Le fichier à télécharger ou une redirection avec un message d'erreur.
    """
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)
    except FileNotFoundError:
        flash('Fichier introuvable', 'danger')
        app.logger.error(f"Le fichier '{filename}' n'a pas été trouvé dans le dossier '{app.config['UPLOAD_FOLDER']}'")
        return redirect(url_for('upload_files'))
    except Exception as e:
        flash(f"Erreur lors du téléchargement du fichier: {e}", 'danger')
        app.logger.error(f"Erreur lors du téléchargement du fichier '{filename}': {e}")
        return redirect(url_for('upload_files'))

if __name__ == "__main__":
    app.run(debug=True)