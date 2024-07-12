from flask import Flask, request, redirect, url_for, send_from_directory, render_template, flash
import os
import pandas as pd
from datetime import timedelta
from zipfile import ZipFile

app = Flask(__name__, template_folder="../templates", static_folder="../static")
app.secret_key = "infokey"
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def allowed_file(filename):
    """
    Vérifie l'extension du fichier passé en paramètre.

    Args:
        filename (str): Le nom du fichier à vérifier.

    Returns:
        bool: True si l'extension est autorisée, False sinon.
    """
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def compare_dates_multiple_sheets(reference_file, comparison_file, output_file):
    """
    Compare des dates à travers toutes les feuilles d'un fichier Excel avec le fichier de référence 
    et écris le résultat de la comparaison dans un fichier d'output.

    Args:
        reference_file (str): Chemin (path) vers le fichier Excel de référence.
        comparison_file (str): Chemin (path) vers le fichier Excel à comparer.
        output_file (str): Chemin (path) vers le fichier Excel d'output pour enregistrer les résultats.

    Raises:
        ValueError: Si les dates sont mal formatées ou s'il manque des colonnes.
        Exception: Pour toute exception.
    """
    try:
        # Lis toutes les feuilles du fichier de comparaison
        df1 = pd.read_excel(comparison_file, sheet_name=None)
        # Lis le fichier de référence
        df2 = pd.read_excel(reference_file)
        
        # Normaliser et renomme le noms des colonnes du fichier de référence
        df2.columns = [col.strip().lower() for col in df2.columns]
        df2 = df2.rename(columns={
            'début': 'start_date',
            'fin': 'end_date',
            'matricule': 'matricule',
            'libellé': 'label'
        })
        
        # Convertit les colonnes de dates de début et de fin en format datetime
        df2['start_date'] = pd.to_datetime(df2['start_date'], format='%Y-%m-%d', errors='coerce').dt.date
        df2['end_date'] = pd.to_datetime(df2['end_date'], format='%Y-%m-%d', errors='coerce').dt.date
        
        # Vérifie que les dates sont bien formatées
        if df2['start_date'].isnull().any():
            raise ValueError("Certaines dates de début sont mal formées dans le fichier de référence.")
        if df2['end_date'].isnull().any():
            raise ValueError("Certaines dates de fin sont mal formées dans le fichier de référence.")

        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

        for sheet_name, df1_sheet in df1.items():
            # Normaliser et renomme le noms des colonnes du fichier de comparaison
            df1_sheet.columns = [col.strip().lower() for col in df1_sheet.columns]
            if 'date à contrôler' in df1_sheet.columns:
                df1_sheet = df1_sheet.rename(columns={'date à contrôler': 'control_date', 'matricule': 'matricule'})
            else:
                raise ValueError(f"La colonne 'Date à contrôler' est manquante dans la feuille {sheet_name} du fichier de comparaison.")
            
            # Convertit les colonnes de dates à contrôler en format datetime
            df1_sheet['control_date'] = pd.to_datetime(df1_sheet['control_date'], format='%d/%m/%Y', errors='coerce').dt.date
            if df1_sheet['control_date'].isnull().any():
                raise ValueError(f"Certaines dates à contrôler sont mal formées dans la feuille {sheet_name} du fichier de comparaison.")
            
            results_df = pd.DataFrame(columns=['Matricule', 'Date à contrôler', 'Libellé'])
            
            # Compare les dates pour chaque matricule et date à contrôler
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
            
            # Ecris les résultats dans le fichier Excel d'output
            results_df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.set_column('B:B', 18)
        
        writer.close()
    except Exception as e:
        print(f"Une erreur s'est produite : {e}")

def compare_dates(reference_file, comparison_file, output_file):
    """
    Compare les dates entre deux fichiers Excel et écris les résultats dans un fichier d'output.

    Args:
        reference_file (str): Chemin (path) vers le fichier Excel de référence.
        comparison_file (str): Chemin (path) vers le fichier Excel de comparaison.
        output_file (str): Chemin (path) vers le fichier Excel d'output pour enregistrer les résultats.

    Raises:
        ValueError: Si les dates sont mal formatées ou s'il manque des colonnes.
        Exception: Pour toute exception.
    """
    try:
        # Lecture du fichier de comparaison et du fichier de référence
        df1 = pd.read_excel(comparison_file)
        df2 = pd.read_excel(reference_file)
        
        # Normalise et renomme le noms des colonnes des deux fichiers
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
        
        required_columns_df1 = {'matricule', 'start_date', 'end_date'}
        required_columns_df2 = {'matricule', 'start_date', 'end_date', 'label'}
        
        # Vérifie si les colonnes nécessaires sont presentes
        if not required_columns_df1.issubset(df1.columns) or not required_columns_df2.issubset(df2.columns):
            raise ValueError("Les fichiers ne contiennent pas les colonnes recherchées.")
        
        # Convertit les colonnes de dates en format datetime
        df1['start_date'] = pd.to_datetime(df1['start_date'], format='%d/%m/%Y', errors='coerce').dt.date
        df1['end_date'] = pd.to_datetime(df1['end_date'], format='%d/%m/%Y', errors='coerce').dt.date
        df2['start_date'] = pd.to_datetime(df2['start_date'], format='%Y-%m-%d', errors='coerce').dt.date
        df2['end_date'] = pd.to_datetime(df2['end_date'], format='%Y-%m-%d', errors='coerce').dt.date
        
        # Vérifie si les dates sont bien formatées
        if df1['start_date'].isnull().any():
            raise ValueError("Certaines dates de début sont mal formées dans le fichier de comparaison.")
        if df1['end_date'].isnull().any():
            raise ValueError("Certaines dates de fin sont mal formées dans le fichier de comparaison.")
        if df2['start_date'].isnull().any():
            raise ValueError("Certaines dates de début sont mal formées dans le fichier de référence.")
        if df2['end_date'].isnull().any():
            raise ValueError("Certaines dates de fin sont mal formées dans le fichier de référence.")
        
        # Transforme les intervalles de dates en une succession de dates uniques de la date de début à la date de fin
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
        
        # Merge les dataframes selon les colonnes 'Matricule' et 'Date'
        merged_df = pd.merge(expanded_df1, expanded_df2, on=['Matricule', 'Date'], how='left')
        merged_df['Libellé'] = merged_df['Libellé'].fillna('La date ne correspond pas')
        
        # Enregistre la dataframe dans un fichier Excel d'output
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
    Route pour gérer le dépôt de fichiers et lancer la comparaison.

    Returns:
        redirect: Redirige vers le téléchargement si l'opération est réussie, sinon recharge la page.
    """
    if request.method == 'POST':
        # Vérifie si des fichiers font partie de la requête
        if 'file1' not in request.files or 'file2' not in request.files or 'file3' not in request.files:
            flash('Aucun fichier trouvé', 'danger')
            return redirect(request.url)
        
        file1 = request.files['file1']
        file2 = request.files['file2']
        file3 = request.files['file3']
        
        # Vérifie si un fichier n'est pas sélectionné
        if file1.filename == '' or file2.filename == '' or file3.filename == '':
            flash('Aucun fichier sélectionné', 'danger')
            return redirect(request.url)
        
        # Vérifie si l'extension du fichier est correcte
        if file1 and allowed_file(file1.filename) and file2 and allowed_file(file2.filename) and file3 and allowed_file(file3.filename):
            filename1 = os.path.join(app.config['UPLOAD_FOLDER'], file1.filename)
            filename2 = os.path.join(app.config['UPLOAD_FOLDER'], file2.filename)
            filename3 = os.path.join(app.config['UPLOAD_FOLDER'], file3.filename)
            file1.save(filename1)
            file2.save(filename2)
            file3.save(filename3)
            
            output_file1 = os.path.join(app.config['UPLOAD_FOLDER'], f'output_{file2.filename}.xlsx')
            output_file2 = os.path.join(app.config['UPLOAD_FOLDER'], f'output_{file3.filename}.xlsx')
            
            # Comparaison des dates
            compare_dates_multiple_sheets(filename1, filename2, output_file1)
            compare_dates(filename1, filename3, output_file2)
            
            flash('Les fichiers ont été comparés avec succès!', 'success')
            
            # Creation d'un fichier zip contenant les deux fichiers résultants
            zip_filename = os.path.join(app.config['UPLOAD_FOLDER'], 'output_files.zip')
            with ZipFile(zip_filename, 'w') as zipf:
                zipf.write(output_file1, os.path.basename(output_file1))
                zipf.write(output_file2, os.path.basename(output_file2))
            
            # Supprime les fichiers temporaires
            os.remove(filename1)
            os.remove(filename2)
            os.remove(filename3)
            os.remove(output_file1)
            os.remove(output_file2)
            
            return redirect(url_for('download_file', filename='output_files.zip'))

    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    """
    Route pour gérer le téléchargment des fichiers.

    Args:
        filename (str): Le nom du fichier à télécharger.

    Returns:
        send_from_directory: Renvoie le fichier du dossier uploads en pièce jointe.
        redirect: Recharge la page si le fichier n'est pas trouvé ou pour tout autre erreur.
    """
    try:
        return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)
    except FileNotFoundError:
        flash('Fichier introuvable', 'danger')
        return redirect(url_for('upload_files'))
    except Exception as e:
        flash(f"Erreur lors du téléchargement du fichier: {e}", 'danger')
        return redirect(url_for('upload_files'))

if __name__ == "__main__":
    app.run(debug=True)
