# Application de Comparaison de Dates dans des Fichiers Excel

Cette application est une application web basée sur Flask permettant de comparer des dates présentes dans plusieurs fichiers Excel. L'application télécharge les fichiers, effectue des comparaisons de dates et génère un fichier de sortie avec les résultats. Cette application est utilisée afin de contrôler des dates de congés des employés et vérifier si elles concordent entre les différentes sources de données.

## Fonctionnalités

- **Téléchargement de fichiers Excel :** Permet de télécharger plusieurs fichiers Excel.
- **Comparaison de dates :** Compare les dates présentes dans différents fichiers Excel et génère un fichier de sortie avec les résultats.
- **Génération d'un fichier ZIP :** Crée un fichier ZIP contenant les fichiers Excel de sortie.

## Prérequis

- Python 3.x
- Flask
- Pandas
- XlsxWriter

## Documentation

Pour accéder à plus de documentation, cliquez sur le lien suivant: [https://krdto.github.io/outil_rapprochement_cong-s/](https://krdto.github.io/outil_rapprochement_cong-s/)

## Installation

1. Clonez le repository :
    ```bash
    git clone https://github.com/Krdto/outil_rapprochement_cong-s.git
    ```

2. Accédez au répertoire du projet :
    ```bash
    cd outil_rapprochement_cong-s
    ```

3. Installez les packages Python requis :
    ```bash
    pip install -r requirements.txt
    ```

## Utilisation

1. Exécutez l'application :
    ```bash
    python app.py
    ```

2. Accédez à l'interface web :
    Ouvrez un navigateur web et allez à [http://localhost:5000](http://localhost:5000).

3. Déposez les fichiers Excel :
    - Sélectionnez les fichiers Excel nécessaires.
    - Cliquez sur le bouton **"Comparer les Fichiers"**.

4. Téléchargez le fichier ZIP de résultats :
    Le fichier ZIP contenant les fichiers Excel générés sera disponible en téléchargement.

## Structure des fichiers

- `app.py` : Script principal de l'application.
- `templates/index.html` : Modèle HTML pour l'interface web.
- `static/` : Fichiers statiques (par exemple, images, feuilles de style).
- `uploads/` : Répertoire pour stocker les fichiers téléchargés et les résultats.
