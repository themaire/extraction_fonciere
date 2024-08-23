# Script Python d'extraction foncier

Ceci est un script utilitaire qui me sert dans le cadre de mon travail. Il prends en entrée un fichier Excel completé par des collegues dans le but de faire une demande d'extraction de données. Le script va donc lire ce fichier Excel et faire des requetes à une base de données pour produire un fichier .shape (couche spatiale de polygone) et un fichier Excel de sortie correspondant à la demande en entrée.

Un environement virtuel Python est mise en oeuvre pour l'installation des 3 modules utilisés.

Les réglages se trouve dans le fichier .env à obtenir depuis le fichier d'exemple où le chemin du dossier de travail (là où les fichiers d'entree et de sortie seront lus et écrits), les informations de connexion à la base de données et le chemin de l'outil en ligne de commande pgsql2shp.


### Pré-requis :
(Methode d'installation sous Windows Powershell.)

## 1 Installer Python (j'utilise la version 3.12.4).

## 2 Ouvrir PowerShell en tant qu'administrateur:

Appuyez sur la touche Windows et tapez "PowerShell".
(Faites un clic droit sur "Windows PowerShell" et sélectionnez "Exécuter en tant qu'administrateur".)
Exécuter la commande suivante et dire oui (O) :
Set-ExecutionPolicy RemoteSigned

## 2 Créer un environnement virtuel. 
Il s'agit de preferer installer des dépendances logiciel appellés "modules Pytthon" dans ce qu'on appel un venv. Pour faire simple, cela conciste à ne pas "salir" son Python. Les modules seront installé que dans le dossier du projet et non de manière générale accessible depuis n'importe où dans l'ordinateur.
python -m venv env_foncier

## 3 Rentrer/activer l'environnement virtuel "env_foncier" (bien etre placé dans le dossier de ce projet):
```powershell
env_foncier\Scripts\activate
```

## 4 Mettre à jour pip (gestionnaire d'installation de modules Python automatisé) :
```powershell
python -m pip install --upgrade pip
```

## 5 Installer les dépendences :
python -m pip install -r requirements.txt
(Ceci va installer les modules pour lire et écrire des fichiers Excel, exploiter une base de données Postgresql et utiliser les variables d'environnement : la prochaine étape.)

## 6 Sortez de l'environnement virtuel, nous en avons plus besoin avec cette commande :
```powershell
deactivate
```

## 7 Dupliquer le fichier .env_exemple en .env avec cette commande :
```powershell
cp .env_exemple .env
```

## 8 Editez ce fichier .env pour correspondre a votre configuration personnelle.

## 9 Executer le script, l'idée est la suivante :

Usage :

"Emplacement du lanceur Python de l'énvironnement virtuel" "Emplacement du script Python" "mettre_0_ou_1_pour_ecrire_dans_la_table_d'historique"


Par exemple, on peut lancer le script de cette manière uniquement en utilisant des chemins absolus :

```powershell
& 'C:\Users\toto\OneDrive - xxxXXXxxxXXXxx\chemin\du\script\env_foncier\Scripts\python.exe' 'C:\Users\toto\OneDrive - xxxXXXxxxXXXxx\chemin\du\script\extraction_foncier.py' 'C:\Users\nelie\OneDrive - xxxXXXxxxXXXxx\chemin\où\les\collegues\mettent\leurs\fichiers\de\demanandes\a_traiter\MBV_FR2100283_51_Marais de St-Gond_Modele_demande_extraction_ffna.xlsx' 0
```