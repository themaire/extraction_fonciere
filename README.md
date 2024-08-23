Pré-requis :
(Methode d'installation sous Windows Powershell.)

1 Installer Python (j'utilise la version 3.12.4).

2 Ouvrir PowerShell en tant qu'administrateur:

Appuyez sur la touche Windows et tapez "PowerShell".
(Faites un clic droit sur "Windows PowerShell" et sélectionnez "Exécuter en tant qu'administrateur".)
Exécuter la commande suivante et dire oui (O) :
Set-ExecutionPolicy RemoteSigned

2 Créer un environnement virtuel. Il s'agit de preferer installer des dépendances logiciel appellés "modules Pytthon" dans ce qu'on appel un venv. Pour faire simple, cela conciste à ne pas "salir" son Python. Les modules seront installé que dans le dossier du projet et non de manière générale accessible depuis n'importe où dans l'ordinateur.
python -m venv env_foncier

2 Rentrer/activer l'environnement virtuel "env_foncier" (bien etre placé dans le dossier de ce projet):
env_foncier\Scripts\activate

3 Mettre à jour pip (gestionnaire d'installation de modules Python automatisé) :
python -m pip install --upgrade pip

4 Installer les dépendences :
python -m pip install -r requirements.txt
(Ceci va installer les modules pour lire et écrire des fichiers Excel, exploiter une base de données Postgresql et utiliser les variables d'environnement : la prochaine étape.)

5 Sortez de l'environnement virtuel, nous en avons plus besoin avec cette commande :
deactivate

6 Dupliquer le fichier .env_exemple en .env avec cette commande :
cp .env_exemple .env

6 Editez ce fichier .env pour correspondre a votre configuration personnelle.

7 Executer le script :
.\env_foncier\Scripts\python .\extraction_foncier.py

