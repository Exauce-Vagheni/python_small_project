# importer le module openpy
from openpyxl import Workbook
# creation du classeur
classeur=Workbook()
# creation de la feuille des etudiants
etudiants=classeur.create_sheet("LISTE_ETUDIANTS")
# creation de l'entete comme premiere ligne de la feuille des etudiants
entete=["NOM","POSTNOM","PRENOM","DATE NAISSANCE","FACULTE","TUTEUR","CONTACT","MATRICULE"]
etudiants.append(entete)
# enregistrement du fichier 
classeur.save("ETUDIANTS.xlsx")
