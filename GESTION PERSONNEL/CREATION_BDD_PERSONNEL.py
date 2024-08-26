# importer le module openpy
from openpyxl import Workbook
# creation du classeur
classeur=Workbook()
# creation de la feuille des agents
agents=classeur.create_sheet("LISTE_AGENTS")
# creation de l'entete comme premiere ligne de la feuille des agents
entete=["NOM","POSTNOM","PRENOM","TYPE DE CONTRAT","DATE EMBAUCHE","ADRESSE","POSTE","CONTACT"]
agents.append(entete)
# enregistrement du fichier 
classeur.save("AGENTS.xlsx")
