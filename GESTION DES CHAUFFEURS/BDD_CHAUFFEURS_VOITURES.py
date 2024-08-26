# importer le module openpy
from openpyxl import Workbook
classeur=Workbook()
chauffeurs_voitures=classeur.create_sheet("ENREGISTREMENT_CHAUFFEURS")
titre=["NOM","POSTNOM","PRENOM","CONTACT","PLAQUE VOITURE","CODE D'ENREGISTREMENT","PARKING","ASSOCIATION"]
chauffeurs_voitures.append(titre)
classeur.save("CHAUFFEURS_VOITURES.xlsx")
