# importer le module openpy
from openpyxl import Workbook
classeur=Workbook()
motards=classeur.create_sheet("ENREGISTREMENT_MOTARDS")
titre=["NOM","POSTNOM","PRENOM","CONTACT","PLAQUE MOTO","CODE D'ENREGISTREMENTS","PARKING","ASSOCIATION"]
motards.append(titre)
classeur.save("MOTARDS.xlsx")
