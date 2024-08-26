from openpyxl import Workbook
classeur=Workbook()
mariages=classeur.create_sheet("LISTE_MARIAGES")
entete_tableau=["CODE","NOMS DU MARI","NOMS DE LA MARIEE","DATE DE MARIAGE","REGIME MATRIMONIAL","NOMS DU PARRAIN","NOMS DE LA MARRAINE","COMMUNE","RELIGION","NOMBRE D'ENFANTS","LIEU DU MARIAGE"]
mariages.append(entete_tableau)
classeur.save("MARIAGES.xlsx")
