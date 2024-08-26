from openpyxl import Workbook
classeur=Workbook()
morts=classeur.create_sheet("LISTE_PERSONNES_DECEDES")
entete_tableau=["CODE","NOMS","DATE DE NAISSANCE","DATE DE DECES","ETAT CIVIL","DOMICILE","CAUSE DE DECES","CIMETIERE","LIEU DE DECES"]
morts.append(entete_tableau)
classeur.save("DECES.xlsx")
