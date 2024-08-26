from openpyxl import Workbook
classeur=Workbook()
naissances=classeur.create_sheet("LISTE_NAISSANCES")
entete_tableau=["CODE","NOMS","DATE DE NAISSANCE","NOM DU PERE","NOM DE LA MERE","DOMICILE","HOPITAL","ETAT","LIEU DE NAISSANCE"]
naissances.append(entete_tableau)
classeur.save("NAISSANCES.xlsx")
