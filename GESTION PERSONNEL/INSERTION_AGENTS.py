# importer le module openpy
from openpyxl import load_workbook
# ouvrir le fichier excel deja enregistré
classeur=load_workbook("AGENTS.xlsx")
# ouvrir la feuille des personnes engages
agents=classeur["LISTE_AGENTS"]
nom=input("Entrez le nom de l'agent: ")
postnom=input("Entrez le postnom de l'agent: ")
prenom=input("Entrez le prenom de l'agent: ")
contrat=input("Entrez le type de contrat: ")
date_embauche=input("Quel est la date d'embauche de cet agent ?': ")
adresse=input("Quel est l'adresse de cet agent ?: ")
poste=input("Entrez le poste de l'agent ': ")
contact=input("Entrez le contact de l'agent': ")
profil=[nom,postnom,prenom,contrat,date_embauche,adresse,poste,contact]
agents.append(profil)
# enregistrer les donnees entrés 
classeur.save('AGENTS.xlsx')                                                          
