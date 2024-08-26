# importer le module openpy
from openpyxl import load_workbook
# ouvrir le fichier excel deja enregistré
classeur=load_workbook("ETUDIANTS.xlsx")
# ouvrir la feuille des etudiants
etudiants=classeur["LISTE_ETUDIANTS"]
nom=input("Entrez le nom de l'etudiant: ")
postnom=input("Entrez le postnom de l'etudiant: ")
prenom=input("Entrez le prenom de l'etudiant: ")
date_naissance=input("Entrez la date de naissance de l'etudiant': ")
faculte=input("Entrez la faculté de l'etudiant': ")
tuteur=input("Entrez le nom du tuteur de l'etudiant': ")
contact=input("Entrez le contact de l'etudiant': ")
matricule=input("Entrez le matricule de l'etudiant': ")
profil=[nom,postnom,prenom,date_naissance,faculte,tuteur,contact,matricule]
etudiants.append(profil)
# enregistrer les donnees entrés 
classeur.save('ETUDIANTS.xlsx')                                                          
