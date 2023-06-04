#!/usr/bin/python3

from openpyxl import Workbook
from openpyxl import load_workbook

def saveandproceed():
    print("La saisie a été enregistré dans le fichier my_comptability_sheet.xlsx.")
    answer = input("Voulez-vous saisir un autre mois (o/N) ?: ")
    if answer == "O":
        menu()
    elif answer == "N":
        print("Fermeture du programme.")
        exit(0)
    elif answer == "o":
        menu()
    elif answer == "n":
        print("Fermeture du programme.")
        exit(0)
    elif answer == "Y":
        menu()
    elif answer == "y":
        menu()
    else:
        print("Fermeture du programme.")
        exit(0)

def menu():
    print(
    """
    Quel mois voulez-vous enregistrer la saisie ?
    > 1 : Janvier
    > 2 : Février
    > 3 : Mars
    > 4 : Avril
    > 5 : Mai
    > 6 : Juin
    > 7 : Juillet
    > 8 : Août
    > 9 : Septembre
    > 10: Octobre
    > 11: Novembre
    > 12: Décembre
    > q : Quitter
    """
    )
    choix = input("Veuillez sélectionnez le numéro du mois: ")
    if choix == "1":
        sheet.cell(row=5, column=3).value = revenus
        wb.save("my_comptability_sheet.xlsx")
        save()
    elif choix == "2":
        sheet.cell(row=6, column=3).value = revenus
        wb.save("my_comptability_sheet.xlsx")
        save()
    elif choix == "3":
        sheet.cell(row=7, column=3).value = revenus
        wb.save("my_comptability_sheet.xlsx")
        save()
    elif choix == "4":
        sheet.cell(row=8, column=3).value = revenus
        wb.save("my_comptability_sheet.xlsx")
        save()
    elif choix == "5":
        sheet.cell(row=9, column=3).value = revenus
        wb.save("my_comptability_sheet.xlsx")
        save()
    elif choix == "6":
        sheet.cell(row=10, column=3).value = revenus
        wb.save("my_comptability_sheet.xlsx")
        save()
    elif choix == "7":
        sheet.cell(row=11, column=3).value = revenus
        wb.save("my_comptability_sheet.xlsx")
        save()
    elif choix == "8":
        sheet.cell(row=12, column=3).value = revenus
        wb.save("my_comptability_sheet.xlsx")
        save()
    elif choix == "9":
        sheet.cell(row=13, column=3).value = revenus
        wb.save("my_comptability_sheet.xlsx")
        save()
    elif choix == "10":
        sheet.cell(row=14, column=3).value = revenus
        wb.save("my_comptability_sheet.xlsx")
        save()
    elif choix == "11":
        sheet.cell(row=15, column=3).value = revenus
        wb.save("my_comptability_sheet.xlsx")
        save()
    elif choix == "12":
        sheet.cell(row=16, column=3).value = revenus
        wb.save("my_comptability_sheet.xlsx")
        save()
    elif choix == "q":
        print("Fermeture du programme.")
        exit(0)
    elif choix == "Q":
        print("Fermeture du programme.")
        exit(0)
    else:
        print("Veuillez sélectionnez une option présente dans la liste")
        menu()

def main():
    global wb
    global sheet
    global salaire
    global locations
    global dividendes
    global interets
    global redevances
    global revenus
    annee = input("En quelle année sommes-nous ? ")
    wb = load_workbook("my_comptability_sheet.xlsx")
    sheet = wb[f'{annee}']
    salaire = input("Quelle est la somme de votre salaire ce mois-ci ?: ")
    locations = input("Quelle est la somme de vos revenus de location ce mois-ci ?: ")
    dividendes = input("Quelle est la somme de vos dividendes ce mois-ci ?: ")
    interets = input("Quelle est la somme de vos intérêts ce mois-ci ?: ")
    redevances = input("Quelle est la somme de vos redevances ce mois-ci ?: ")
    revenus = float(salaire) + float(locations) + float(dividendes) + float(interets) + float(redevances)
    menu()

if __name__ == '__main__':
    main()