#!/usr/bin/python3

from openpyxl import Workbook
from openpyxl import load_workbook
from termcolor import colored

def depensefixe():
    global depenses
    telephone = input("Combien coûte votre abonnement téléphonique ce mois-ci ?: ")
    internet = input("Combien coûte votre abonnement internet ce mois-ci ?: ")
    loyer = input("Combien coûte votre loyer ce mois-ci ?: ")
    transport = input("Combien coûte votre abonnement de transport ce mois-ci ?: ")
    auto = input("Combien coûte votre traite auto ce mois-ci ?: ")
    electricite = input("Quelle est la somme de votre facture d'électricité ce mois-ci ?: ")
    eau = input("Quelle est la somme de votre facture d'eau ce mois-ci ?:")
    depenses = float(telephone) + float(internet) + float(loyer) + float(transport) + float(auto) + float(electricite) + float(eau)
    depensesfixes = round(depenses, 2)
    print(colored(f"Vos dépenses fixes s'élèvent à {depensesfixes}€ ce mois-ci. \n", 'yellow'))
    menu()

def investissement():
    global investissements
    investissement = input("Combien avez-vous investi ce mois-ci ?: ")
    investissements = float(investissement)
    print(colored(f"Vous avez investi {investissements}€ ce mois-ci. \n", 'green'))
    depensefixe()

def epargne():
    global epargnes
    epargne = input("Combien avez-vous épargné ce mois-ci ?: ")
    epargnes = float(epargne)
    print(colored(f"Vous avez épargné {epargnes}€ ce mois-ci. \n", 'green'))
    investissement()

def revenu():
    global revenus
    salaire = input("Quelle est la somme de votre salaire ce mois-ci ?: ")
    prime = input("Quelle a été la somme de votre prime ce mois-ci ?: ")
    locations = input("Quelle est la somme de vos revenus de location ce mois-ci ?: ")
    dividendes = input("Quelle est la somme de vos dividendes ce mois-ci ?: ")
    interets = input("Quelle est la somme de vos intérêts ce mois-ci ?: ")
    redevances = input("Quelle est la somme de vos redevances ce mois-ci ?: ")
    revenus = float(salaire) + float(locations) + float(dividendes) + float(interets) + float(redevances)
    revenues = round(revenus, 2)
    print(colored(f"Vos revenus ce mois-ci s'élève à {revenues}€. \n", 'green'))
    epargne()

def saveandproceed():
    print(colored("La saisie a été enregistré dans le fichier my_comptability_sheet.xlsx.", 'green'))
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
        sheet.cell(row=5, column=3).value = revenues
        sheet.cell(row=5, column=4).value = epargnes
        sheet.cell(row=5, column=5).value = investissements
        sheet.cell(row=5, column=6).value = depensesfixes
        wb.save("my_comptability_sheet.xlsx")
        saveandproceed()
    elif choix == "2":
        sheet.cell(row=6, column=3).value = revenues
        sheet.cell(row=6, column=4).value = epargnes
        sheet.cell(row=6, column=5).value = investissements
        sheet.cell(row=6, column=6).value = depensesfixes
        wb.save("my_comptability_sheet.xlsx")
        saveandproceed()
    elif choix == "3":
        sheet.cell(row=7, column=3).value = revenues
        sheet.cell(row=7, column=4).value = epargnes
        sheet.cell(row=7, column=5).value = investissements
        sheet.cell(row=7, column=6).value = depensesfixes
        wb.save("my_comptability_sheet.xlsx")
        saveandproceed()
    elif choix == "4":
        sheet.cell(row=8, column=3).value = revenues
        sheet.cell(row=8, column=4).value = epargnes
        sheet.cell(row=8, column=5).value = investissements
        sheet.cell(row=8, column=6).value = depensesfixes
        wb.save("my_comptability_sheet.xlsx")
        saveandproceed()
    elif choix == "5":
        sheet.cell(row=9, column=3).value = revenues
        sheet.cell(row=9, column=4).value = epargnes
        sheet.cell(row=9, column=5).value = investissements
        sheet.cell(row=9, column=6).value = depensesfixes
        wb.save("my_comptability_sheet.xlsx")
        saveandproceed()
    elif choix == "6":
        sheet.cell(row=10, column=3).value = revenues
        sheet.cell(row=10, column=4).value = epargnes
        sheet.cell(row=10, column=5).value = investissements
        sheet.cell(row=10, column=6).value = depensesfixes
        wb.save("my_comptability_sheet.xlsx")
        saveandproceed()
    elif choix == "7":
        sheet.cell(row=11, column=3).value = revenues
        sheet.cell(row=11, column=4).value = epargnes
        sheet.cell(row=11, column=5).value = investissements
        sheet.cell(row=11, column=6).value = depensesfixes
        wb.save("my_comptability_sheet.xlsx")
        saveandproceed()
    elif choix == "8":
        sheet.cell(row=12, column=3).value = revenues
        sheet.cell(row=12, column=4).value = epargnes
        sheet.cell(row=12, column=5).value = investissements
        sheet.cell(row=12, column=6).value = depensesfixes
        wb.save("my_comptability_sheet.xlsx")
        saveandproceed()
    elif choix == "9":
        sheet.cell(row=13, column=3).value = revenues
        sheet.cell(row=13, column=4).value = epargnes
        sheet.cell(row=13, column=5).value = investissements
        sheet.cell(row=13, column=6).value = depensesfixes
        wb.save("my_comptability_sheet.xlsx")
        saveandproceed()
    elif choix == "10":
        sheet.cell(row=14, column=3).value = revenues
        sheet.cell(row=14, column=4).value = epargnes
        sheet.cell(row=14, column=5).value = investissements
        sheet.cell(row=14, column=6).value = depensesfixes
        wb.save("my_comptability_sheet.xlsx")
        saveandproceed()
    elif choix == "11":
        sheet.cell(row=15, column=3).value = revenues
        sheet.cell(row=15, column=4).value = epargnes
        sheet.cell(row=15, column=5).value = investissements
        sheet.cell(row=15, column=6).value = depensesfixes
        wb.save("my_comptability_sheet.xlsx")
        saveandproceed()
    elif choix == "12":
        sheet.cell(row=16, column=3).value = revenues
        sheet.cell(row=16, column=4).value = epargnes
        sheet.cell(row=16, column=5).value = investissements
        sheet.cell(row=16, column=6).value = depensesfixes
        wb.save("my_comptability_sheet.xlsx")
        saveandproceed()
    elif choix == "q":
        print("Fermeture du programme.")
        exit(0)
    elif choix == "Q":
        print("Fermeture du programme.")
        exit(0)
    else:
        print(colored("Veuillez sélectionnez une option présente dans la liste !", 'yellow'))
        menu()

def main():
    global wb
    global sheet
    annee = input("En quelle année sommes-nous ? ")
    wb = load_workbook("my_comptability_sheet.xlsx")
    sheet = wb[f'{annee}']
    revenu()

if __name__ == '__main__':
    main()