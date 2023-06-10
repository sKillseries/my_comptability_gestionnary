# my_comptability_gestionnary

Mon outil de gestion comptable personnel.

## Installer les prérequis

Pour installer les prérequis pour l'utilisation de l'outil il faut au préalable tapez la commande ci-dessous:

```
pip install -U -r requirements.txt
```

## Guide d'utilisation

### Lancer le script

Pour exécuter le script deux façons de procéder:

#### Méthode 1

Pour lancer le script tapez:

```
python3 my_comptability_gestionnary.py
```

#### Méthode 2
 
Vous devrez d'abord rendre le script python exécutable en utilisant la commande ci-dessous:

```
chmod +x my_comptability_gestionnary.py
```

> __Info:__ Vous n'aurez à faire la commande ci-dessus qu'une fois.

Puis pour lancer le script tapez:

```
./my_comptability_gestionnary.py
```

### Fonctionnement

Une fois que vous avez lancer le script répondez aux questions posées.
Pour saisir les chiffres à virgule utiliser le point `(.)` et non la virgule `(,)`.

Une fois la saisie fini, celle-ci sera enregistré automatiquement dans le fichier excel `my_comptability_sheet.xlsx`, que vous pourrez consulter par la suite pour vérifier l'exactitude de votre saisie.

#### Le fichier excel

Avant de commencer votre saisie il faudra créer une feuille avec l'année en cours.
Le fichier est initialisé avec l'année 2023 car c'est l'année à laquelle a été développé l'outil.
Mais si vous utilisez l'outil en 2024 il faudra créer une feuille qui s'appelle 2024.
Il faudra dès lors copier le contenu de la feuille `template`.

> __Attention:__ Veuillez bien respecté la disposition de la feuille au niveau de l'emplacement des colonnes et des lignes, sinon l'enregistrement de la saisie ne sera pas bon.