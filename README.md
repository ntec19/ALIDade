# Projet 'ALIDade'

_Automatisation pour les Livrets Individuels Dématérialisés, v2023-02-07f_

![image alidade](https://marine-data.co.uk/wp-content/uploads/2016/03/MD69BC-800x600.1-300x225.png)

----

## Objectifs généraux

L'objectif de ces scripts Python est d'automatiser autant que possible le traitement des livrets dématérialisés pour les diplômes suivants :
- [31212] baccalaureat professionnel "Métiers de l'accueil"
- [31213] baccalaureat professionnel "Métiers du commerce et de la vente - Option A : Animation et gestion de l'espace commercial"
- [31214] baccalaureat professionnel "Métiers du commerce et de la vente - Option B : Prospection clientèle et valorisation de l'offre commerciale"
- [31224] CAP "Équipier polyvalent du commerce"

----

## Prérequis

Pour que ce script fonctionne correctement, il est nécessaire de vérifier les prérequis suivants :
- Python 3 installé ;
- bibliothèque Python ['openpyxl'](https://pypi.org/project/openpyxl/) installée ;
- dans le même répertoire que le script, se trouvent :
    - 👉 les fichiers d'export Cyclade, qui possèdent l'extension ".csv"  
      et dont le nom commence par "cyclade".  
      Il peut y en avoir plusieurs (typiquement, un pour le CAP,  
      un pour le bac pro) ; les candidats seront recherchés  
      dans chacun d'entre eux ;
    - 👉 les fichiers modèles nécessaires :  
         - fichiers modèles candidats : 31212.xlsx, 31213.xlsx, 31214.xlsx, 31224.xlsx
         - fichiers modèles synthèse établissement : 31212_etab.xlsx, 31213_etab.xlsx, 31214_etab.xlsx, 31224_etab.xlsx

Le fichier 'setup.py' est particulièrement important, car il contient les CONSTANTES qu'il conviendra de modifier pour changer les valeurs par défaut du programme.
⚠ En particulier, le dictionnaire CORRESPONDANCE_CANDIDATS_SYNTHESE assure la correspondance des références de cellules entre les fichiers individuels de candidat et les fichiers de synthese établissement : il devra être défini avec attention. ⚠

----

## Temps 1 : génération des livrets

Le script **`genere_grilles_indiv.py`** permet de générer les livrets dématérialisés individuels. Ce sont des fichiers Excel créés par copie d'un modèle dans una arborescence cohérente, puis modifiés pour y insérer les informations personnelles des candidats.

----

## Temps 2 : consolidation des notes

**à faire :**

Le script **`consolidation.py`** permet de parcourir tous les livrets individuels des candidats présents dans un dossier, de récupérer les notes obtenues et de consolider l'ensemble dans un document unique pourl'établissement.

----

fin