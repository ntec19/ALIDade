# Projet 'ALIDade'

_Automatisation pour les Livrets Individuels Dématérialisés, v20230319_

![image alidade](https://marine-data.co.uk/wp-content/uploads/2016/03/MD69BC-800x600.1-300x225.png)

----

## Objectifs généraux

L'objectif de ces scripts Python est d'automatiser autant que possible le traitement des livrets dématérialisés pour les diplômes suivants :
- [31212] baccalauréat professionnel "Métiers de l'accueil"
- [31213] baccalauréat professionnel "Métiers du commerce et de la vente - Option A : Animation et gestion de l'espace commercial"
- [31214] baccalauréat professionnel "Métiers du commerce et de la vente - Option B : Prospection clientèle et valorisation de l'offre commerciale"
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

Le script **`consolidation.py`** permet de parcourir tous les livrets individuels des candidats présents dans un dossier, de récupérer les notes obtenues et de consolider l'ensemble dans un document unique pourl'établissement.

----

## Documentation

Une documentation _orientée utilisateur_ est fournie au format Word à la racine du projet.

----

## Notes, règles et questions...

### Les codes diplômes :

- '31212': "bacpro_MA"
- '31213': "bacpro_MVC_A_AGEC"
- '31214': "bacpro_MVC_B_PC"
- '31224': "CAP_EPC"

### Les noms des fichiers MODELES

```
    Livret individuel dématérialisé
=   "fichier Excel candidat"
=   31224.xlsx  (par exemple, pour le CAP EPC)
```

```
    Synthèse établissement
=   "fichier Excel établissement"
=   31224_etab.xlsx  (par exemple, pour le CAP EPC)
```

donc 8 fichiers à préparer, ainsi nommés.

### Contraintes / livret individuel dématérialisé (fichiers `xxxxx.xlsx`)

- Les modification (cellules à modifier avec les valeurs issues de Cyclade)
doivent être réalisées sur **une seule feuille**, nommée de la même
manière pour **tous les diplômes**. Généralement : la première feuille du classeur.

- Les reports d'informations (identité du candidat, etc.) sont opérés par formules Excel entre feuilles du même classeur.

### Contraintes / fichier de synthèse établissement (fichiers `xxxxx_etab.xlsx`)

- il doit y avoir une relation 'injective' entre les infos lues dans les livrets des candidats et les cellules d'un candidat sur le fichier de synthèse (contre-exemple : 'nom', 'prénom' -> 'nom prénom').

- le dictionnaire `CORRESPONDANCE_CANDIDATS_SYNTHESE` est un élément-clé :

```
{   # source :          fichier individuel du candidat
    # destination :     fichier de synthese établissement
    
    '31224': {
       #'champ' : [ [feuille_source, cellule_source], [feuille_destination, première_ligne_des_données, colonne ] ]
        'nom':      [ ['1-Candidat, établissement', 'E26'], ['RECAPNOTES', 12, 2] ],
        'prenom':   [ ['1-Candidat, établissement', 'E28'], ['RECAPNOTES', 12, 3] ],
        'n_cand':   [ ['1-Candidat, établissement', 'E30'], ['RECAPNOTES', 12, 1] ],
        'noteEP1':  [ ['5- Synthèse', 'Q12'], ['RECAPNOTES', 12, 4] ],
        'noteEP2':  [ ['5- Synthèse', 'Q16'], ['RECAPNOTES', 12, 5] ],
        'noteEP3':  [ ['5- Synthèse', 'Q20'], ['RECAPNOTES', 12, 6] ],
        'pfmp1':  [ ['5- Récapitulatif PFMP', 'B13'], ['PFMP', 10, 5] ],
        'pfmp2':  [ ['5- Récapitulatif PFMP', 'B14'], ['PFMP', 10, 6] ],
        'pfmp3':  [ ['5- Récapitulatif PFMP', 'B15'], ['PFMP', 10, 7] ],
        'pfmp4':  [ ['5- Récapitulatif PFMP', 'B16'], ['PFMP', 10, 8] ]
    },
    
    '31224': { etc.
    }
}
```

- deux fichiers d'export Cyclades fictifs sont fournis, pour exemple _(cyclade1.csv, cyclade2.csv)_.

<br>

fin