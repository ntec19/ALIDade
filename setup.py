#!/usr/bin/env python3
#
# module setup.py 
# v20230403


################################################################
# CONSTANTES :

# préfixe répertoire des dossiers des candidats
CANDIDATS_FOLDER_PREFIX     = 'candidats_'

# préfixe répertoire des dossiers de synthèse établissement
ETAB_FOLDER_PREFIX          = 'synthese_'

# préfixe répertoire des fichiers modèles
TEMPLATES_FOLDER            = 'MODELES'

# fichiers "source de données" : CYCLADE
CYCLADE_PREFIX              = "cyclade"  # les fichiers CSV exportés de Cyclade doivent être commencés par ...

# diplômes possibles :
DIPLOMES                    = { '31212': "Bac pro Métiers de l'accueil",
                                '31213': "Bac pro Met.com.ven.op.A Ani.ges.esp.com.",
                                '31214': "Bac pro Met.com.ven.Op.B Pr.cl.va.of.com.",
                                '31224': "CAP Équipier polyvalent du commerce"}

DIPLOMES_COURTS             = { '31212': "bacpro_MA",
                                '31213': "bacpro_MVC_A_AGEC",
                                '31214': "bacpro_MVC_B_PVOC",
                                '31224': "CAP_EPC"}

# fichiers "modèles" candidats :
CANDIDATS_TEMPLATE_SHEET    = '1-Candidat, établissement'
CANDIDATS_TEMPLATE_DICT     = { 'nom': 'E17',
                                'prenom': 'E19',
                                'numcandidat': 'E21',
                                'division': 'E27',
                                'etab': 'E29'       # dans cette cellule : "UAI + établissement",
                                #'session': 'G2',   -> info déjà présente
                                #'daten': 'A9',     -> non utilisé                                
                                }

# dictionnaire de correspondance entre
#      ->   fichiers individuels des candidats
#  et  ->   fichiers de synthese établissement

CORRESPONDANCE_CANDIDATS_SYNTHESE = {
    # source :          fichier individuel du candidat
    # destination :     fichier de synthese établissement
    
    '31212': {
       #'champ' : [ [feuille_source, cellule_source], [feuille_destination, première_ligne_des_données, colonne ] ]
        'nom':      [ ['1-Candidat, établissement', 'E17'], ['RECAPNOTES', 12, 2] ],
        'prenom':   [ ['1-Candidat, établissement', 'E19'], ['RECAPNOTES', 12, 3] ],
        'n_cand':   [ ['1-Candidat, établissement', 'E21'], ['RECAPNOTES', 12, 1] ],
        'E31':      [ ['4- Synthèse', 'Q12'], ['RECAPNOTES', 12, 4] ],
        'E32':      [ ['4- Synthèse', 'Q16'], ['RECAPNOTES', 12, 5] ],
        'pfmpdéro': [ ['5- Récapitulatif PFMP', 'D11'], ['PFMP', 10, 4] ],
        'pfmp1':    [ ['5- Récapitulatif PFMP', 'C14'], ['PFMP', 10, 5] ],
        'pfmp2':    [ ['5- Récapitulatif PFMP', 'C15'], ['PFMP', 10, 6] ],
        'pfmp3':    [ ['5- Récapitulatif PFMP', 'C16'], ['PFMP', 10, 7] ],
        'pfmp4':    [ ['5- Récapitulatif PFMP', 'C17'], ['PFMP', 10, 8] ],
        'pfmp5':    [ ['5- Récapitulatif PFMP', 'C18'], ['PFMP', 10, 9] ],
        'pfmp6':    [ ['5- Récapitulatif PFMP', 'C19'], ['PFMP', 10, 10] ]
    },
    
    '31213': {
       #'champ' : [ [feuille_source, cellule_source], [feuille_destination, première_ligne_des_données, colonne ] ]
        'nom':      [ ['1-Candidat, établissement', 'E17'], ['RECAPNOTES', 12, 2] ],
        'prenom':   [ ['1-Candidat, établissement', 'E19'], ['RECAPNOTES', 12, 3] ],
        'n_cand':   [ ['1-Candidat, établissement', 'E21'], ['RECAPNOTES', 12, 1] ],
        'E31':      [ ['4- Synthèse', 'Q12'], ['RECAPNOTES', 12, 4] ],
        'E32':      [ ['4- Synthèse', 'Q16'], ['RECAPNOTES', 12, 5] ],
        'E33':      [ ['4- Synthèse', 'Q20'], ['RECAPNOTES', 12, 6] ],
        'pfmpdéro': [ ['5- Récapitulatif PFMP', 'E11'], ['PFMP', 10, 4] ],
        'pfmp1':    [ ['5- Récapitulatif PFMP', 'D14'], ['PFMP', 10, 5] ],
        'pfmp2':    [ ['5- Récapitulatif PFMP', 'D15'], ['PFMP', 10, 6] ],
        'pfmp3':    [ ['5- Récapitulatif PFMP', 'D16'], ['PFMP', 10, 7] ],
        'pfmp4':    [ ['5- Récapitulatif PFMP', 'D17'], ['PFMP', 10, 8] ],
        'pfmp5':    [ ['5- Récapitulatif PFMP', 'D18'], ['PFMP', 10, 9] ],
        'pfmp6':    [ ['5- Récapitulatif PFMP', 'D19'], ['PFMP', 10, 10] ]
    },
    
    '31214': {
       #'champ' : [ [feuille_source, cellule_source], [feuille_destination, première_ligne_des_données, colonne ] ]
        'nom':      [ ['1-Candidat, établissement', 'E17'], ['RECAPNOTES', 12, 2] ],
        'prenom':   [ ['1-Candidat, établissement', 'E19'], ['RECAPNOTES', 12, 3] ],
        'n_cand':   [ ['1-Candidat, établissement', 'E21'], ['RECAPNOTES', 12, 1] ],
        'E31':      [ ['4- Synthèse', 'Q12'], ['RECAPNOTES', 12, 4] ],
        'E32':      [ ['4- Synthèse', 'Q16'], ['RECAPNOTES', 12, 5] ],
        'E33':      [ ['4- Synthèse', 'Q20'], ['RECAPNOTES', 12, 6] ],
        'pfmpdéro': [ ['5- Récapitulatif PFMP', 'E11'], ['PFMP', 10, 4] ],
        'pfmp1':    [ ['5- Récapitulatif PFMP', 'D14'], ['PFMP', 10, 5] ],
        'pfmp2':    [ ['5- Récapitulatif PFMP', 'D15'], ['PFMP', 10, 6] ],
        'pfmp3':    [ ['5- Récapitulatif PFMP', 'D16'], ['PFMP', 10, 7] ],
        'pfmp4':    [ ['5- Récapitulatif PFMP', 'D17'], ['PFMP', 10, 8] ],
        'pfmp5':    [ ['5- Récapitulatif PFMP', 'D18'], ['PFMP', 10, 9] ],
        'pfmp6':    [ ['5- Récapitulatif PFMP', 'D19'], ['PFMP', 10, 10] ]
    },
    
    '31224': {
       #'champ' : [ [feuille_source, cellule_source], [feuille_destination, première_ligne_des_données, colonne ] ]
        'nom':      [ ['1-Candidat, établissement', 'E17'], ['RECAPNOTES', 12, 2] ],
        'prenom':   [ ['1-Candidat, établissement', 'E19'], ['RECAPNOTES', 12, 3] ],
        'n_cand':   [ ['1-Candidat, établissement', 'E21'], ['RECAPNOTES', 12, 1] ],
        'noteEP1':  [ ['5- Synthèse', 'Q12'], ['RECAPNOTES', 12, 4] ],
        'noteEP2':  [ ['5- Synthèse', 'Q16'], ['RECAPNOTES', 12, 5] ],
        'noteEP3':  [ ['5- Synthèse', 'Q20'], ['RECAPNOTES', 12, 6] ],
        'pfmpdéro': [ ['5- Récapitulatif PFMP', 'C12'], ['PFMP', 10, 4] ],
        'pfmp1':    [ ['5- Récapitulatif PFMP', 'C15'], ['PFMP', 10, 5] ],
        'pfmp2':    [ ['5- Récapitulatif PFMP', 'C16'], ['PFMP', 10, 6] ],
        'pfmp3':    [ ['5- Récapitulatif PFMP', 'C17'], ['PFMP', 10, 7] ],
        'pfmp4':    [ ['5- Récapitulatif PFMP', 'C18'], ['PFMP', 10, 8] ],
        'pfmp5':    [ ['5- Récapitulatif PFMP', 'C19'], ['PFMP', 10, 9] ],
        'pfmp6':    [ ['5- Récapitulatif PFMP', 'C20'], ['PFMP', 10, 10] ]
    }

}


TEMPO           = 0.1
NSEP            = 32  # pour l'affichage des séparateurs
NEWLINE         = "\n"  # pour le saut de ligne dans le print de DOC

# pour sanitize()
CHAR_SEP        = "+"
CHAR_SUB        = "_"
MAJ             = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
SPE             = "ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÑÒÓÔÕÖŒÙÚÛÜ"
NBR             = "0123456789-"
LEG_CHAR        = MAJ + MAJ.lower() + SPE + SPE.lower() + NBR


################################################################
# FONCTIONS :

try:
    import sys
    import os
    import time
    import shutil
    import csv
except:
    print("❌ Une des bibliothèques standards est manquante !\n")
    sys.exit(1)


try:
    import openpyxl
except:
    print("❌ La bibliothèque suivante n'est pas installée : openpyxl\n")
    sys.exit(2)


def clear():
    if os.name == 'nt':
        os.system('cls')
    else:
        os.system('clear')


def stamp():
    now = time.localtime()
    res = time.strftime("%Y%m%d_%H%M%S", now)
    return res


def sanitize(s):
    res = ""
    for letter in s:
        if letter in LEG_CHAR:
            res += letter
        else:
            res += CHAR_SUB
    return res


def touche():
    input("-" * NSEP + "\n⌨️ Appuyez sur la touche 'Entrée' pour continuer.\n" + "_" * NSEP + "\n")


def info(s):
    print(f"\n{'-'*NSEP}\nℹ️ : {s}\n{'_'*NSEP}\n")


################################################################
# fin du module 'setup.py'
