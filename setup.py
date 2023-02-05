#!/usr/bin/env python3
#
# module setup.py
# v2023-02-05
# 🟢⚠❌📌‼❓🔷👉⌨️


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
DIPLOMES                    = { '31212': "Métiers de l'accueil",
                                '31213': "Met.com.ven.op.A Ani.ges.esp.com.",
                                '31214': "Met.com.ven.Op.B Pr.cl.va.of.com.",
                                '31224': "CAP Équipier polyvalent du commerce"}  # !!todo!! : vérifier code CAP EPC

DIPLOMES_COURTS             = { '31212': "bacpro_MA",
                                '31213': "bacpro_MVC_A_AGEC",
                                '31214': "bacpro_MVC_B_PC",
                                '31224': "CAP_EPC"}                              # !!todo!! : vérifier code CAP EPC

# fichiers "modèles" candidats :
CANDIDATS_TEMPLATE_SHEET    = '1-Candidat, établissement'
CANDIDATS_TEMPLATE_DICT     = { 'session': 'A3',
                                'etab': 'A4',
                                'UAI': 'A5',

                                'nom': 'A7',
                                'prenom': 'A8',
                                'daten': 'A9',
                                'numcandidat': 'A10',
                                'division': 'A11',
                                'code': 'A12'}  # !!PB!! à modifier !

# fichiers "modèles" établissement :
ETAB_TEMPLATE_SHEET         = '1-Candidat, établissement'
ETAB_TEMPLATE_DICT          = { 'session': 'A3',
                                'etab': 'A4',
                                'UAI': 'A5',

                                'nom': 'A7',
                                'prenom': 'A8',
                                'daten': 'A9',
                                'numcandidat': 'A10',
                                'division': 'A11',
                                'code': 'A12'}  # !!PB!! à modifier !

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
