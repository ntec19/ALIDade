#!/usr/bin/env python3
#
# script genere_grilles_indiv.py
# v20230405
# doc openpyxl : https://openpyxl.readthedocs.io


# import du module 'setup', contenant les constantes et fonctions communes 
from setup import *


################################################################
# CONSTANTES SPECIFIQUES :

DOC = f"""

* - * - * - * - * - * - * - * - * - * - * - * - * - * - * - *
  Script pour générer les fichiers individuels des candidats
* - * - * - * - * - * - * - * - * - * - * - * - * - * - * - *

Pour que ce script fonctionne correctement, il est nécessaire de vérifier
les prérequis suivants :
- Python 3 installé ; si ce message s'affiche, c'est sûrement le cas ;-).
- La bibliothèque Python 'openpyxl' est installée.
- Dans le même répertoire que le script, se trouvent les
  👉 fichiers d'export Cyclade, qui possèdent l'extension ".csv"
  et dont le nom commence par "{CYCLADE_PREFIX}".
  Il peut y en avoir plusieurs (typiquement, un pour le CAP, un pour le
  bac pro) ; les candidats seront recherchés dans chacun d'entre eux.
- Dans le répertoire "{TEMPLATES_FOLDER}", se trouvent les
  👉 fichiers modèles "candidats" nécessaires :
{NEWLINE.join(f'        "{key}.xlsx" -> {value}' for key, value in DIPLOMES.items())}
  avec une feuille "{CANDIDATS_TEMPLATE_SHEET}".

Appuyez sur [Entrée] pour continuer, [CTRL+C] pour arrêter.

"""

################################################################
# affichage de la documentation
print(DOC)
input()


################################################################
# C'est parti ;-) :

clear()
info("Début du traitement...")


################################################################
# lire les données depuis le(s) fichier(s) Cyclade :


################################################################
# récupérer une liste de tous les fichiers du répertoire courant
files = os.listdir(".")
files = [f for f in files if os.path.isfile(f)]  # exclure les répertoires


################################################################
# construire une liste de tous les fichiers dont le nom commence par CYCLADE_PREFIX et d'extension .csv
files_cyclade = []
for f in files:
    if f[0:7].lower() == CYCLADE_PREFIX.lower() and f.split('.')[-1].lower() == 'csv':
        files_cyclade.append(f)


################################################################
# si aucun fichier cycladeXYZ.csv : sortie en erreur
if len(files_cyclade) == 0:
    print(f"❌ Fichier(s) Cyclade ({CYCLADE_PREFIX}XYZ.csv) inexistant(s)\n")
    sys.exit(3)


################################################################
# construire une liste "candidats" à partir de tous les fichiers Cyclade
candidats = []
for data in files_cyclade:
    with open(data, encoding='utf-8-sig') as f:     # utf-8-sig car le fichier Cyclades est encodé en utf-8 avec BOM
        reader = list(csv.reader(f, delimiter=';', quotechar="'"))
        # liste de liste ; reader[r] : chaque ligne ; reader[r][c] : chaque cellule
        session     = reader[0][0]
        
        
        # modification du script pour la prise en compte du problème suivant :
        # dans certains exports Cyclades, une info s'intercale en ligne 2 : "BACCALAURÉAT PROFESSIONNEL;;;;;;;;;;;;;;"
        # ce qui engendre un décalage de ligne.
        # 
        #     situation analysée initialement :
        # 2023;;;;;;;;;;;;;; 
        # Listes simples;;;;;;;;;;;;;;
        # LYC TEST – VILLEFICTIVE  (0921234A);;;;;;;;;;;;;;     <---- ligne 3 (n°2)
        # Type d'édition : Liste simple;;;;;;;;;;;;;;        
        # 
        #     solution apparue dans certains établissements :
        # 2023;;;;;;;;;;;;;; 
        # BACCALAURÉAT PROFESSIONNEL;;;;;;;;;;;;;;
        # Listes simples;;;;;;;;;;;;;;
        # LYC TEST – VILLEFICTIVE  (0921234A);;;;;;;;;;;;;;     <---- ligne 4 (n°3)
        # Type d'édition : Liste simple;;;;;;;;;;;;;;
        # 
        # solution : calculer un 'delta'
        # par rapport à la présence de '(' dans la ligne :
        
        if      '(' in reader[2][0]:
            delta = 0
        elif    '(' in reader[3][0]:
            delta = 1
        else:
            print(f"❌ Problème dans le format de l'export Cyclades : {data}\n")
            sys.exit(3)
        
        etab        = reader[2+delta][0]
        etab_nom    = etab.split('(')[0][:-1]
        etab_uai    = etab.split('(')[1][:-1]
        
        print("|||", data, "|||", delta, "|||", etab_nom, "|||", etab_uai, "|||")
        touche()
        
        ''' les données sont structurées ainsi :
        ['Division de classe', 'N° Candidat', 'N° Inscription', 'N° Océan', 'Nom de famille',
        "Nom d'usage", 'Prénom(s)', 'Date de Naissance', 'Division de classe', 'INE',
        'Catégorie Candidat', 'Code Spécialité', 'Spécialité',
        'Etat', 'Enseignements']
        exemple :
        ['TMA', '01216557741', '002 Version 2', ' -', 'DURAND',
        ' -', 'Bryan', '09/03/2005', 'TMA', '081277848GG',
        'SCOLAIRE BACPRO 3 ANS (132)', '31212', 'Métiers de l'accueil',
        'Inscrit', 'Non renseigné'] '''
        reader = reader[9+delta:]     # les données commencent à la ligne 10+delta
        for line in reader:
            # candidat : liste au format
            # ['Nom', 'Prénom', 'Date de Naissance', 'N° Candidat', 'Division', 'Code']
            candidat = [line[4], line[6], line[7], line[1], line[0], line[11]]
            candidats.append(candidat)


################################################################
# la liste 'candidats' ne doit pas être vide
if len(candidats) == 0:
    print(f"❌ Un problème est survenu (pas de candidat trouvé) !\n")
    sys.exit(4)


################################################################
# afficher les infos 'établissement' et 'candidats'
info_etab   =   "Infos établissement trouvées :\n"
info_etab   +=  "session            : " + session + "\n"
info_etab   +=  "Nom établissement  : " + etab_nom + "\n"
info_etab   +=  "UAI établissement  : " + etab_uai
info(f"{len(candidats)} candidats trouvés.")
info(info_etab)


touche()


################################################################
# on dispose désormais d'une liste "candidats" ;
# [ [ 'Nom', 'Prénom', 'Date de Naissance', 'N° Candidat', 'Division', 'Code' ], etc. ]
# chaque élément est lui-même une liste, ie. un candidat
#
# ainsi que des variables "globales" : session ; etab_nom ; etab_uai


################################################################
# extraction des diplômes qui concernent l'établissement
diplomes = []
for candidat in candidats:
    diplome = candidat[5]
    if not (diplome in diplomes):
        diplomes.append(diplome)


################################################################
# la liste 'diplomes' ne doit pas être vide
if len(diplomes) == 0:
    print(f"❌ Un problème est survenu (pas de diplôme trouvé) !\n")
    sys.exit(5)


################################################################
# analyse : chaque diplôme dans la liste 'diplomes' est-il dans les clés du dictionnaire 'DIPLOMES' ?
# sinon : message d'avertissement, mais on continue...
for diplome in diplomes:
    if not (diplome in DIPLOMES.keys()):
        print(f"⚠️ diplôme inconnu trouvé : {diplome} (non pris en charge)")
        # sys.exit(6)
        # pas de fin de script si code diplôme inconu, mais message d'avertissement


################################################################
# Ne conserver dans la liste que les codes de diplômes présents dans les clés du dictionnaire 'DIPLOMES'
# NB : aurait pu être fait ci-dessus, cf # extraction des diplômes qui concernent l'établissement
diplomes_tout = diplomes.copy()
diplomes = []
for diplome in diplomes_tout:
    if diplome in DIPLOMES.keys():
        diplomes.append(diplome)


################################################################
# la liste 'diplomes' ne doit pas être vide (bis ;-)
if len(diplomes) == 0:
    print(f"❌ Un problème est survenu (pas de diplôme conforme trouvé) !\n")
    sys.exit(5)


################################################################
# affichage des diplômes trouvés (code + intitulé)
info_diplomes = str(len(diplomes)) + " diplôme(s) trouvé(s) :\n"
for d in diplomes:
    info_diplomes += d + ' : '
    info_diplomes += DIPLOMES[d] + '\n'
info(info_diplomes[:-1])


touche()


################################################################
# récupérer une liste de tous les fichiers du répertoire TEMPLATES_FOLDER
files = os.listdir("./"+TEMPLATES_FOLDER)
files = [f for f in files if os.path.isfile("./"+TEMPLATES_FOLDER+'/'+f)]  # exclure les répertoires


################################################################
# vérification de l'existence des fichiers modèles pour chaque diplôme dans le répertoire TEMPLATES_FOLDER
for d in diplomes:
    if d + '.xlsx' not in files:
        print(f"❌ Un fichier modèle est manquant : {d + '.xlsx'} !\n")
        sys.exit(7)


################################################################
# vérification de l'existence d'une feuille CANDIDATS_TEMPLATE_SHEET dans chaque fichier modèle
for d in diplomes:
    classeur = './' + TEMPLATES_FOLDER + '/' + d + ".xlsx"
    wb = openpyxl.load_workbook(classeur, read_only=True, data_only=True)
    if CANDIDATS_TEMPLATE_SHEET not in wb.sheetnames:
        print(f"❌ Le fichier modèle  \"{classeur}\" doit posséder une feuille \"{CANDIDATS_TEMPLATE_SHEET}\" !\n")
        sys.exit(8)
    wb.close()


################################################################
# création de l'arborescence pour les fichiers individuels des candidats
#
# si le dossier existe, le renommer
if os.path.exists(CANDIDATS_FOLDER_PREFIX + etab_uai):
    t = stamp()
    print(f"⚠️ Le répertoire \"{CANDIDATS_FOLDER_PREFIX + etab_uai}\" existe déjà :\nil a été renommé en \"{CANDIDATS_FOLDER_PREFIX}_old_" + t + "\".\n")
    os.rename(CANDIDATS_FOLDER_PREFIX + etab_uai, CANDIDATS_FOLDER_PREFIX + "_old_" + t)
# créer le dossier candidats_UAI
print(f"🟢 Création du répertoire \"{CANDIDATS_FOLDER_PREFIX + etab_uai}\".\n")
os.mkdir(CANDIDATS_FOLDER_PREFIX + etab_uai)
# créer un sous dossier par diplôme
for diplome in diplomes:
    folderName  =   CANDIDATS_FOLDER_PREFIX + etab_uai
    folderName  +=  "/"
    folderName  +=  diplome + "-"
    folderName  +=  DIPLOMES_COURTS[diplome]
    os.mkdir(folderName)


touche()


################################################################
# création des fichiers individuels des candidats dans l'arborescence

# pour mémoire :
# candidats = [ [ 'Nom', 'Prénom', 'Date de Naissance', 'N° Candidat', 'Division', 'Code' ], etc. ]
# + variables "globales" : session ; etab_nom ; etab_uai
# arbo =    .    /    CANDIDATS_FOLDER_PREFIX + etab_uai    /    diplome + "-" DIPLOMES_COURTS[diplome]
# arbo =    .    /    candidats_0921500F         /    31212-bacpro_MA
# nom+prenom+code+ncandidat.xlsx
info("Traitement : création des fichiers individuels des candidats")
for candidat in candidats:

    if candidat[5] in diplomes:
        ################################################################
        # copie du fichier 'modèle' vers le fichier 'candidat' dans le bon sous-dossier
        folder      =  "./" + CANDIDATS_FOLDER_PREFIX + etab_uai + "/"
        folder      += candidat[5] + "-"
        folder      += DIPLOMES_COURTS[candidat[5]] + "/"
        filename    =  sanitize(candidat[0]) + CHAR_SEP
        filename    += sanitize(candidat[1]) + CHAR_SEP
        filename    += candidat[5] + CHAR_SEP
        filename    += candidat[3] + ".xlsx"
        # exemple source :
        #   ./MODELES/31212.xlsx
        # exemple destination :
        #   ./0921234A_candidats/31212-bacpro_MA/DURAND+Clara+31212+06916557742.xlsx
        source      = './' + TEMPLATES_FOLDER + '/' + candidat[5] + ".xlsx"
        destination = folder + filename
        shutil.copyfile(source, destination)
        time.sleep(TEMPO)  # pour 'terminer' l'écriture du fichier

        ################################################################
        # personnalisation des fichiers candidats (insertion des valeurs)
        # pour mémoire :
        # clés de CANDIDATS_TEMPLATE_DICT : 'nom', 'prenom', 'numcandidat', 'division', 'etab'
        # pour mémoire :
        # candidats = [ [ 'Nom', 'Prénom', 'Date de Naissance', 'N° Candidat', 'Division', 'Code' ], etc. ]
        wb = openpyxl.load_workbook(destination, read_only=False)
        sheet = wb[CANDIDATS_TEMPLATE_SHEET]
        sheet[CANDIDATS_TEMPLATE_DICT['etab']]        = etab_uai + " (" + etab_nom + ")"
        sheet[CANDIDATS_TEMPLATE_DICT['nom']]         = candidat[0]
        sheet[CANDIDATS_TEMPLATE_DICT['prenom']]      = candidat[1]
        sheet[CANDIDATS_TEMPLATE_DICT['numcandidat']] = candidat[3]
        sheet[CANDIDATS_TEMPLATE_DICT['division']]    = candidat[4]
        wb.save(destination)
        wb.close()
        ################################################################
        # affichage rassurant ;-)
        print("\n" + "-" * 32)
        print(f"Candidat traité : {candidat[0]} {candidat[1]}, né(e) le {candidat[2]}")
        print(f"    Diplôme : {DIPLOMES_COURTS[candidat[5]]} (code : {candidat[5]})")
        print(f"    Division : {candidat[4]} - Numéro de candidat : {candidat[3]}")
        print(f"    Nom du dossier : {folder}")
        print(f"    Nom du fichier : {filename}")


################################################################
# affichage final
msg_fin = f"""🟢 🟢 🟢 🟢 🟢 🟢 🟢 🟢

🟢 Les fichiers des candidats sont créés !

Dans le dossier "{CANDIDATS_FOLDER_PREFIX}{etab_uai}", un sous-dossier est
préparé par diplôme.
Chacun d'entre eux contient les fichiers individuels des candidats,
avec les informations nominatives mises à jour.


"""

info(msg_fin)

# fin
