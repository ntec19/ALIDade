#!/usr/bin/env python3
#
# script consolide_grilles.py
# v2023-01-30
# 🟢⚠❌📌‼❓🔷👉⌨️
# doc openpyxl : https://openpyxl.readthedocs.io


# import du module 'setup', contenant les constantes et fonctions communes
from setup import *


################################################################
# CONSTANTES SPECIFIQUES :

DOC = f"""

* - * - * - * - * - * - * - * - * - * - * - * - * - * - * - *
  Script pour générer les fichiers de synthèse établissement
* - * - * - * - * - * - * - * - * - * - * - * - * - * - * - *

Pour que ce script fonctionne correctement, il est nécessaire de vérifier
les prérequis suivants :
- Python 3 installé ; si ce message s'affiche, c'est sûrement le cas ;-).
- La bibliothèque Python 'openpyxl' est installée.
- Dans le même répertoire que le script, se trouvent le répertoire
  qui contient les sous-répertoires de diplômes, puis les fichiers
  individuels des candidats.
  👉 Ce répertoire commence par "{CANDIDATS_FOLDER_PREFIX}".
- Dans le répertoire "{TEMPLATES_FOLDER}", se trouvent les
  👉 fichiers modèles "établissement" nécessaires :
{NEWLINE.join(f'        "{key}_etab.xlsx" -> {value}' for key, value in DIPLOMES.items())}
  avec une feuille "{ETAB_TEMPLATE_SHEET}".

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
# déterminer le répertoire 'établissement'
# ie. : candidats_0921234A
# récupérer une liste de tous les répertoires du répertoire courant
folders     = os.listdir(".")
folders     = [f for f in folders if os.path.isdir(f)]  # exclure les fichiers
# filtrer les répertoires dont le nom commence par CANDIDATS_FOLDER_PREFIX
folders     = [f for f in folders if f[0:len(CANDIDATS_FOLDER_PREFIX)] == CANDIDATS_FOLDER_PREFIX]
# filtrer les répertoires dont le nom continue par 8 caractères (UAI)
folders     = [f for f in folders if len(f[len(CANDIDATS_FOLDER_PREFIX):]) == 8]
# filtrer les répertoires dont le nom continue par 7 chiffres
folders     = [f for f in folders if f[len(CANDIDATS_FOLDER_PREFIX):len(CANDIDATS_FOLDER_PREFIX) + 7].isdigit()]
# filtrer les répertoires dont le nom continue par 1 lettre
folders     = [f for f in folders if f[len(CANDIDATS_FOLDER_PREFIX) + 7:].isalpha()]
if len(folders) == 0:
    print(f"❌ Un problème est survenu (pas de dossier candidats trouvé) !\n")
    sys.exit()
if len(folders) != 1:
    print(f"❌ Un problème est survenu (ambiguïté sur les dossiers candidats) !\n")
    print(folders)
    sys.exit()
candidats_folder = folders[0] 
etab_uai = candidats_folder[len(CANDIDATS_FOLDER_PREFIX):]
info(f"UAI : {etab_uai} - Dossier candidats trouvé : {candidats_folder}")


################################################################
# déterminer les sous-répertoires 'diplômes' du répertoire 'établissement'

# récupérer une liste de tous les fichiers du répertoire 'établissement'
folders_diplomes    = os.listdir("./" + candidats_folder)
folders_diplomes     = [f for f in folders_diplomes if os.path.isdir("./" + candidats_folder + "/" + f)]  # exclure les fichiers
# vérifier que les noms des dossiers diplômes commencent bien par un code connu (32212, etc.)
for folder in folders_diplomes:
    if folder[0:5] not in DIPLOMES.keys():
        print(f"❌ Un problème est survenu avec ce dossier diplôme non conforme : {folder} !\n")
        sys.exit()
# trier par ordre alpha
folders_diplomes.sort()

# affichage intermédiaire : nombre de dossiers et noms de ces dossiers
message     =  "Nombre de dossiers 'diplômes' trouvés : "
message     += str(len(folders_diplomes)) + " :\n\t"
message     += "\n\t".join(folders_diplomes)
info(message)
touche()


################################################################
# vérifier l'existence des fichiers modèle ETABLISSEMENT nécessaires

# récupérer une liste de tous les fichiers du répertoire TEMPLATES_FOLDER
files = os.listdir("./"+TEMPLATES_FOLDER)
files = [f for f in files if os.path.isfile("./"+TEMPLATES_FOLDER+'/'+f)]  # exclure les répertoires

for folder in folders_diplomes:
    if folder[:5] + "_etab.xlsx" not in files:
        print(f"❌ Un problème est survenu : fichier modèle établisement  {folder[:5]}_etab.xlsx inexistant !\n")
        sys.exit()


################################################################
# création du répertoire pour les fichiers de synthèse établissement
#
# si le dossier existe, le renommer
if os.path.exists(ETAB_FOLDER_PREFIX + etab_uai):
    t = stamp()
    print(f"⚠️ Le répertoire \"{ETAB_FOLDER_PREFIX + etab_uai}\" existe déjà :\nil a été renommé en \"{ETAB_FOLDER_PREFIX}_old_" + t + "\".\n")
    os.rename(ETAB_FOLDER_PREFIX + etab_uai, ETAB_FOLDER_PREFIX + "_old_" + t)
# créer le dossier synthese_UAI
print(f"🟢 Création du répertoire \"{ETAB_FOLDER_PREFIX + etab_uai}\".\n")
os.mkdir(ETAB_FOLDER_PREFIX + etab_uai)


info("!!!!    STOP    !!!!")
touche()
# le dossier synthese_0921234A est créé : à suivre ;-)
# to be continued


################################################################
# pour chaque dossier diplome trouvé,
#     pour chaque fichier dans le dossier diplome
#         traitement :

# on boucle sur chaque dossier :
for folder in folders_diplomes:
    # récupérer le code diplôme
    code_diplome = folder[:5]
    # copie du fichier modèle
    source = "./" + TEMPLATES_FOLDER + "/" + code_diplome + "_etab.xlsx"
    destination = code_diplome + "_" + etab_uai + "_" + DIPLOMES_COURTS[code_diplome] + ".xlsx"
    shutil.copyfile(source, destination)
    # récupérer le chemin relatif du dossier
    current_folder = "./" + candidats_folder + "/" + folder + "/"
    # récupérer la liste de tous ses fichiers
    files = os.listdir(current_folder)
    files = [f for f in files if os.path.isfile(current_folder + "/" + f)]  # exclure les répertoires
    if len(files) == 0:
        print(f"❌ Un problème est survenu : dossier {current_folder} vide !\n")
        sys.exit()
    # trier par ordre alpha
    files.sort()
    
    # on boucle ensuite sur chaque fichier :
    for file in files:
        print(file)
        print("*********************************\n")

        wb_candidat = openpyxl.load_workbook(current_folder + file, read_only=True, data_only=True)
        sheet = wb_candidat[CANDIDATS_TEMPLATE_SHEET]
        valeur = sheet[CANDIDATS_TEMPLATE_DICT['session']]  # etc.
        wb_candidat.close()

        wb_etab = openpyxl.load_workbook(destination, read_only=False, data_only=True)
        sheet = wb_etab[CANDIDATS_TEMPLATE_SHEET]
        valeur = sheet[CANDIDATS_TEMPLATE_DICT['session']]  # etc.
        wb_etab.close()


info("!!!!    STOP    !!!!")
touche()



'''
CANDIDATS_TEMPLATE_DICT   = {'session': 'A3',
                   'etab': 'A4',
                   'UAI': 'A5',

                   'nom': 'A7',
                   'prenom': 'A8',
                   'daten': 'A9',
                   'numcandidat': 'A10',
                   'division': 'A11',
                   'code': 'A12'}  # !!PB!! à modifier !
                   
folders     = [f for f in folders if os.path.isdir(f)]  # exclure les fichiers
# filtrer les répertoires dont le nom commence par CANDIDATS_FOLDER_PREFIX
folders     = [f for f in folders if f[0:len(CANDIDATS_FOLDER_PREFIX)] == CANDIDATS_FOLDER_PREFIX]
# filtrer les répertoires dont le nom continue par 8 caractères (UAI)
folders     = [f for f in folders if len(f[len(CANDIDATS_FOLDER_PREFIX):]) == 8]
# filtrer les répertoires dont le nom continue par 7 chiffres
folders     = [f for f in folders if f[len(CANDIDATS_FOLDER_PREFIX):len(CANDIDATS_FOLDER_PREFIX) + 7].isdigit()]
# filtrer les répertoires dont le nom continue par 1 lettre
folders     = [f for f in folders if f[len(CANDIDATS_FOLDER_PREFIX) + 7:].isalpha()]

'''

info("STOP")
touche()





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
    with open(data, encoding='utf-8-sig') as f:
        reader = list(csv.reader(f, delimiter=';', quotechar="'"))
        # liste de liste ; reader[r] : chaque ligne ; reader[r][c] : chaque cellule
        session     = reader[0][0]
        etab        = reader[2][0]
        etab_nom    = etab.split('(')[0][:-1]
        etab_uai    = etab.split('(')[1][:-1]
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
        reader = reader[9:]     # les données commencent à la ligne 10
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
# tous les éléments de la liste 'diplomes' doivent être dans les clés du dictionnaire 'DIPLOMES' (cf. constantes)
for diplome in diplomes:
    if not (diplome in DIPLOMES.keys()):
        print(f"Un diplôme inconnu est trouvé : {diplome}.")
        sys.exit(6)


################################################################
# affichage des diplômes trouvés (code + intitulé)
info_diplomes = str(len(diplomes)) + " diplôme(s) trouvé(s) :\n"
for d in diplomes:
    info_diplomes += d + ' : '
    info_diplomes += DIPLOMES[d] + '\n'
info(info_diplomes[:-1])


touche()


################################################################
# vérification de l'existence des fichiers modèles pour chaque diplôme
for d in diplomes:
    if d + '.xlsx' not in files:
        print(f"❌ Un fichier modèle est manquant : {d + '.xlsx'} !\n")
        sys.exit(7)


################################################################
# vérification de l'existence d'une feuille CANDIDATS_TEMPLATE_SHEET dans chaque fichier modèle
for d in diplomes:
    classeur = d + ".xlsx"
    wb = openpyxl.load_workbook(classeur, read_only=True, data_only=True)
    if CANDIDATS_TEMPLATE_SHEET not in wb.sheetnames:
        print(f"❌ Le fichier \"{classeur}\" doit posséder une feuille \"{CANDIDATS_TEMPLATE_SHEET}\" !\n")
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
    ################################################################
    # copie du fichier 'modèle' vers le fichier 'candidat' dans le bon sous-dossier
    folder      =  "./" + CANDIDATS_FOLDER_PREFIX + etab_uai + "/"
    folder      += candidat[5] + "-"
    folder      += DIPLOMES_COURTS[candidat[5]] + "/"
    filename    =  sanitize(candidat[0]) + "+"
    filename    += sanitize(candidat[1]) + "+"
    filename    += candidat[5] + "+"
    filename    += candidat[3] + ".xlsx"
    print("\n" + "-" * 32)
    print(f"Candidat traité : {candidat[0]} {candidat[1]}, né(e) le {candidat[2]}")
    print(f"Diplôme : {DIPLOMES_COURTS[candidat[5]]} (code : {candidat[5]})")
    print(f"Division : {candidat[4]} - Numéro de candidat : {candidat[3]}")
    print(f"Nom du dossier : {folder}")
    print(f"Nom du fichier : {filename}")
    # exemple source :
    #   31212.xlsx
    # exemple destination :
    #   ./0921234A_candidats/31212-bacpro_MA/DURAND+Clara+31212+06916557742.xlsx
    source      = candidat[5] + ".xlsx"
    destination = folder + filename
    # print(f"{source} -> {destination}")
    shutil.copyfile(source, destination)
    time.sleep(TEMPO)  # pour 'terminer' l'écriture du fichier
    #
    ################################################################
    # personnalisation des fichiers candidats (insertion des valeurs)
    # pour mémoire :
    # clés de CANDIDATS_TEMPLATE_DICT :
    # 'session', 'etab', 'UAI', 'nom', 'prenom', 'daten', 'numcandidat', 'division', 'code'
    # pour mémoire :
    # candidats = [ [ 'Nom', 'Prénom', 'Date de Naissance', 'N° Candidat', 'Division', 'Code' ], etc. ]
    #
    wb = openpyxl.load_workbook(destination, read_only=False)
    sheet = wb[CANDIDATS_TEMPLATE_SHEET]
    sheet[CANDIDATS_TEMPLATE_DICT['session']]     = session
    sheet[CANDIDATS_TEMPLATE_DICT['etab']]        = etab_nom
    sheet[CANDIDATS_TEMPLATE_DICT['UAI']]         = etab_uai
    sheet[CANDIDATS_TEMPLATE_DICT['nom']]         = candidat[0]
    sheet[CANDIDATS_TEMPLATE_DICT['prenom']]      = candidat[1]
    sheet[CANDIDATS_TEMPLATE_DICT['daten']]       = candidat[2]
    sheet[CANDIDATS_TEMPLATE_DICT['numcandidat']] = candidat[3]
    sheet[CANDIDATS_TEMPLATE_DICT['division']]    = candidat[4]
    sheet[CANDIDATS_TEMPLATE_DICT['code']]        = candidat[5]
    wb.save(destination)
    wb.close()

msg_fin = f"""

🟢 Les fichiers des candidats sont créés :

Dans le dossier "{CANDIDATS_FOLDER_PREFIX}{etab_uai}", un sous-dossier est
préparé par diplôme.
Chacun d'entre eux contient les fichiers individuels des candidats,
avec les informations nominatives mises à jour.

That's all folks!

"""

info(msg_fin)

# fin



