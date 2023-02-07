#!/usr/bin/env python3
#
# script consolide_grilles.py
# v2023-02-07f
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
# déterminer le répertoire candidats de l'établissement : 'candidats_rootfolder'
# ex: 'candidats_0921234A'
# et l'UAI 'etab_uai', ex: '0921234A'
#
# récupérer une liste de tous les répertoires du répertoire courant
folders     = os.listdir(".")
folders     = [f for f in folders if os.path.isdir(f)]  # exclure les fichiers
# filtrer les répertoires dont le nom commence par candidats_rootfolder_PREFIX
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
candidats_rootfolder = folders[0] 
etab_uai = candidats_rootfolder[len(CANDIDATS_FOLDER_PREFIX):]
info(f"UAI : {etab_uai} - Dossier candidats trouvé : {candidats_rootfolder}")


################################################################
# récupérer une liste 'candidats_subfolders' de tous les fichiers du répertoire 'établissement'
# ex: [ '31212-bacpro_MA', etc. ]
candidats_subfolders    = os.listdir("./" + candidats_rootfolder)
candidats_subfolders     = [f for f in candidats_subfolders if os.path.isdir("./" + candidats_rootfolder + "/" + f)]  # exclure les fichiers
# vérifier que les noms des dossiers diplômes commencent bien par un code connu (32212, etc.)
for folder in candidats_subfolders:
    if folder[0:5] not in DIPLOMES.keys():
        print(f"❌ Un problème est survenu avec ce dossier diplôme non conforme : {folder} !\n")
        sys.exit()
# trier par ordre alpha
candidats_subfolders.sort()


################################################################
# construire la liste 'etab_diplomes', sous-ensemble des clés de 'DIPLOMES'
# qui ne contient que les *codes* de diplômes *réellement trouvés*
# dans le dossier 'candidats_rootfolder'
# ex: [ '31212', '31213', etc ]
etab_diplomes = []
for folder in candidats_subfolders:
    etab_diplomes.append(folder[0:5])


################################################################
# construction d'un dictionnaire 'candidats_folders'
# ex: { '31212': './candidats_0921234A/31212-bacpro_MA' }
candidats_folders = {}
for diplome in etab_diplomes:
    candidats_folders[diplome] = './' + candidats_rootfolder + '/' + diplome + "-" + DIPLOMES_COURTS[diplome]


################################################################
# vérification de la cohérence
# entre les noms de dossiers candidats réels
# et les noms attendus
candidats_subfolders_attendus = [v.split('/')[2] for v in candidats_folders.values()]
if candidats_subfolders_attendus != candidats_subfolders:
    print(f"❌ Un problème est survenu avec le nommage des sous-dossiers des candidats :")
    print(set(candidats_subfolders_attendus).symmetric_difference(set(candidats_subfolders)))
    sys.exit()


################################################################
# affichage intermédiaire : nombre de dossiers et noms de ces dossiers
message     =  "Nombre de dossiers 'diplômes' trouvés : "
message     += str(len(candidats_subfolders)) + " :\n\t"
message     += "\n\t".join(candidats_subfolders)
info(message)
touche()


################################################################
# vérifier l'existence des fichiers modèle ETABLISSEMENT nécessaires
# récupérer une liste de tous les fichiers du répertoire TEMPLATES_FOLDER
files = os.listdir("./"+TEMPLATES_FOLDER)
files = [f for f in files if os.path.isfile("./"+TEMPLATES_FOLDER+'/'+f)]  # exclure les répertoires

for diplome in etab_diplomes:
    if diplome + "_etab.xlsx" not in files:
        print(f"❌ Un problème est survenu : fichier modèle établisement  {diplome}_etab.xlsx inexistant !\n")
        sys.exit()


################################################################
# création du répertoire pour les fichiers de synthèse établissement
etab_folder = ETAB_FOLDER_PREFIX + etab_uai  # ex: synthese_0921234A
# si le dossier existe, le renommer
if os.path.exists(etab_folder):
    t = stamp()
    os.rename(etab_folder, etab_folder + "_old_" + t)
    print(f"⚠️ Le répertoire \"{etab_folder}\" existe déjà :\nil a été renommé en \"{etab_folder}_old_" + t + "\".\n")
# créer le dossier synthese_UAI
os.mkdir(etab_folder)
print(f"🟢 Création du répertoire \"{etab_folder}\".\n")


################################################################
# créer le fichier de synthèse de l'établissement pour tous les diplômes
# de la liste 'etab_diplomes' (ex: ['31212', '31213', etc])
#, dans le dossier 'etab_folder' (ex: 'synthese_0921234A')
# par copie du fichier modèle, source = 'TEMPLATES_FOLDER'
# et
# construction d'un dictionnaire 'etab_syntheses' :
# { '31212': './synthese_0921500F/31212_0921234A_bacpro_MA.xlsx', etc}
etab_syntheses = {}
for diplome in etab_diplomes:
    source = './' + TEMPLATES_FOLDER + '/' + diplome + '_etab.xlsx'  # ex: ./MODELES/31212_etab.xlsx
    destination = './' + etab_folder + '/' + diplome + "_" + etab_uai + "_" + DIPLOMES_COURTS[diplome] + ".xlsx"  # ex: ./synthese_0921500F/31212_0921234A_bacpro_MA.xlsx
    etab_syntheses[diplome] = destination
    shutil.copyfile(source, destination)


################################################################
# traitement des fichiers individuels des candidats :
# pour chaque diplome (répertoire) :
#     pour chaque candidat (fichier xlsx) :
#         lire les infos (nom, prenom, ncand, note1, etc.

for diplome in etab_diplomes:
    
    data_candidats = []  # contiendra des dictionnaires qui stockeront les données de chaque candidat du diplôme
    print("\nTraitement du dossier '" + candidats_folders[diplome] + "'...\n")
    time.sleep(TEMPO*10)
    files = os.listdir(candidats_folders[diplome])
    files.sort()
    for file in files:
        candidat = {}  # contiendra les données du candidat
        # print("\tOn traite le fichier :", file)
        # on ouvre le fichier 'file' avec openpyxl
        wb_candidat = openpyxl.load_workbook(candidats_folders[diplome]+'/'+file, read_only=True, data_only=True)
        for k in CORRESPONDANCE_CANDIDATS_SYNTHESE[diplome].keys():
            # CORRESPONDANCE_CANDIDATS_SYNTHESE[diplome][k] -> [['1-Candidat, établissement', 'A7'], ['RECAPNOTES', 12, 2]]
            v = wb_candidat[CORRESPONDANCE_CANDIDATS_SYNTHESE[diplome][k][0][0]][CORRESPONDANCE_CANDIDATS_SYNTHESE[diplome][k][0][1]].value
            candidat[k] = v
        wb_candidat.close()
        data_candidats.append(candidat)
    #info("Informations de debug : " + str(data_candidats))
    
    # écriture dans le fichier de synthèse établissement
    wb_synthese = openpyxl.load_workbook(etab_syntheses[diplome], read_only=False, data_only=True)
    for k in CORRESPONDANCE_CANDIDATS_SYNTHESE[diplome].keys():
        # k prend ses valeurs dans : 'nom', 'prenom', 'date_n', etc.
        first_line  = CORRESPONDANCE_CANDIDATS_SYNTHESE[diplome][k][1][1]
        colonne     = CORRESPONDANCE_CANDIDATS_SYNTHESE[diplome][k][1][2]
        
        line = first_line
        for candidat in data_candidats:
            v = candidat[k]
            wb_synthese[CORRESPONDANCE_CANDIDATS_SYNTHESE[diplome][k][1][0]].cell(row=line, column=colonne).value = v
            line += 1
    
    wb_synthese.save(etab_syntheses[diplome])
    wb_synthese.close()


msg_fin = f"""🟢 🟢 🟢 🟢 🟢 🟢 🟢 🟢

🟢 Les fichiers de synthèse de l'établissement sont créés :

Dans le dossier "{etab_folder}", un fichier Excel est
préparé par diplôme.
Chacun d'entre eux contient la consolidation des informations
des candidats : état civil, numéro de candidat, notes, etc.

That's all folks again!

"""

info(msg_fin)

# fin
