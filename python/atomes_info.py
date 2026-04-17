# -*- coding: utf-8 -*-

"""
LLM tools (Claude, Gemini, GPT, Lechat) were used at different occasions to prepare this file
"""

"""
atomes_info.py — Configuration centralisée de l'extension atomes pour LibreOffice
==================================================================================

Ce fichier est l'endroit unique où figurent toutes les constantes, identifiants,
commandes, options et paramètres spécifiques au logiciel atomes.

Objectif générique
------------------
Cette extension a été conçue comme un modèle générique d'extension LibreOffice
permettant d'insérer des fichiers d'un logiciel tiers dans un document ODF.
Pour l'adapter à un autre logiciel, il suffit de modifier ce fichier uniquement.

Organisation des sections
--------------------------
  1. Identifiants de l'extension (UNO, Java, menus)
  2. Commande exécutable et options subprocess
  3. Version minimale requise
  4. Format de fichier natif (extension, filtres, icône)
  5. Marqueurs et préfixes ODF (stockage dans le document)
  6. Paramètres d'insertion graphique
  7. Notes sur les fichiers XML de configuration (hors Python)

IMPORTANT : ce fichier ne doit importer AUCUN autre module de l'extension
(ni atomes_extension, ni atomes_i18n, ni atomes_options) afin d'éviter
tout risque de dépendance cyclique lors du chargement par LibreOffice.
"""

import re
import os
import tempfile

# ══════════════════════════════════════════════════════════════════════
# 1. IDENTIFIANTS DE L'EXTENSION
# ══════════════════════════════════════════════════════════════════════
#
# Ces identifiants sont utilisés par LibreOffice pour enregistrer et localiser
# l'extension. Ils doivent correspondre exactement aux valeurs déclarées dans
# les fichiers XML de configuration (Addons.xcu, Jobs.xcu, description.xml).
#
# Convention : identifiants en notation inversée de domaine (reverse-DNS),
# à la manière des packages Java/Android.

# Identifiant principal pour la création des variables de ce fichier
atomes = "atomes"

# Identifiant principal de l'extension (correspond à <identifier> dans description.xml
# et à getPackageLocation() dans le code Python)
atomes_EXTENSION_ID = "fr.ipcms.atomes.extension"

# Identifiant UNO du service principal (déclenché par les entrées de menu dans Addons.xcu)
# Correspond à g_ImplementationHelper.addImplementation(...) dans atomes_service.py
# et à <value>service:fr.ipcms.atomes.atomesService?insert</value> dans Addons.xcu
atomes_SERVICE_ID = "fr.ipcms.atomes.atomesService"

# Identifiant UNO du job OnLoad (déclenché automatiquement à l'ouverture d'un document)
# Correspond à <value>fr.ipcms.atomes.atomesJob</value> dans Jobs.xcu
atomes_JOB_ID = "fr.ipcms.atomes.atomesJob"

# Nom du nœud de menu dans Addons.xcu (oor:name du nœud OfficeMenuBar)
# Utilisé uniquement dans Addons.xcu, documenté ici pour référence
atomes_MENU_NODE_ID = "fr.ipcms.atomes.extension.menu"


# ══════════════════════════════════════════════════════════════════════
# 2. COMMANDE EXÉCUTABLE ET OPTIONS SUBPROCESS
# ══════════════════════════════════════════════════════════════════════
#
# Tous les appels à subprocess.run() de l'extension passent par ces variables.
# Pour adapter l'extension à un autre logiciel, remplacer l'exécutable et
# adapter les options de ligne de commande correspondantes.

# Nom de l'exécutable tel qu'il doit être accessible dans le PATH système
atomes_EXECUTABLE = atomes + ".exe" if os.name == 'nt' else atomes

# Option pour afficher la version du logiciel (utilisée dans _check_atomes_version)
# La sortie attendue doit contenir le pattern défini par atomes_VERSION_PATTERN ci-dessous
atomes_OPT_VERSION = "-v"

# Option pour générer un rendu PNG d'un fichier projet lors de l'insertion
# Syntaxe complète : atomes --render-png <fichier.apf> -o <sortie.png>
atomes_OPT_RENDER_PNG = "--render-png"

# Option pour ouvrir un fichier projet avec l'interface graphique et générer
# une image de prévisualisation mise à jour (utilisée au double-clic)
# Syntaxe complète : atomes --libreoffice -o <sortie.png> <fichier.apf>
atomes_OPT_LIBREOFFICE = "--libreoffice"

# Option commune pour spécifier le fichier de sortie (PNG)
atomes_OPT_OUTPUT = "-o"

# Délai maximum (en secondes) pour le rendu PNG lors de l'insertion d'un fichier
# Augmenter cette valeur pour les projets très volumineux
atomes_CMD_TIMEOUT = 120

# Délai maximum (en secondes) pour la vérification de la version installée
atomes_VERSION_TIMEOUT = 10


# ══════════════════════════════════════════════════════════════════════
# 3. VERSION MINIMALE REQUISE
# ══════════════════════════════════════════════════════════════════════
#
# L'extension vérifie au démarrage qu'une version suffisante du logiciel
# est installée. Modifier ces valeurs pour changer la version minimale acceptée.

# Version minimale sous forme de tuple (majeure, mineure, patch)
atomes_MIN_VERSION = (1, 3, 0)

# Expression régulière pour extraire la version depuis la sortie de atomes -v
# Le pattern doit capturer trois groupes numériques : majeure, mineure, patch
# Exemple de sortie attendue : "atomes version         : 1.3.0"
atomes_VERSION_PATTERN = re.compile(
    r'atomes version\s*:\s*(\d+)\.(\d+)\.(\d+)'
)


# ══════════════════════════════════════════════════════════════════════
# 4. FORMAT DE FICHIER NATIF
# ══════════════════════════════════════════════════════════════════════
#
# Paramètres décrivant le format de fichier lu par le logiciel.
# Ces valeurs alimentent les filtres des boîtes de dialogue "Ouvrir / Insérer",
# les noms de fichiers temporaires, et le type MIME dans le manifest ODF.

atomes_apf = "apf"

# Extension de fichier native du logiciel (avec le point)
atomes_FILE_EXTENSION = "." + atomes_apf

# Glob pour le filtre de fichiers dans les boîtes de dialogue FilePicker
atomes_FILE_FILTER = "*" + atomes_FILE_EXTENSION

# Type MIME inséré dans META-INF/manifest.xml pour les fichiers embarqués en mode ZIP
# Doit être un identifiant unique pour le format ; utiliser application/octet-stream
# si aucun MIME officiel n'existe pour le logiciel cible
atomes_MIME_TYPE = "application/octet-stream"

# Nom du fichier icône de repli (fallback) lorsque le rendu PNG échoue
# Ce fichier doit exister dans le sous-dossier icons/ de l'extension
atomes_ICON_FILENAME = atomes + ".svg"

# Préfixe et suffixe utilisés pour les fichiers PNG temporaires créés lors du rendu
atomes_TEMP_PNG_PREFIX = atomes + "_"
atomes_TEMP_PNG_SUFFIX = ".png"

# Répertoire temporaire du système
atomes_TEMP_DIR = tempfile.gettempdir()

# Préfixe des fichiers PNG temporaires utilisés lors de la mise à jour d'une shape
atomes_TEMP_UPDATE_PREFIX = atomes + "_update_"


# ══════════════════════════════════════════════════════════════════════
# 5. MARQUEURS ET PRÉFIXES ODF (STOCKAGE DANS LE DOCUMENT)
# ══════════════════════════════════════════════════════════════════════
#
# Ces chaînes sont utilisées comme marqueurs/clés dans les propriétés ODF du document
# et dans les métadonnées des shapes LibreOffice. Elles permettent à l'extension
# de retrouver et identifier les objets qu'elle a insérés.
#
# ATTENTION : modifier ces valeurs rendra incompatibles les documents existants
# (les anciennes shapes ne seront plus reconnues).

# Préfixe du nom (shape.Name) des GraphicObjectShapes insérées par l'extension
# Exemple : shape.Name = "atomesFile_a1b2c3d4_monprojet.apf"
atomes_SHAPE_NAME_PREFIX = atomes + "File_"

# Préfixe de la description (shape.Description) — préfixe commun identifiant
# les shapes de l'extension indépendamment du mode de stockage
atomes_SHAPE_DESCRIPTION_PREFIX = atomes + "File:"

# Préfixe de description pour le mode stockage interne (fichier embarqué)
# shape.Description = "atomesEmbed:<unique_name>"
atomes_EMBED_PREFIX = atomes + "Embed:"

# Préfixe de description pour le mode liens externes
# shape.Description = "atomesLink:/chemin/absolu/vers/fichier.apf"
atomes_LINK_PREFIX = atomes + "Link:"

# Nom du sous-dossier de stockage créé dans l'archive ODF (mode ZIP)
# Correspond à un StorageElement dans le DocumentStorage LibreOffice
atomes_ODF_STORAGE_FOLDER = atomes + "Files"

# Préfixe des propriétés UserDefined du document (mode stockage par propriétés)
# Exemple : propriété "apf_a1b2c3d4_monprojet.apf" contient le contenu Base64
atomes_PROP_FILE_PREFIX = atomes_apf + "_"

# Nom de la propriété UserDefined stockant le mode de stockage global
# Valeurs possibles : "internal" (défaut) ou "external"
atomes_PROP_STORAGE_MODE = atomes + "StorageMode"

# Nom de la propriété UserDefined stockant le sous-mode de stockage interne
# Valeurs possibles : "properties" (défaut) ou "zip"
atomes_PROP_INTERNAL_MODE = atomes + "InternalMode"

# Nom de la propriété UserDefined stockant la carte {unique_name: chemin_disque}
# Valeur : chaîne JSON sérialisée
atomes_PROP_FILE_MAP = atomes + "FileMap"


# ══════════════════════════════════════════════════════════════════════
# 6. PARAMÈTRES D'INSERTION GRAPHIQUE
# ══════════════════════════════════════════════════════════════════════
#
# Paramètres contrôlant la taille et la mise en page des shapes insérées.

# Taille par défaut (largeur, hauteur) en centièmes de millimètre (1/100 mm)
# 6000 = 6 cm. Cette valeur est utilisée avant que l'image réelle soit chargée.
atomes_DEFAULT_SHAPE_WIDTH  = 6000
atomes_DEFAULT_SHAPE_HEIGHT = 6000

# Facteur de conversion pixels → centièmes de mm (twips LibreOffice)
# Formule : taille_lo = pixels * atomes_PX_TO_LO_SCALE
# Basé sur 96 DPI standard et 1440 unités LO par pouce
atomes_PX_TO_LO_SCALE = 1440 / 96


# ══════════════════════════════════════════════════════════════════════
# 7. RÉFÉRENCE AUX FICHIERS XML DE CONFIGURATION (hors Python)
# ══════════════════════════════════════════════════════════════════════
#
# Les fichiers suivants contiennent également des identifiants spécifiques
# à atomes. Ils ne sont PAS modifiés par ce fichier Python, mais sont
# documentés ici pour guider l'adaptation de l'extension à un autre logiciel.
#
# ┌─────────────────────┬──────────────────────────────────────────────────────┐
# │ Fichier             │ Valeurs à adapter                                    │
# ├─────────────────────┼──────────────────────────────────────────────────────┤
# │ description.xml     │ <identifier> → atomes_EXTENSION_ID                   │
# │                     │ <name>, <description> → nom du logiciel              │
# ├─────────────────────┼──────────────────────────────────────────────────────┤
# │ Addons.xcu          │ oor:name du nœud menu → atomes_MENU_NODE_ID          │
# │                     │ <value>service:…?insert</value> → atomes_SERVICE_ID  │
# │                     │ Titres FR/EN des entrées de menu                     │
# ├─────────────────────┼──────────────────────────────────────────────────────┤
# │ Jobs.xcu            │ <value>fr.ipcms.atomes.atomesJob</value>             │
# │                     │   → atomes_JOB_ID                                    │
# ├─────────────────────┼──────────────────────────────────────────────────────┤
# │ META-INF/           │ Fichiers Python référencés dans manifest.xml         │
# │ manifest.xml        │ (ajouter atomes_info.py si nécessaire)               │
# └─────────────────────┴──────────────────────────────────────────────────────┘
