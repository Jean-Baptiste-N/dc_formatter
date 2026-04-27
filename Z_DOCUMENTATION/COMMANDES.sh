#!/bin/bash

################################################################################
#
# DC FORMATTER - RÉPERTOIRE COMPLET DES COMMANDES
#
# Ce fichier documente toutes les commandes disponibles via CLI pipeline.py
# avec exemples d'utilisation pratiques
#
################################################################################

# =============================================================================
# SECTION 1: PIPELINE COMPLÈTE (RECOMMANDÉ)
# =============================================================================

# Commande: full
# Description: Orchestre les 4 phases complètes en une seule commande
#
# Syntaxe:
#   python3 -m tools3.pipeline full -s DOCUMENT.docx [-o OUTPUT_DIR]
#
# Paramètres:
#   -s, --source DOCUMENT.docx    : Nom du fichier dans DC_SOURCES/ ou chemin complet
#   -o, --output_dir OUTPUT_DIR   : Dossier parent pour OUTPUT1_*, OUTPUT2_*, etc.
#                                   (défaut: répertoire courant)

# Exemple 1: Pipeline simple
python3 -m tools3.pipeline full -s DC_JNZ_2026.docx

# Exemple 2: Pipeline avec output custom
python3 -m tools3.pipeline full -s DC_BM2.docx -o results/

# Exemple 3: Fichier avec chemin complet
python3 -m tools3.pipeline full -s /chemin/complet/document.docx

# Résultat:
#   OUTPUT1_XML-RAW/DC_JNZ_2026_GLOBAL.xml
#   OUTPUT2_JSON-RAW/DC_JNZ_2026_GLOBAL_raw.json
#   OUTPUT3_JSON-TRANSFORMED/DC_JNZ_2026_GLOBAL_transformed.json
#   OUTPUT4_DOCX-RESULT/DC_JNZ_2026_GLOBAL_formatted.docx


# =============================================================================
# SECTION 2: PIPELINE EN DEUX PHASES
# =============================================================================

# Commande: extract
# Description: Phase 1 seulement - Extraction + Conversion XML → JSON RAW
# Sortie: XML global + JSON RAW
# Utile: Quand vous voulez examiner la structure brute avant transformation

# Syntaxe:
#   python3 -m tools3.pipeline extract -s DOCUMENT.docx [-o OUTPUT_DIR]

# Exemple:
python3 -m tools3.pipeline extract -s DC_JNZ_2026.docx

# Résultat:
#   OUTPUT1_XML-RAW/DC_JNZ_2026_GLOBAL.xml
#   OUTPUT2_JSON-RAW/DC_JNZ_2026_GLOBAL_raw.json


# Commande: transform-render
# Description: Phase 2 seulement - Transformation JSON + Rendu DOCX
# Prérequis: Phase 1 doit avoir été exécutée d'abord
# Sortie: JSON transformé + DOCX final

# Syntaxe:
#   python3 -m tools3.pipeline transform-render -s DOCUMENT.docx [-o OUTPUT_DIR]

# Exemple:
python3 -m tools3.pipeline transform-render -s DC_JNZ_2026.docx

# Résultat:
#   OUTPUT3_JSON-TRANSFORMED/DC_JNZ_2026_GLOBAL_transformed.json
#   OUTPUT4_DOCX-RESULT/DC_JNZ_2026_GLOBAL_formatted.docx


# =============================================================================
# SECTION 3: COMMANDES INDIVIDUELLES (RAREMENT UTILISÉES)
# =============================================================================

# Commande: extract-dims
# Description: Extrait uniquement les dimensions du template
# Aucun paramètre requis
# Sortie: Affichage des dimensions du template par défaut

python3 -m tools3.pipeline extract-dims

# Résultat:
#   Affiche:
#   - page_width: 11906 (twips)
#   - page_height: 16838 (twips)
#   - top_margin: 1440
#   - bottom_margin: 1440
#   - left_margin: 1440
#   - right_margin: 1440
#   - usable_width: 8226


# Commande: extract-xml
# Description: Extrait le XML brut d'un DOCX et crée le XML global
# Sortie: XML global uniquement (pas de conversion JSON)

# Syntaxe:
#   python3 -m tools3.pipeline extract-xml -s DOCUMENT.docx [-o OUTPUT_DIR]

# Exemple:
python3 -m tools3.pipeline extract-xml -s DC_JNZ_2026.docx

# Résultat:
#   OUTPUT1_XML-RAW/DC_JNZ_2026_GLOBAL.xml


# Commande: xml-to-json
# Description: Convertit un XML en JSON RAW
# Entrée: Chemin du fichier XML global
# Sortie: JSON RAW avec structure complète

# Syntaxe:
#   python3 -m tools3.pipeline xml-to-json -s XML_FILE.xml [-o OUTPUT_DIR]

# Exemple:
python3 -m tools3.pipeline xml-to-json -s OUTPUT1_XML-RAW/DC_JNZ_2026_GLOBAL.xml

# Résultat:
#   OUTPUT2_JSON-RAW/DC_JNZ_2026_GLOBAL_raw.json


# Commande: transform
# Description: Transforme un JSON RAW en JSON avec tags et styles
# Entrée: Chemin du fichier JSON RAW
# Sortie: JSON TRANSFORMED avec structure enrichie

# Syntaxe:
#   python3 -m tools3.pipeline transform -s JSON_RAW.json [-o OUTPUT_DIR]

# Exemple:
python3 -m tools3.pipeline transform -s OUTPUT2_JSON-RAW/DC_JNZ_2026_GLOBAL_raw.json

# Résultat:
#   OUTPUT3_JSON-TRANSFORMED/DC_JNZ_2026_GLOBAL_transformed.json


# Commande: render
# Description: Génère un DOCX à partir d'un JSON transformé
# Entrée: Chemin du fichier JSON TRANSFORMED
# Sortie: DOCX final formaté

# Syntaxe:
#   python3 -m tools3.pipeline render -s JSON_TRANSFORMED.json [-o OUTPUT_DIR]

# Exemple:
python3 -m tools3.pipeline render -s OUTPUT3_JSON-TRANSFORMED/DC_JNZ_2026_GLOBAL_transformed.json

# Résultat:
#   OUTPUT4_DOCX-RESULT/DC_JNZ_2026_GLOBAL_formatted.docx


# =============================================================================
# SECTION 4: AIDE ET DOCUMENTATION
# =============================================================================

# Afficher l'aide générale
python3 -m tools3.pipeline --help

# Afficher l'aide d'une commande spécifique
python3 -m tools3.pipeline full --help
python3 -m tools3.pipeline extract --help
python3 -m tools3.pipeline transform-render --help


# =============================================================================
# SECTION 5: AUTRES MODULES (hors pipeline.py)
# =============================================================================

# NOTE: Ces modules peuvent être utilisés directement en Python
# ou via des scripts séparés (pas via CLI pipeline.py)

# Module: zip_docx.py
# Fonction: archive_docx(docx_file, archive_folder='archive')
# Description: Archive un DOCX avec timestamp

# Python:
# from tools3.zip_docx import archive_docx
# result = archive_docx('document.docx', 'archive/')


# =============================================================================
# SECTION 6: FLUX DE TRAVAIL COMPLET (ÉTAPE PAR ÉTAPE)
# =============================================================================

# Scénario: Traiter un nouveau document du début à la fin

# Étape 1: Vérifier que le fichier existe
ls -la DC_SOURCES/DC_JNZ_2026.docx

# Étape 2: Exécuter le pipeline complet
python3 -m tools3.pipeline full -s DC_JNZ_2026.docx

# Étape 3: Vérifier les résultats
ls -la OUTPUT1_XML-RAW/
ls -la OUTPUT2_JSON-RAW/
ls -la OUTPUT3_JSON-TRANSFORMED/
ls -la OUTPUT4_DOCX-RESULT/

# Étape 4: Ouvrir le document final
open OUTPUT4_DOCX-RESULT/DC_JNZ_2026_GLOBAL_formatted.docx


# =============================================================================
# SECTION 7: WORKFLOW PERSONNALISÉ (DEBUGGING)
# =============================================================================

# Scénario: Vous voulez examiner le JSON RAW avant transformation

# Étape 1: Exécuter seulement Phase 1 (extraction)
python3 -m tools3.pipeline extract -s DC_JNZ_2026.docx

# Étape 2: Examiner le JSON RAW
cat OUTPUT2_JSON-RAW/DC_JNZ_2026_GLOBAL_raw.json | head -100

# Étape 3: Si satisfait, continuer Phase 2
python3 -m tools3.pipeline transform-render -s DC_JNZ_2026.docx


# =============================================================================
# SECTION 8: VARIABLES ET CHEMINS PAR DÉFAUT
# =============================================================================

# Chemin source par défaut
DC_SOURCES/

# Dossiers de sortie par défaut
OUTPUT1_XML-RAW/        # Phase 1: XML global
OUTPUT2_JSON-RAW/       # Phase 2: JSON brut
OUTPUT3_JSON-TRANSFORMED/  # Phase 3: JSON transformé
OUTPUT4_DOCX-RESULT/    # Phase 4: DOCX final

# Template par défaut
assets/TEMPLATE.docx


# =============================================================================
# SECTION 9: RÉSOLUTION DE PROBLÈMES COURANTS
# =============================================================================

# Erreur: "Fichier non trouvé"
# Solution: Assurez-vous que le fichier est dans DC_SOURCES/ ou utilisez chemin complet
python3 -m tools3.pipeline full -s DC_SOURCES/DC_JNZ_2026.docx

# Erreur: "Module not found: tools3"
# Solution: Assurez-vous d'être dans le répertoire /home/jbn/dc_formatter
cd /home/jbn/dc_formatter
python3 -m tools3.pipeline full -s DC_JNZ_2026.docx

# Erreur: "Transformation failed"
# Solution: Exécutez Phase 1 séparément pour vérifier XML/JSON RAW
python3 -m tools3.pipeline extract -s DC_JNZ_2026.docx
# Vérifiez les fichiers avant de continuer


# =============================================================================
# SECTION 10: BONNES PRATIQUES
# =============================================================================

# ✅ BON: Utiliser le pipeline complet (plus simple)
python3 -m tools3.pipeline full -s document.docx

# ✅ BON: Utiliser chemins relatifs si fichier dans DC_SOURCES/
python3 -m tools3.pipeline full -s document.docx

# ✅ BON: Spécifier output pour éviter mélange avec autres résultats
python3 -m tools3.pipeline full -s document.docx -o results/project1/

# ❌ MAUVAIS: Mélanger phases sans vérifier intermédiaires
# (Difficile de déboguer si erreur)

# ❌ MAUVAIS: Oublier extension .docx
# Utilisez: -s DC_JNZ_2026.docx (pas DC_JNZ_2026)


# =============================================================================
# SECTION 11: EXEMPLES AVANCÉS
# =============================================================================

# Traiter plusieurs documents en boucle
for doc in DC_SOURCES/*.docx; do
    echo "Traitement: $doc"
    python3 -m tools3.pipeline full -s "$doc" -o "results/$(basename "$doc" .docx)/"
done

# Extraire dimensions du template et sauvegarder dans fichier
python3 -m tools3.pipeline extract-dims > template_dimensions.txt

# Pipeline avec vérification d'erreur
python3 -m tools3.pipeline full -s document.docx && \
    echo "✓ Pipeline réussi" || \
    echo "✗ Pipeline échoué"


# =============================================================================
# SECTION 12: VERSIONS ET INFOS
# =============================================================================

# Structure des répertoires attendue:
# /home/jbn/dc_formatter/
# ├── README.md
# ├── requirements.txt
# ├── assets/
# │   └── TEMPLATE.docx
# ├── DC_SOURCES/
# │   ├── DC_JNZ_2026.docx
# │   └── DC_BM2.docx
# ├── tools3/
# │   ├── __init__.py
# │   ├── pipeline.py
# │   ├── extract_xml_raw.py
# │   ├── parse_template.py
# │   ├── parse_xml_raw_to_json_raw.py
# │   ├── process_json_raw_to_json_transformed.py
# │   ├── render_json_transformed_to_docx.py
# │   └── zip_docx.py
# ├── OUTPUT1_XML-RAW/
# ├── OUTPUT2_JSON-RAW/
# ├── OUTPUT3_JSON-TRANSFORMED/
# └── OUTPUT4_DOCX-RESULT/

# Version: 1.0
# Dernière mise à jour: Avril 2026
# Python: 3.8+
# Dépendances: python-docx, lxml, (voir requirements.txt)

