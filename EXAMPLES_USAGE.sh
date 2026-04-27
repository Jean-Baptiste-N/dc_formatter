#!/bin/bash
# Exemples d'utilisation du pipeline refactorisé

# Structure des résultats:
# - pipeline-1_XML-RAW/        → XML global
# - pipeline-2_JSON-RAW/       → JSON RAW
# - pipeline-3_JSON-TRANSFORMED/ → JSON transformé
# - pipeline-4_DOCX-RESULT/    → DOCX final

# Template utilisé par défaut: assets/TEMPLATE.docx

# ============================================================================
# EXEMPLE 1: Pipeline Complète (Une Seule Commande)
# ============================================================================
echo "📊 EXEMPLE 1: Pipeline Complète"
python -m tools3.pipeline full test/BM2/input.docx


# ============================================================================
# EXEMPLE 2: Pipeline Complète avec Dossier Custom
# ============================================================================
echo "📊 EXEMPLE 2: Pipeline Complète avec Dossier Custom"
python -m tools3.pipeline full test/BM2/input.docx -o results/bm2/


# ============================================================================
# EXEMPLE 3: Phase 1 - Extraction (dims + xml + json raw)
# ============================================================================
echo "📦 EXEMPLE 3: Phase 1 - Extraction"
python -m tools3.pipeline extract test/BM2/input.docx


# ============================================================================
# EXEMPLE 4: Phase 2 - Transformation + Rendu (après Phase 1)
# ============================================================================
echo "🎨 EXEMPLE 4: Phase 2 - Transformation + Rendu"
python -m tools3.pipeline transform-render test/BM2/input.docx


# ============================================================================
# EXEMPLE 5: Extraction + Transformation séparées
# ============================================================================
echo "🔧 EXEMPLE 5: Extraction + Transformation séparées"

# Phase 1
python -m tools3.pipeline extract test/BM2/input.docx

# (Travail optionnel sur le JSON RAW)

# Phase 2
python -m tools3.pipeline transform-render test/BM2/input.docx


# ============================================================================
# EXEMPLE 6: Commandes Individuelles (avancé)
# ============================================================================
echo "🔧 EXEMPLE 6: Commandes Individuelles"

# 6.1 Extraire les dimensions du template (défaut: assets/TEMPLATE.docx)
python -m tools3.pipeline extract-dims

# 6.2 Extraire le XML brut
python -m tools3.pipeline extract-xml test/BM2/input.docx

# 6.3 Convertir XML en JSON RAW
python -m tools3.pipeline xml-to-json \
    pipeline-1_XML-RAW/BM2_GLOBAL.xml \
    pipeline-2_JSON-RAW/BM2_GLOBAL_raw.json

# 6.4 Transformer le JSON RAW
python -m tools3.pipeline transform \
    pipeline-2_JSON-RAW/BM2_GLOBAL_raw.json

# 6.5 Rendre le JSON en DOCX
python -m tools3.pipeline render \
    pipeline-3_JSON-TRANSFORMED/BM2_GLOBAL_transformed.json


# ============================================================================
# EXEMPLE 7: Traiter Plusieurs Documents (Efficace)
# ============================================================================
echo "📁 EXEMPLE 7: Traiter Plusieurs Documents"

# Phase 1: Extraction pour tous (crée les JSON RAW une fois)
for doc in test/BM2 test/CBD test/CBD24; do
    echo "Extraction de $doc..."
    python -m tools3.pipeline extract "$doc/input.docx"
done

# Phase 2: Transformation et rendu (utilise les JSON RAW)
for doc in test/BM2 test/CBD test/CBD24; do
    echo "Transformation et rendu de $doc..."
    python -m tools3.pipeline transform-render "$doc/input.docx"
done

# Tous les résultats sont dans:
# - pipeline-4_DOCX-RESULT/
# contient les DOCX générés pour BM2, CBD, CBD24


# ============================================================================
# EXEMPLE 8: Affichage de l'aide
# ============================================================================
echo "❓ EXEMPLE 8: Affichage de l'aide"

# Aide générale
python -m tools3.pipeline --help

# Aide pour une commande spécifique
python -m tools3.pipeline full --help
python -m tools3.pipeline extract --help
python -m tools3.pipeline transform-render --help
python -m tools3.pipeline extract-xml --help


# ============================================================================
# NOTES IMPORTANTES
# ============================================================================
# 1. RÉSULTATS:
#    Chaque étape crée un dossier distinct et prévisible
#
# 2. TEMPLATE:
#    Utilise toujours assets/TEMPLATE.docx (pas besoin de le spécifier)
#
# 3. DOSSIERS PAR DÉFAUT:
#    Les 4 dossiers sont créés à la racine par défaut
#    Utilisez -o pour changer le dossier parent
#
# 4. PIPELINE RECOMMANDÉE:
#    Pour traiter rapidement: python -m tools3.pipeline full document.docx
#    Pour plus de contrôle: extract → (vérifier JSON RAW) → transform-render
#
# 5. RÉUTILISATION:
#    Phase 1 crée le JSON RAW qui peut être:
#    - Modifié manuellement
#    - Traité avec transform-render
#    - Réutilisé pour plusieurs documents

