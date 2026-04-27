#!/bin/bash
# QUICK START - Points d'Entrée Principaux

echo "🚀 QUICK START - PIPELINE REFACTORISÉ"
echo "======================================"
echo ""

# Test que le pipeline est prêt
echo "1️⃣  Vérification de l'installation..."
python3 -c "from tools3 import pipeline; print('✅ Pipeline importé avec succès')" || {
    echo "❌ Erreur d'importation"
    exit 1
}

echo ""
echo "2️⃣  Points d'entrée principaux:"
echo ""
echo "   📚 DOCUMENTATION:"
echo "      • FINAL_REFACTORING_SUMMARY.md  → Résumé complet des changements"
echo "      • EXAMPLES_USAGE.sh             → 8 exemples pratiques"
echo "      • ARCHITECTURE.md               → Diagrammes et flux"
echo ""

echo "   🧪 VALIDATION:"
echo "      • python3 quick_validate.py     → Vérifie la syntaxe et les imports"
echo ""

echo "   🎯 UTILISATION:"
echo "      • python3 -m tools3.pipeline --help"
echo "      • python3 -m tools3.pipeline full document.docx"
echo "      • python3 -m tools3.pipeline extract document.docx"
echo "      • python3 -m tools3.pipeline transform-render document.docx"
echo ""

echo "3️⃣  Structure des résultats:"
echo ""
echo "      pipeline-1_XML-RAW/           → XML extrait"
echo "      pipeline-2_JSON-RAW/          → JSON brut"
echo "      pipeline-3_JSON-TRANSFORMED/  → JSON transformé"
echo "      pipeline-4_DOCX-RESULT/       → DOCX final"
echo ""

echo "4️⃣  Configuration implicite:"
echo ""
echo "      • Template: assets/TEMPLATE.docx (modifiable via TEMPLATE_PATH)"
echo "      • Dimensions: Extraites une seule fois et réutilisées"
echo "      • Aucun paramètre par défaut caché"
echo ""

echo "✅ PRÊT À L'EMPLOI!"
echo ""
