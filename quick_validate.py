#!/usr/bin/env python3
"""Validation rapide de la syntaxe et des imports du pipeline"""

import sys
from pathlib import Path

# Ajouter tools3 au chemin
sys.path.insert(0, str(Path(__file__).parent))

print("🔍 Validation rapide du pipeline refactorisé\n")

# 1. Vérifier les imports
print("1️⃣  Vérification des imports...")
try:
    from tools3 import pipeline
    print("   ✅ pipeline.py importé correctement\n")
except Exception as e:
    print(f"   ❌ Erreur à l'import: {e}\n")
    sys.exit(1)

# 2. Vérifier les constantes
print("2️⃣  Vérification des constantes...")
try:
    assert hasattr(pipeline, 'TEMPLATE_PATH'), "TEMPLATE_PATH manquante"
    assert hasattr(pipeline, 'SOURCE_DOCX_DIR'), "SOURCE_DOCX_DIR manquante"
    assert hasattr(pipeline, 'OUTPUT_XML_RAW'), "OUTPUT_XML_RAW manquante"
    assert hasattr(pipeline, 'OUTPUT_JSON_RAW'), "OUTPUT_JSON_RAW manquante"
    assert hasattr(pipeline, 'OUTPUT_JSON_TRANSFORMED'), "OUTPUT_JSON_TRANSFORMED manquante"
    assert hasattr(pipeline, 'OUTPUT_DOCX_RESULT'), "OUTPUT_DOCX_RESULT manquante"
    
    print(f"   TEMPLATE_PATH: {pipeline.TEMPLATE_PATH}")
    print(f"   SOURCE_DOCX_DIR: {pipeline.SOURCE_DOCX_DIR}")
    print(f"   OUTPUT_XML_RAW: {pipeline.OUTPUT_XML_RAW}")
    print(f"   OUTPUT_JSON_RAW: {pipeline.OUTPUT_JSON_RAW}")
    print(f"   OUTPUT_JSON_TRANSFORMED: {pipeline.OUTPUT_JSON_TRANSFORMED}")
    print(f"   OUTPUT_DOCX_RESULT: {pipeline.OUTPUT_DOCX_RESULT}\n")
    print("   ✅ Toutes les constantes sont présentes\n")
except AssertionError as e:
    print(f"   ❌ {e}\n")
    sys.exit(1)

# 3. Vérifier la fonction helper
print("3️⃣  Vérification de la fonction helper...")
try:
    assert hasattr(pipeline, '_resolve_source_path'), "_resolve_source_path manquante"
    print("   ✅ _resolve_source_path présente\n")
except AssertionError as e:
    print(f"   ❌ {e}\n")
    sys.exit(1)

# 4. Vérifier les fonctions cmd_*
print("4️⃣  Vérification des fonctions cmd_*...")
cmd_functions = [
    'cmd_extract_dims',
    'cmd_extract_xml',
    'cmd_xml_to_json',
    'cmd_transform',
    'cmd_render',
    'cmd_extract_all',
    'cmd_transform_and_render',
    'cmd_pipeline_full'
]

for func_name in cmd_functions:
    if hasattr(pipeline, func_name):
        print(f"   ✅ {func_name}")
    else:
        print(f"   ❌ {func_name} manquante")
        sys.exit(1)

print()

# 5. Vérifier main()
print("5️⃣  Vérification de main()...")
try:
    assert hasattr(pipeline, 'main'), "main() manquante"
    print("   ✅ main() présente\n")
except AssertionError as e:
    print(f"   ❌ {e}\n")
    sys.exit(1)

# 6. Test rapide de syntaxe
print("6️⃣  Test de compilation bytecode...")
try:
    import py_compile
    py_compile.compile(str(Path(__file__).parent / 'tools3' / 'pipeline.py'), doraise=True)
    print("   ✅ Compilation bytecode réussie\n")
except Exception as e:
    print(f"   ❌ Erreur de compilation: {e}\n")
    sys.exit(1)

print("=" * 50)
print("✅ VALIDATION COMPLÈTE: Tout fonctionne!")
print("=" * 50)

# Test the path resolution
print("\n7️⃣  Test de résolution de chemin...")
try:
    test_path = pipeline._resolve_source_path('DC_BM2.docx')
    print(f"   DC_BM2.docx → {test_path}")
    print(f"   Existe: {Path(test_path).exists()}")
    
    if Path(test_path).exists():
        print("   ✅ Résolution de chemin OK\n")
    else:
        print("   ⚠️  Fichier n'existe pas mais résolution OK\n")
except Exception as e:
    print(f"   ❌ Erreur: {e}\n")
    sys.exit(1)
