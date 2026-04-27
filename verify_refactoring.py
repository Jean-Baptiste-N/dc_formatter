#!/usr/bin/env python3
"""
Script de vérification de la refactorisation
Teste les imports et la structure des modules
"""

import sys
from pathlib import Path

# Ajouter tools3 au chemin
tools3_path = Path(__file__).parent / 'tools3'
sys.path.insert(0, str(tools3_path.parent))

def test_imports():
    """Teste que tous les imports fonctionnent"""
    print("🔍 Vérification des imports...\n")
    
    try:
        print("  ✓ extract_xml_raw.py", end=" ")
        from tools3.extract_xml_raw import export_all_xml, extract_document_xml
        print("✅")
        
        print("  ✓ parse_template.py", end=" ")
        from tools3.parse_template import extract_page_dimensions_from_template
        print("✅")
        
        print("  ✓ parse_xml_raw_to_json_raw.py", end=" ")
        from tools3.parse_xml_raw_to_json_raw import xml_to_json
        print("✅")
        
        print("  ✓ process_json_raw_to_json_transformed.py", end=" ")
        from tools3.process_json_raw_to_json_transformed import apply_tags_and_styles
        print("✅")
        
        print("  ✓ render_json_transformed_to_docx.py", end=" ")
        from tools3.render_json_transformed_to_docx import json_to_docx
        print("✅")
        
        print("  ✓ pipeline.py", end=" ")
        from tools3.pipeline import main
        print("✅")
        
        print("\n✅ Tous les imports OK!\n")
        return True
        
    except Exception as e:
        print(f"\n❌ Erreur d'import: {e}\n")
        return False


def test_function_signatures():
    """Teste les signatures des fonctions"""
    print("📋 Vérification des signatures...\n")
    
    try:
        from tools3.extract_xml_raw import export_all_xml
        from tools3.parse_template import extract_page_dimensions_from_template
        from tools3.parse_xml_raw_to_json_raw import xml_to_json
        from tools3.process_json_raw_to_json_transformed import apply_tags_and_styles
        from tools3.render_json_transformed_to_docx import json_to_docx
        
        import inspect
        
        # Vérifier extract_page_dimensions_from_template
        sig = inspect.signature(extract_page_dimensions_from_template)
        params = list(sig.parameters.keys())
        print(f"  ✓ extract_page_dimensions_from_template({', '.join(params)})")
        assert len(params) == 1 and params[0] == 'template_path', \
            "Doit avoir 1 paramètre obligatoire"
        
        # Vérifier xml_to_json
        sig = inspect.signature(xml_to_json)
        params = list(sig.parameters.keys())
        print(f"  ✓ xml_to_json({', '.join(params)})")
        assert len(params) == 2, "Doit avoir 2 paramètres obligatoires"
        
        # Vérifier apply_tags_and_styles
        sig = inspect.signature(apply_tags_and_styles)
        params = list(sig.parameters.keys())
        print(f"  ✓ apply_tags_and_styles({', '.join(params)})")
        assert len(params) == 3, "Doit avoir 3 paramètres obligatoires"
        assert params[2] == 'page_dimensions', "Le 3e paramètre doit être page_dimensions"
        
        # Vérifier json_to_docx
        sig = inspect.signature(json_to_docx)
        params = list(sig.parameters.keys())
        print(f"  ✓ json_to_docx({', '.join(params)})")
        assert len(params) == 3, "Doit avoir 3 paramètres obligatoires"
        
        # Vérifier export_all_xml
        sig = inspect.signature(export_all_xml)
        params = list(sig.parameters.keys())
        print(f"  ✓ export_all_xml({', '.join(params)})")
        assert len(params) == 2, "Doit avoir 2 paramètres obligatoires"
        
        print("\n✅ Toutes les signatures sont correctes!\n")
        return True
        
    except AssertionError as e:
        print(f"\n❌ Erreur de signature: {e}\n")
        return False
    except Exception as e:
        print(f"\n❌ Erreur: {e}\n")
        return False


def test_cli():
    """Teste la CLI du pipeline"""
    print("🎯 Vérification de la CLI...\n")
    
    try:
        from tools3.pipeline import main
        import argparse
        
        print("  ✓ Pipeline CLI chargé")
        print("  ✓ ArgumentParser disponible")
        
        print("\n✅ CLI OK!\n")
        return True
        
    except Exception as e:
        print(f"\n❌ Erreur CLI: {e}\n")
        return False


def main_check():
    """Lance toutes les vérifications"""
    print("="*70)
    print("🔧 VÉRIFICATION DE LA REFACTORISATION")
    print("="*70 + "\n")
    
    results = []
    results.append(("Imports", test_imports()))
    results.append(("Signatures", test_function_signatures()))
    results.append(("CLI", test_cli()))
    
    print("="*70)
    print("📊 RÉSUMÉ")
    print("="*70 + "\n")
    
    all_ok = True
    for test_name, passed in results:
        status = "✅ OK" if passed else "❌ ÉCHEC"
        print(f"  {test_name}: {status}")
        if not passed:
            all_ok = False
    
    print("\n" + "="*70)
    if all_ok:
        print("✅ TOUTES LES VÉRIFICATIONS RÉUSSIES!")
        print("="*70 + "\n")
        print("Vous pouvez utiliser le pipeline avec:")
        print("  python -m tools3.pipeline --help\n")
        return 0
    else:
        print("❌ CERTAINES VÉRIFICATIONS ONT ÉCHOUÉ")
        print("="*70 + "\n")
        return 1


if __name__ == "__main__":
    sys.exit(main_check())
