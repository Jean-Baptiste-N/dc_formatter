#!/usr/bin/env python3
"""
Workflow complet: Détecte les styles → Génère le code → Crée un template

Usage:
    python analyze_and_generate_styles.py monfichierfichier.docx
"""

import sys
from pathlib import Path
from detect_custom_styles import detect_styles
from tools2.generate_style_code import generate_from_json


def analyze_and_generate(docx_path: str):
    """Pipeline complet: détection → génération → template."""
    
    docx_file = Path(docx_path)
    
    if not docx_file.exists():
        print(f"❌ Fichier non trouvé: {docx_path}")
        sys.exit(1)
    
    print("\n" + "="*70)
    print("🔧 WORKFLOW COMPLET D'ANALYSE ET GÉNÉRATION DE STYLES")
    print("="*70)
    
    # Étape 1: Détection
    print("\n[1/3] 🔍 DÉTECTION DES STYLES...")
    detector = detect_styles(str(docx_file))
    json_file = str(docx_file.stem) + "_styles.json"
    
    # Étape 2: Génération du code
    print("\n[2/3] 📝 GÉNÉRATION DU CODE...")
    output_script = str(docx_file.stem) + "_create_template.py"
    output_guide = str(docx_file.stem) + "_styles_guide.md"
    generate_from_json(json_file, output_script, output_guide)
    
    # Étape 3: Créer le template
    print("\n[3/3] 🎨 CRÉATION DU TEMPLATE DOCX...")
    try:
        # Charger le module généré dynamiquement
        import importlib.util
        spec = importlib.util.spec_from_file_location("generated_styles", output_script)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        
        template_path = str(docx_file.stem) + "_TEMPLATE.docx"
        module.create_styles_template(template_path)
        print(f"✓ Template créé: {template_path}")
    except Exception as e:
        print(f"⚠️  Le template DOCX n'a pas pu être créé automatiquement")
        print(f"   Mais les fichiers Python et Markdown sont générés")
        print(f"   Erreur: {e}")
    
    # Résumé final
    print("\n" + "="*70)
    print("✓ WORKFLOW TERMINÉ!")
    print("="*70)
    print(f"\n📁 Fichiers générés:")
    print(f"   1. {json_file} - Données JSON des styles")
    print(f"   2. {output_script} - Script Python pour recréer les styles")
    print(f"   3. {output_guide} - Guide de référence Markdown")
    
    print(f"\n💡 Prochaines étapes:")
    print(f"   • Exécuter le script: python {output_script}")
    print(f"   • Consulter le guide: {output_guide}")
    print(f"   • Adapter selon vos besoins spécifiques")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python analyze_and_generate_styles.py <fichier.docx>")
        print("\nExemple:")
        print("   python analyze_and_generate_styles.py test/BMO/DC_BMO.docx")
        sys.exit(1)
    
    analyze_and_generate(sys.argv[1])
