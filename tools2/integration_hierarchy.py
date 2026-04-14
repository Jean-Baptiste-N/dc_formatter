#!/usr/bin/env python
"""
Intégrateur complet de reformatage avec détection de hiérarchie.
Flux: Document original → Reformatage → Détection de hiérarchie → Application de styles
"""

import sys
from pathlib import Path
from docx import Document

# Import local modules
sys.path.insert(0, str(Path(__file__).parent))
from tools2.hierarchy_detector import HierarchyDetector


def reformat_and_apply_styles(input_path: str, output_path: str, 
                              template_path: str = None, 
                              verbose: bool = True) -> dict:
    """
    Pipeline complet:
    1. Charge le document reformaté (ou original)
    2. Détecte la hiérarchie
    3. Applique les styles du template (optionnel)
    4. Sauvegarde le résultat
    
    Args:
        input_path: Document source (reformaté)
        output_path: Fichier de sortie
        template_path: Chemin du template avec styles (optionnel)
        verbose: Afficher les détails
    
    Returns:
        Dict avec stats: {
            'headings_detected': int,
            'heading1_count': int,
            'heading2_count': int,
            'applied': bool
        }
    """
    
    if verbose:
        print(f"\n{'='*70}")
        print(f"PIPELINE REFORMATAGE + HIÉRARCHIE")
        print(f"{'='*70}\n")
        print(f"📄 Document source: {input_path}")
        print(f"📦 Template:        {template_path or 'Aucun'}")
        print(f"💾 Sortie:          {output_path}\n")
    
    # Étape 1: Charger le document
    try:
        doc = Document(input_path)
    except Exception as e:
        print(f"❌ Erreur chargement: {e}")
        return {'error': str(e)}
    
    # Étape 2: Charger le template si fourni
    if template_path and Path(template_path).exists():
        try:
            template_doc = Document(template_path)
            if verbose:
                print(f"✅ Template chargé desde {template_path}")
            # Les styles du template seront accessibles
        except Exception as e:
            print(f"⚠️  Impossible charger template: {e}")
            template_doc = None
    else:
        template_doc = None
    
    # Étape 3: Détecter la hiérarchie
    if verbose:
        print(f"\n🔍 Analyse de la hiérarchie...")
    
    # Créer détecteur sur le document chargé
    detector = HierarchyDetector(input_path)
    detections = detector.detect_all()
    
    if verbose:
        detector.print_analysis(limit=15)
        detector.report()
    
    # Compter les types
    heading1_count = sum(1 for _, level, _ in detections if level == "Heading1")
    heading2_count = sum(1 for _, level, _ in detections if level == "Heading2")
    
    # Étape 4: Appliquer les styles
    if verbose:
        print(f"\n🎨 Application des styles Heading1/Heading2...")
    
    applied = detector.apply_all_detected()
    
    if verbose:
        print(f"✅ {applied} styles appliqués")
        print(f"   - {heading1_count} × Heading1")
        print(f"   - {heading2_count} × Heading2")
    
    # Étape 5: Sauvegarder
    try:
        detector.save(output_path)
        if verbose:
            print(f"\n✅ Document sauvegardé: {output_path}\n")
    except Exception as e:
        print(f"❌ Erreur sauvegarde: {e}")
        return {'error': str(e)}
    
    return {
        'success': True,
        'headings_detected': len(detections),
        'heading1_count': heading1_count,
        'heading2_count': heading2_count,
        'applied': applied,
        'input': input_path,
        'output': output_path,
        'template_used': bool(template_doc)
    }


def main():
    """CLI interface."""
    if len(sys.argv) < 2:
        print("""
Utilisation:
  python integration_hierarchy.py <input.docx> [output.docx] [--template template.docx]

Exemples:
  # Reformater + appliquer styles Heading
  python integration_hierarchy.py test/DC_JNZ_2026_reformated.docx reformatted_with_headings.docx
  
  # Avec template personnalisé
  python integration_hierarchy.py test/DC_JNZ_2026_reformated.docx output.docx --template TEMPLATE_DC.docx

Options:
  --template PATH    Utiliser un template avec styles personnalisés
  --quiet            Pas de verbose
        """)
        sys.exit(1)
    
    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else input_path.replace(".docx", "_with_headings.docx")
    
    # Vérifier les args optionnels
    template_path = None
    verbose = True
    
    for i, arg in enumerate(sys.argv[3:]):
        if arg == "--template" and i+4 < len(sys.argv):
            template_path = sys.argv[i+4]
        elif arg == "--quiet":
            verbose = False
    
    # Exécuter le pipeline
    result = reformat_and_apply_styles(input_path, output_path, template_path, verbose)
    
    if result.get('success'):
        print(f"📊 RÉSUMÉ:")
        print(f"   Headings détectés: {result['headings_detected']}")
        print(f"   H1: {result['heading1_count']}, H2: {result['heading2_count']}")


if __name__ == "__main__":
    main()
