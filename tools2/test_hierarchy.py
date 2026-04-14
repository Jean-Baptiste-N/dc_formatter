#!/usr/bin/env python
"""
Script de test du détecteur de hiérarchie sur tous les DC.
Analyse et génère un rapport de couverture.
"""

import sys
from pathlib import Path
from tabulate import tabulate  # pip install tabulate

sys.path.insert(0, str(Path(__file__).parent.parent / "tools"))

from tools2.hierarchy_detector import HierarchyDetector


def test_document(doc_path: str) -> dict:
    """Teste un document et retourne les résultats."""
    if not Path(doc_path).exists():
        return {
            'file': Path(doc_path).name,
            'status': '❌ NOT FOUND',
            'h1': 0,
            'h2': 0,
            'total': 0,
            'error': f'File not found: {doc_path}'
        }
    
    try:
        detector = HierarchyDetector(doc_path)
        detections = detector.detect_all()
        
        h1_count = sum(1 for _, level, _ in detections if level == "Heading1")
        h2_count = sum(1 for _, level, _ in detections if level == "Heading2")
        
        return {
            'file': Path(doc_path).name,
            'status': '✅ OK' if detections else '⚠️  AUCUN',
            'h1': h1_count,
            'h2': h2_count,
            'total': len(detections),
            'error': None
        }
    
    except Exception as e:
        return {
            'file': Path(doc_path).name,
            'status': '❌ ERROR',
            'h1': 0,
            'h2': 0,
            'total': 0,
            'error': str(e)[:50]
        }


def main():
    """Lance les tests sur tous les DC availables."""
    
    test_dir = Path("/home/jbn/dc_formatter/test")
    
    # Lister les documents reformatés
    reformatted_docs = sorted(test_dir.glob("*_reformated.docx"))
    
    print("\n" + "="*80)
    print("TEST DE DÉTECTION DE HIÉRARCHIE")
    print("="*80 + "\n")
    
    if not reformatted_docs:
        print("❌ Aucun document reformaté trouvé dans test/")
        return
    
    results = []
    
    for doc_path in reformatted_docs:
        print(f"🔍 Test: {doc_path.name}...")
        result = test_document(str(doc_path))
        results.append(result)
        
        if result['error']:
            print(f"   ⚠️  {result['error']}\n")
        else:
            print(f"   ✅ H1: {result['h1']}, H2: {result['h2']}, Total: {result['total']}\n")
    
    # Afficher le tableau récapitulatif
    print("\n" + "="*80)
    print("RÉSUMÉ DES TESTS")
    print("="*80 + "\n")
    
    table_data = [[
        r['file'][:30],
        r['status'],
        r['h1'],
        r['h2'],
        r['total']
    ] for r in results]
    
    headers = ['Document', 'Status', 'H1', 'H2', 'Total']
    print(tabulate(table_data, headers=headers, tablefmt='grid'))
    
    # Statistiques globales
    total_h1 = sum(r['h1'] for r in results if r['status'] == '✅ OK')
    total_h2 = sum(r['h2'] for r in results if r['status'] == '✅ OK')
    total_docs = len([r for r in results if r['status'] == '✅ OK'])
    
    print(f"\n📊 STATISTIQUES GLOBALES:")
    print(f"   Documents testés:     {total_docs}")
    print(f"   Total Heading1:       {total_h1}")
    print(f"   Total Heading2:       {total_h2}")
    print(f"   Total hiérarchies:    {total_h1 + total_h2}\n")


if __name__ == "__main__":
    main()
