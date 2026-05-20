#!/usr/bin/env python3
"""
Exemples concrets d'utilisation du parsing détaillé XML
"""

import json
from pathlib import Path

def example_1_extract_hierarchy():
    """Exemple 1: Extraire la hiérarchie avec niveaux de liste"""
    
    print("=" * 80)
    print("EXEMPLE 1: Extraire hiérarchie avec niveaux de liste")
    print("=" * 80)
    
    with open('structures/DC_JNZ_2026_RAW.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    content = data['document']['content']
    
    print("\n📋 Liste hiérarchique avec niveaux:\n")
    
    for elem in content:
        if elem['type'] == 'Paragraph':
            props = elem.get('properties', {})
            
            # Si c'est une liste
            if 'ilvl' in props:
                ilvl = int(props['ilvl'])
                numId = props.get('numId', '?')
                
                # Récupérer le texte
                text = ''.join(r['text'] for r in elem.get('runs', []))
                
                # Indentation selon le niveau
                indent = "  " * ilvl
                
                print(f"{indent}├─ [{ilvl}] numId:{numId} | {text[:50]}")


def example_2_extract_formatted_text():
    """Exemple 2: Extraire le texte avec son formatage"""
    
    print("\n" + "=" * 80)
    print("EXEMPLE 2: Extraire texte formaté avec propriétés")
    print("=" * 80)
    
    with open('structures/DC_JNZ_2026_RAW.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    content = data['document']['content']
    
    print("\n📝 Paragraphes avec formatage:\n")
    
    count = 0
    for elem in content[:30]:
        if elem['type'] == 'Paragraph':
            props = elem.get('properties', {})
            
            # Afficher les propriétés du paragraphe
            para_props_str = ""
            if props.get('style'):
                para_props_str += f"Style:{props['style']}"
            if props.get('alignment'):
                para_props_str += f" | Align:{props['alignment']}"
            
            # Traiter les runs
            print(f"\n[Para {elem['index']}] {para_props_str}")
            
            for run_idx, run in enumerate(elem.get('runs', [])):
                text = run['text']
                run_props = run.get('properties', {})
                
                # Déterminer le formatage
                fmt = []
                if run_props.get('bold'):
                    fmt.append('B')
                if run_props.get('italic'):
                    fmt.append('I')
                if run_props.get('color'):
                    fmt.append(f"#{run_props['color']}")
                
                fmt_str = f" [{','.join(fmt)}]" if fmt else ""
                print(f"  Run {run_idx}: '{text}' {fmt_str}")
                print(f"    Size: {run_props.get('size')} | Font: {run_props.get('font')}")
            
            count += 1
            if count >= 5:
                break


def example_3_extract_by_style():
    """Exemple 3: Extraire tous les paragraphes avec un style spécifique"""
    
    print("\n" + "=" * 80)
    print("EXEMPLE 3: Extraire par style/tag")
    print("=" * 80)
    
    with open('structures/DC_JNZ_2026_RAW.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    content = data['document']['content']
    
    # Récupérer tous les styles uniques
    styles = set()
    for elem in content:
        if elem['type'] == 'Paragraph':
            style = elem.get('properties', {}).get('style')
            if style:
                styles.add(style)
    
    print(f"\n📌 Styles trouvés dans le document ({len(styles)}):\n")
    for style in sorted(styles):
        count = sum(1 for e in content 
                   if e['type'] == 'Paragraph' 
                   and e.get('properties', {}).get('style') == style)
        print(f"  • {style}: {count} paragraphe(s)")
    
    # Exemple: Extraire les Heading2
    print(f"\n🔍 Extrait - Tous les 'Heading2':\n")
    for elem in content:
        if elem['type'] == 'Paragraph':
            if elem.get('properties', {}).get('style') == 'Heading2':
                text = ''.join(r['text'] for r in elem.get('runs', []))
                color = elem['runs'][0].get('properties', {}).get('color', 'N/A') if elem.get('runs') else 'N/A'
                print(f"  • {text}")
                print(f"    Color: {color}\n")


def example_4_filter_lists():
    """Exemple 4: Filtrer les listes par niveau"""
    
    print("\n" + "=" * 80)
    print("EXEMPLE 4: Filtrer listes par niveau (ilvl)")
    print("=" * 80)
    
    with open('structures/DC_JNZ_2026_RAW.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    content = data['document']['content']
    
    # Grouper par ilvl
    levels = {}
    for elem in content:
        if elem['type'] == 'Paragraph':
            ilvl = elem.get('properties', {}).get('ilvl')
            if ilvl:
                if ilvl not in levels:
                    levels[ilvl] = []
                text = ''.join(r['text'] for r in elem.get('runs', []))
                levels[ilvl].append(text)
    
    print(f"\n📊 Listes groupées par niveau:\n")
    for level in sorted(levels.keys(), key=lambda x: int(x)):
        print(f"  Niveau {level} ({len(levels[level])} éléments):")
        for item in levels[level][:3]:
            print(f"    • {item[:60]}")
        if len(levels[level]) > 3:
            print(f"    ... et {len(levels[level]) - 3} autres")


def example_5_extract_tables():
    """Exemple 5: Extraire les tableaux"""
    
    print("\n" + "=" * 80)
    print("EXEMPLE 5: Extraire et analyser les tableaux")
    print("=" * 80)
    
    with open('structures/DC_JNZ_2026_RAW.json', 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    content = data['document']['content']
    
    tables = [e for e in content if e['type'] == 'Table']
    
    print(f"\n📊 Nombre de tableaux: {len(tables)}\n")
    
    for table in tables[:2]:
        rows = table['row_count']
        cols = table.get('col_count', '?')
        
        print(f"Tableau [{table['index']}]: {rows} rows × {cols} cols")
        
        # Afficher la première ligne
        if table['rows']:
            first_row = table['rows'][0]
            print("  Première ligne:")
            for cell in first_row['cells'][:3]:
                if cell['paragraphs']:
                    para = cell['paragraphs'][0]
                    text = ''.join(r['text'] for r in para.get('runs', []))
                    print(f"    • {text[:40]}")
        print()


if __name__ == "__main__":
    # Changer de répertoire vers le dossier du projet
    import os
    os.chdir('/home/jbn/dc_formatter')
    
    # Lancer les exemples
    example_1_extract_hierarchy()
    example_2_extract_formatted_text()
    example_3_extract_by_style()
    example_4_filter_lists()
    example_5_extract_tables()
    
    print("\n" + "=" * 80)
    print("✅ Tous les exemples exécutés")
    print("=" * 80)
