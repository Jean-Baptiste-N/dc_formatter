"""
Script pour convertir le JSON structuré (paragraphes + tableaux) en dataframe pandas.
Permet l'exploration et l'analyse interactive du document.
"""

import pandas as pd
import json
from pathlib import Path
from typing import Dict, Tuple


def json_to_dataframe(json_file: str, flatten: bool = False) -> pd.DataFrame:
    """
    Convertit le JSON (paragraphes + tableaux) en un unique dataframe.
    
    Args:
        json_file (str): Chemin du fichier JSON
        flatten (bool): Si True, crée une vue aplatie (une ligne par cellule tableau)
        
    Returns:
        pd.DataFrame: Dataframe avec tous les éléments
    """
    
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    doc = data['document']
    rows = []
    
    for elem in doc['content']:
        if elem['type'] == 'Paragraph':
            rows.append({
                'id': elem['id'],
                'type': 'Paragraph',
                'style': elem['style'],
                'text': elem['text'],
                'table_id': None,
                'row': None,
                'col': None
            })
        
        elif elem['type'] == 'Table':
            if flatten:
                # Vue aplatie: une ligne par cellule
                for row_data in elem['rows']:
                    for cell_data in row_data['cells']:
                        rows.append({
                            'id': elem['id'],
                            'type': 'Table',
                            'style': None,
                            'text': cell_data['content'],
                            'table_id': elem['id'],
                            'row': row_data['row_index'],
                            'col': cell_data['col_index']
                        })
            else:
                # Vue structurée: une ligne par cellule de tableau
                for row_data in elem['rows']:
                    for cell_data in row_data['cells']:
                        rows.append({
                            'id': f"{elem['id']}_r{row_data['row_index']}_c{cell_data['col_index']}",
                            'type': 'Table',
                            'style': None,
                            'text': cell_data['content'],
                            'table_id': elem['id'],
                            'row': row_data['row_index'],
                            'col': cell_data['col_index']
                        })
    
    df = pd.DataFrame(rows)
    df['source'] = Path(json_file).name
    
    return df


def get_table_dataframe(json_file: str, table_idx: int) -> pd.DataFrame:
    """
    Extrait un tableau spécifique en tant que dataframe pivot (par index, pas par ID).
    
    Args:
        json_file (str): Chemin du fichier JSON
        table_idx (int): Index du tableau (0, 1, 2, ...)
        
    Returns:
        pd.DataFrame: Tableau pivotté
    """
    
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    doc = data['document']
    tables = [e for e in doc['content'] if e['type'] == 'Table']
    
    if table_idx >= len(tables) or table_idx < 0:
        raise ValueError(f"Index tableau {table_idx} invalide (total: {len(tables)})")
    
    table = tables[table_idx]
    
    # Créer une structure pour le pivot
    rows_data = []
    for row_data in table['rows']:
        row_dict = {}
        for cell_data in row_data['cells']:
            row_dict[f"Col_{cell_data['col_index']}"] = cell_data['content']
        rows_data.append(row_dict)
    
    return pd.DataFrame(rows_data)


def explore_json(json_file: str):
    """
    Explore interactivement le JSON en affichant des statistiques et aperçus.
    
    Args:
        json_file (str): Chemin du fichier JSON
    """
    
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    doc = data['document']
    
    print(f"\n{'='*70}")
    print(f"📊 EXPLORATION DATAFRAME - {doc['source']}")
    print(f"{'='*70}\n")
    
    # Charger en dataframe
    df = json_to_dataframe(json_file)
    
    # ===== STATISTIQUES GLOBALES =====
    print(f"📈 Statistiques globales:")
    print(f"   Lignes: {len(df)}")
    print(f"   Colonnes: {len(df.columns)}")
    print(f"   Types d'éléments:")
    print(df['type'].value_counts().to_string().replace('\n', '\n      '))
    
    # ===== APERÇU PARAGRAPHES =====
    print(f"\n📝 Paragraphes ({(df['type']=='Paragraph').sum()} lignes):")
    df_para = df[df['type'] == 'Paragraph'][['id', 'style', 'text']].head(10)
    for idx, row in df_para.iterrows():
        text_preview = row['text'][:60] if row['text'] else '[VIDE]'
        print(f"   [{row['id']}] {row['style']:15} | {text_preview}...")
    
    # ===== APERÇU TABLEAUX =====
    print(f"\n📊 Tableaux ({(df['type']=='Table').sum()} lignes de tableau):")
    df_tables = df[df['type'] == 'Table']
    if not df_tables.empty:
        with open(json_file, 'r', encoding='utf-8') as f:
            doc = json.load(f)['document']
        tables = [e for e in doc['content'] if e['type'] == 'Table']
        print(f"   Tableaux trouvés: {len(tables)}")
        
        for table_idx in range(min(3, len(tables))):  # Afficher premiers 3 tableaux
            print(f"\n   Table {table_idx}:")
            table_df = get_table_dataframe(json_file, table_idx)
            print(f"   {table_df.shape[0]} rows × {table_df.shape[1]} cols")
            print(table_df.to_string().replace('\n', '\n   ')[:300] + '...')
    
    # ===== COLONNES =====
    print(f"\n📋 Colonnes disponibles:")
    for col in df.columns:
        dtype = df[col].dtype
        non_null = df[col].notna().sum()
        print(f"   • {col:15} ({dtype}, {non_null} non-null)")
    
    print(f"\n{'='*70}\n")
    
    return df


if __name__ == "__main__":
    import sys
    
    # Fichier par défaut ou argument
    json_file = sys.argv[1] if len(sys.argv) > 1 else 'structures/DC_JNZ_2026_RAW.json'
    
    try:
        # Exploration interactive
        df = explore_json(json_file)
        
        # Sauvegarder le dataframe en CSV
        csv_file = Path(json_file).with_suffix('.csv')
        df.to_csv(csv_file, index=False, encoding='utf-8')
        print(f"✓ CSV généré: {csv_file}")
        
        # Afficher le dataframe complet si demandé
        if len(sys.argv) > 2 and sys.argv[2] == '--show-all':
            print(f"\n📋 DATAFRAME COMPLET:")
            print(df.to_string())
    
    except Exception as e:
        print(f"✗ Erreur: {e}")
        import traceback
        traceback.print_exc()