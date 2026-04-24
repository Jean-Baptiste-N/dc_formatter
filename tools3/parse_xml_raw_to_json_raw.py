"""
Script simplifié pour extraire un fichier XML global en JSON avec tous les détails.
Entrée: fichier _GLOBAL.xml
Sortie: fichier _GLOBAL_raw.json (xml brut traduit en json)
Sortie: fichier _GLOBAL_transformed.json (après taggings et transformations)
"""

from argparse import ArgumentParser
import argparse
import re
import xml.etree.ElementTree as ET
import json
import zipfile
from pathlib import Path
from typing import Dict, Any, List

# ===== NAMESPACES =====
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
}

def extract_run_properties(run, ns: Dict) -> Dict[str, Any]:
    """Extrait les propriétés d'un run (texte formaté)"""
    props = {}

    rPr = run.find('w:rPr', ns)
    if rPr is not None:
        # Bold - vérifier l'attribut val (0 ou false = désactivé)
        b = rPr.find('w:b', ns)
        if b is not None:
            val = b.get(f'{{{ns["w"]}}}val', '1')
            if val not in ('0', 'false'):
                props['bold'] = True

        # Italic - vérifier l'attribut val (0 ou false = désactivé)
        i = rPr.find('w:i', ns)
        if i is not None:
            val = i.get(f'{{{ns["w"]}}}val', '1')
            if val not in ('0', 'false'):
                props['italic'] = True

        # Taille
        sz = rPr.find('w:sz', ns)
        if sz is not None:
            props['size'] = sz.get(f'{{{ns["w"]}}}val')

        # Couleur
        color = rPr.find('w:color', ns)
        if color is not None:
            props['color'] = color.get(f'{{{ns["w"]}}}val')

        # Font
        rFonts = rPr.find('w:rFonts', ns)
        if rFonts is not None:
            props['font'] = rFonts.get(f'{{{ns["w"]}}}ascii')

    return props

def extract_paragraph_properties(paragraph, ns: Dict) -> Dict[str, Any]:
    """Extrait les propriétés d'un paragraphe"""
    props = {}

    pPr = paragraph.find('.//w:pPr', ns)
    if pPr is not None:
        # Style
        pStyle = pPr.find('w:pStyle', ns)
        if pStyle is not None:
            style = pStyle.get(f'{{{ns["w"]}}}val', 'Normal')
            props['style'] = style
        else:
            props['style'] = 'Normal'  # défaut

        # Justification
        jc = pPr.find('w:jc', ns)
        if jc is not None:
            alignment = jc.get(f'{{{ns["w"]}}}val', 'left')
            props['alignment'] = alignment
        else:
            props['alignment'] = 'left'  # défaut

        # Numérotation (listes)
        numPr = pPr.find('w:numPr', ns)
        if numPr is not None:
            ilvl_elem = numPr.find('w:ilvl', ns)
            numId_elem = numPr.find('w:numId', ns)
            if ilvl_elem is not None:
                props['ilvl'] = ilvl_elem.get(f'{{{ns["w"]}}}val', '0')
            if numId_elem is not None:
                props['numId'] = numId_elem.get(f'{{{ns["w"]}}}val', None)

        # Section break
        sectPr = pPr.find('w:sectPr', ns)
        if sectPr is not None:
            section_type = sectPr.find('w:type', ns)
            if section_type is not None:
                type_val = section_type.get(f'{{{ns["w"]}}}val')
                props['section_break'] = type_val  # nextPage, continuous, etc.
            else:
                props['section_break'] = 'nextPage'  # défaut

    # paraId
    para_id = paragraph.get(f'{{{NS["w14"]}}}paraId')
    if para_id:
        props['paraId'] = para_id

    return props

def extract_runs_from_paragraph(paragraph, ns: Dict) -> List[Dict[str, Any]]:
    """Extrait tous les runs d'un paragraphe"""
    runs = []

    for run in paragraph.findall('w:r', ns):
        run_obj = {}

        # Vérifier les sauts de page
        br_elem = run.find('w:br', ns)
        if br_elem is not None:
            br_type = br_elem.get(f'{{{ns["w"]}}}type')
            if br_type == 'page':
                run_obj['page_break'] = True
                runs.append(run_obj)
                continue

        # Propriétés du run
        run_props = extract_run_properties(run, ns)
        if run_props:
            run_obj['properties'] = run_props

        # Texte
        text_elem = run.find('w:t', ns)
        if text_elem is not None and text_elem.text:
            run_obj['text'] = text_elem.text

        if run_obj:
            runs.append(run_obj)

    return runs

def parse_paragraph(paragraph, ns: Dict, index: int) -> Dict[str, Any]:
    """Parse un paragraphe"""
    para_obj = {
        'index': index,
        'type': 'Paragraph'
    }

    # Propriétés
    para_props = extract_paragraph_properties(paragraph, ns)
    if para_props:
        para_obj['properties'] = para_props

    # Runs
    runs = extract_runs_from_paragraph(paragraph, ns)
    if runs:
        para_obj['runs'] = runs
    else:
        # Text simple si pas de runs
        texts = []
        for t in paragraph.findall('.//w:t', ns):
            if t.text:
                texts.append(t.text)
        if texts:
            para_obj['text'] = ''.join(texts)

    return para_obj

def extract_table_properties(table, ns: Dict) -> Dict[str, Any]:
    """Extrait les propriétés du tableau (largeur globale, etc.)"""
    tblPr = table.find('w:tblPr', ns)
    props = {}

    if tblPr is not None:
        # Largeur du tableau
        tblW = tblPr.find('w:tblW', ns)
        if tblW is not None:
            props['table_width'] = tblW.get(f'{{{ns["w"]}}}w')
            props['table_width_type'] = tblW.get(f'{{{ns["w"]}}}type', 'dxa')

    return props

def extract_cell_width(cell, ns: Dict) -> int:
    """Extrait la largeur d'une cellule en twips"""
    tcPr = cell.find('w:tcPr', ns)
    if tcPr is not None:
        tcW = tcPr.find('w:tcW', ns)
        if tcW is not None:
            width = tcW.get(f'{{{ns["w"]}}}w')
            if width:
                try:
                    return int(width)
                except ValueError:
                    return None
    return None

def extract_row_height(row, ns: Dict) -> int:
    """Extrait la hauteur d'une ligne"""
    trPr = row.find('w:trPr', ns)
    if trPr is not None:
        trHeight = trPr.find('w:trHeight', ns)
        if trHeight is not None:
            height = trHeight.get(f'{{{ns["w"]}}}val')
            if height:
                try:
                    return int(height)
                except ValueError:
                    return None
    return None

def normalize_paragraph_runs(para: Dict[str, Any]) -> Dict[str, Any]:
    """
    Normalise et fusionne les runs consécutifs d'un paragraphe.

    ⚠️ IMPORTANT: Cette fonction FUSIONNE les runs avec les mêmes propriétés
    en concaténant leurs textes, pour que:
    1. La détection de keywords fonctionne mieux
    2. L'application des styles soit plus cohérente
    3. La structure JSON soit plus propre

    Logique:
    - Fusionne les runs consécutifs avec EXACTEMENT les mêmes propriétés
    - Concatène les textes avec un espace si nécessaire
    - Préserve les runs avec propriétés différentes

    Exemple:
    - Input: [{"text": "dossier ", "properties": {}},
              {"text": "de ", "properties": {}},
              {"text": "compétences", "properties": {}}]
    - Output: [{"text": "dossier de compétences", "properties": {}}]

    Args:
        para: Paragraphe JSON avec runs

    Returns:
        Paragraphe normalisé avec runs fusionnés
    """
    if 'runs' not in para or not para['runs']:
        return para

    runs = para['runs']
    normalized_runs = []

    for run in runs:
        # Ignorer les runs sans texte
        if 'text' not in run or not run['text']:
            continue

        # Obtenir les propriétés du run (ou dict vide si aucune)
        run_props = json.dumps(run.get('properties', {}), sort_keys=True)

        # Vérifier s'il faut fusionner avec le dernier run
        if normalized_runs and 'text' in normalized_runs[-1]:
            last_run_props = json.dumps(normalized_runs[-1].get('properties', {}), sort_keys=True)

            if run_props == last_run_props:
                # Mêmes propriétés: fusionner les textes
                last_text = normalized_runs[-1]['text']
                new_text = run['text']

                # Fusionner intelligemment: ajouter un espace seulement si nécessaire
                if last_text and not last_text.endswith(' ') and not new_text.startswith(' '):
                    normalized_runs[-1]['text'] = last_text + ' ' + new_text
                else:
                    normalized_runs[-1]['text'] = last_text + new_text

                continue  # Pas besoin d'ajouter un nouveau run

        # Ajouter le run normalisé
        normalized_runs.append({
            'text': run['text'],
            'properties': run.get('properties', {})
        })

    # Remplacer les runs
    para['runs'] = normalized_runs

    return para

def parse_table(table, ns: Dict, index: int) -> Dict[str, Any]:
    """Parse un tableau avec extraction des dimensions"""
    table_obj = {
        'index': index,
        'type': 'Table',
        'rows': []
    }

    # Extraire les propriétés du tableau
    table_props = extract_table_properties(table, ns)
    if table_props:
        table_obj['properties'] = table_props

    rows = table.findall('w:tr', ns)
    table_obj['row_count'] = len(rows)

    for row_idx, row in enumerate(rows):
        row_obj = {
            'row_index': row_idx,
            'cells': []
        }

        # Extraire la hauteur de la ligne
        row_height = extract_row_height(row, ns)
        if row_height is not None:
            row_obj['height'] = row_height

        cells = row.findall('w:tc', ns)
        for col_idx, cell in enumerate(cells):
            cell_obj = {
                'col_index': col_idx,
                'paragraphs': []
            }

            # Extraire la largeur de la cellule
            cell_width = extract_cell_width(cell, ns)
            if cell_width is not None:
                cell_obj['width'] = cell_width

            for para_idx, para in enumerate(cell.findall('w:p', ns)):
                cell_obj['paragraphs'].append(
                    parse_paragraph(para, ns, para_idx)
                )

            row_obj['cells'].append(cell_obj)

        table_obj['rows'].append(row_obj)

    if rows:
        table_obj['col_count'] = len(rows[0].findall('w:tc', ns))

    return table_obj

def parse_global_xml(xml_file: str) -> Dict[str, Any]:
    """Parse le fichier global.xml et extrait tous les éléments"""
    tree = ET.parse(xml_file)
    root = tree.getroot()

    document_structure = {
        "document": {
            "type": "Document",
            "source": Path(xml_file).name,
            "content": []
        }
    }

    body = root.find('.//w:body', NS)
    if body is None:
        return document_structure

    content_index = 0

    # Traiter chaque élément du body
    for element in body:
        # Paragraphes
        if element.tag == f'{{{NS["w"]}}}p':
            para_obj = parse_paragraph(element, NS, content_index)
            # ⭐ NORMALISER LES RUNS du paragraphe
            para_obj = normalize_paragraph_runs(para_obj)
            document_structure["document"]["content"].append(para_obj)
            content_index += 1

        # Tableaux
        elif element.tag == f'{{{NS["w"]}}}tbl':
            table_obj = parse_table(element, NS, content_index)

            # ⭐ NORMALISER LES RUNS de tous les paragraphes dans la table
            for row in table_obj.get('rows', []):
                for cell in row.get('cells', []):
                    for para in cell.get('paragraphs', []):
                        normalize_paragraph_runs(para)

            document_structure["document"]["content"].append(table_obj)
            content_index += 1

    # Ajouter statistiques
    document_structure["document"]["stats"] = {
        "total_elements": len(document_structure["document"]["content"]),
        "paragraphs": len([e for e in document_structure["document"]["content"] if e["type"] == "Paragraph"]),
        "tables": len([e for e in document_structure["document"]["content"] if e["type"] == "Table"]),
    }

    return document_structure

def xml_to_json(xml_file: str, output_file: str = None) -> str:
    """
    Convertit un fichier global.xml en JSON RAW (sans tags ni styles).

    Args:
        xml_file (str): Chemin du fichier global.xml
        output_file (str): Chemin du fichier JSON (optionnel)

    Returns:
        str: Chemin du fichier créé
    """
    xml_path = Path(xml_file)

    # Générer le nom de sortie
    if output_file is None:
        # Créer le JSON au même répertoire que la source
        output_file = xml_path.parent / (xml_path.stem + "_raw.json")

    output_path = Path(output_file)

    # Parser le XML (sans tagging)
    structure = parse_global_xml(str(xml_path))

    # Sauvegarder en JSON RAW
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(structure, f, ensure_ascii=False, indent=2)

    print(f"✅ {output_path.name} créé (RAW)")

    return str(output_path)

def main():
    parser = argparse.ArgumentParser(description="Convertir un fichier XML en JSON RAW")
    parser.add_argument("xml_file", help="Chemin du fichier global.xml")
    parser.add_argument("-o", "--output", help="Chemin du fichier JSON de sortie")

    args = parser.parse_args()

    xml_to_json(args.xml_file, args.output)