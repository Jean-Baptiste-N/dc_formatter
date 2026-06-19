#!/usr/bin/env python3
"""
Post-processor pour conserver les propriétés XML de numérotation/indentation.

Après que python-docx génère le DOCX, on injecte les attributs numPr et ilvl
directement dans le XML basé sur:
1. Les données du JSON transformé (qui conserve numId et ilvl originaux)
2. Les styles appliqués au paragraphe

Cela garantit que lors d'une 2ème passe de traitement, les niveaux d'indentation
et les références de numérotation sont correctement restaurés.
"""

import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, Optional

# Namespaces
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}

# Mapping de secours: Style → (numId, ilvl)
# Utilisé si le JSON transformé n'a pas les données
# NOTE: numId=1 est la numérotation principale. Les niveaux sont gérés par ilvl.
STYLE_TO_DEFAULT_NUMBERING = {
    'DC1stbullet': (1, 0),      # Premier niveau (ilvl=0)
    'DC2ndbullet': (1, 1),      # Deuxième niveau (ilvl=1)
    'DC3rdbullet': (1, 2),      # Troisième niveau (ilvl=2)
    'DC4thbullet': (1, 3),      # Quatrième niveau (ilvl=3)
}

def register_namespaces() -> None:
    """Register all namespaces to preserve them in XML output."""
    ET.register_namespace('w', NS['w'])
    ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    ET.register_namespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
    ET.register_namespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    ET.register_namespace('pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture')
    ET.register_namespace('wp14', 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing')
    ET.register_namespace('w14', 'http://schemas.microsoft.com/office/word/2010/wordml')
    ET.register_namespace('w15', 'http://schemas.microsoft.com/office/word/2012/wordml')


def load_json_numbering_data(json_path: Path) -> Dict[int, Dict[str, any]]:
    """
    Charge le JSON transformé et extrait les données de numérotation.
    
    Retourne un dict: para_index → {'style': ..., 'numId': ..., 'ilvl': ...}
    
    Note: On ne prend que l'ilvl du JSON. Le numId est déduit du style.
    """
    if not json_path.exists():
        return {}
    
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except:
        return {}
    
    numbering_data = {}
    para_idx = 0
    
    for elem in data.get('document', {}).get('content', []):
        if elem.get('type') == 'Paragraph':
            props = elem.get('properties', {})
            style = props.get('style')
            ilvl = props.get('ilvl')
            
            # Stocker les données si c'est un style bullet
            if style and any(s in style for s in ['bullet', '1st', '2nd', '3rd', '4th']):
                numbering_data[para_idx] = {
                    'style': style,
                    'ilvl': ilvl
                }
            
            para_idx += 1
        elif elem.get('type') == 'Table':
            # Tables n'affectent pas l'indexation des paragraphes du document principal
            pass
    
    return numbering_data


def inject_numbering_properties(docx_path: Path, json_path: Optional[Path] = None) -> int:
    """
    Inject numPr (numId, ilvl) into paragraphs based on applied styles.
    
    Si json_path est fourni, utilise les données du JSON transformé.
    Sinon, utilise les mappings de secours.
    
    Retourne le nombre de paragraphes avec numPr injectés.
    """
    register_namespaces()
    
    # Charger les données du JSON si disponible
    json_numbering = {}
    if json_path:
        json_numbering = load_json_numbering_data(json_path)
    
    # Extract and modify document.xml
    with zipfile.ZipFile(docx_path, 'r') as zip_read:
        with zip_read.open('word/document.xml') as f:
            doc_content = f.read()
    
    # Parse XML
    root = ET.fromstring(doc_content)
    
    # Find all paragraphs
    paragraphs = root.findall('.//w:p', NS)
    print(f"🔍 Processing {len(paragraphs)} paragraphs...")
    
    injected_count = 0
    para_idx = 0
    
    for para in paragraphs:
        # Get paragraph properties
        pPr = para.find('w:pPr', NS)
        if pPr is None:
            para_idx += 1
            continue
        
        # Get style
        pStyle = pPr.find('w:pStyle', NS)
        if pStyle is None:
            para_idx += 1
            continue
        
        style_val = pStyle.get('{%s}val' % NS['w'])
        if not style_val:
            para_idx += 1
            continue
        
        # Check if numPr already exists (preserve if present)
        existing_numPr = pPr.find('w:numPr', NS)
        if existing_numPr is not None:
            # Already has numbering, skip
            para_idx += 1
            continue
        
        # Determine numId and ilvl
        numId = None
        ilvl = None
        
        # Chercher d'abord dans les données JSON (pour ilvl)
        if para_idx in json_numbering:
            json_data = json_numbering[para_idx]
            ilvl = json_data.get('ilvl')
        
        # Déterminer numId et ilvl par défaut basés sur le style
        if style_val in STYLE_TO_DEFAULT_NUMBERING:
            default_numId, default_ilvl = STYLE_TO_DEFAULT_NUMBERING[style_val]
            numId = default_numId
            # Préférer ilvl du JSON, sinon utiliser le défaut du style
            if ilvl is None:
                ilvl = default_ilvl
        
        # Si on a un numId ET ilvl, ajouter numPr
        if numId is not None and ilvl is not None:
            numPr = ET.Element('{%s}numPr' % NS['w'])
            
            # Ajouter ilvl si disponible
            if ilvl is not None:
                ilvl_elem = ET.SubElement(numPr, '{%s}ilvl' % NS['w'])
                ilvl_elem.set('{%s}val' % NS['w'], str(ilvl))
            
            # Ajouter numId
            numId_elem = ET.SubElement(numPr, '{%s}numId' % NS['w'])
            numId_elem.set('{%s}val' % NS['w'], str(numId))
            
            # Insérer après pStyle
            style_index = list(pPr).index(pStyle)
            pPr.insert(style_index + 1, numPr)
            
            injected_count += 1
        
        para_idx += 1
    
    print(f"✅ Injected numbering properties into {injected_count} paragraphs")
    
    # Rewrite the DOCX file with modified document.xml
    temp_docx = docx_path.parent / (docx_path.stem + '_temp.docx')
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as zip_read:
            with zipfile.ZipFile(temp_docx, 'w', zipfile.ZIP_DEFLATED) as zip_write:
                for item in zip_read.infolist():
                    if item.filename == 'word/document.xml':
                        # Write modified document.xml
                        zip_write.writestr(item, ET.tostring(root, encoding='utf-8', xml_declaration=True))
                    else:
                        # Copy other files as-is
                        zip_write.writestr(item, zip_read.read(item.filename))
        
        # Replace original with modified
        docx_path.unlink()
        temp_docx.rename(docx_path)
        print(f"✅ Updated {docx_path}")
    except Exception as e:
        print(f"❌ Error: {e}")
        if temp_docx.exists():
            temp_docx.unlink()
        raise
    
    return injected_count


def main():
    import sys
    if len(sys.argv) < 2:
        print("Usage: python preserve_xml_numbering.py <docx_file> [json_file]")
        sys.exit(1)
    
    docx_path = Path(sys.argv[1])
    if not docx_path.exists():
        print(f"❌ File not found: {docx_path}")
        sys.exit(1)
    
    json_path = None
    if len(sys.argv) >= 3:
        json_path = Path(sys.argv[2])
        if not json_path.exists():
            print(f"⚠️  JSON file not found, using default mappings: {json_path}")
            json_path = None
    
    inject_numbering_properties(docx_path, json_path)


if __name__ == '__main__':
    main()
