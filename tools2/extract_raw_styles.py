#!/usr/bin/env python3
import zipfile
import xml.etree.ElementTree as ET
import json
from pathlib import Path

docx_path = 'assets/TEMPLATE_DC_new.docx'

# Extraire styles.xml
with zipfile.ZipFile(docx_path, 'r') as z:
    styles_xml = z.read('word/styles.xml')

root = ET.fromstring(styles_xml)

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# Chercher tous les styles DC_
dc_styles_raw = {}

for style in root.findall('.//w:style', ns):
    style_id = style.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId')
    name_elem = style.find('w:name', ns)
    name = name_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if name_elem is not None else style_id
    
    # Filtrer sur DC_
    if name and name.startswith('DC_'):
        # Conserver le XML brut de ce style
        dc_styles_raw[style_id] = {
            'id': style_id,
            'name': name,
            'xml_raw': ET.tostring(style, encoding='unicode')
        }

print(f"✓ {len(dc_styles_raw)} styles DC_ extraits")

# Exporter en JSON
output_file = Path('styles/DC_styles_RAW.json')
with open(output_file, 'w', encoding='utf-8') as f:
    json.dump(dc_styles_raw, f, indent=2, ensure_ascii=False)

print(f"✓ Fichier généré: {output_file}")
print("\nExemple DC_Ligne:")
for sid, info in dc_styles_raw.items():
    if 'Ligne' in info['name']:
        print(info['xml_raw'])
        break
