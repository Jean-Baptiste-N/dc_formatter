#!/usr/bin/env python3
import zipfile
import xml.etree.ElementTree as ET

docx = 'assets/TEMPLATE_DC_new.docx'

with zipfile.ZipFile(docx, 'r') as z:
    styles_xml = z.read('word/styles.xml')
    root = ET.fromstring(styles_xml)
    
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    custom_styles = []
    for style in root.findall('.//w:style', ns):
        style_id = style.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId')
        custom = style.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}customStyle')
        name_elem = style.find('w:name', ns)
        name = name_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if name_elem is not None else style_id
        
        if custom == '1' or (name and name.startswith('DC_')) or (style_id and style_id.startswith('DC_')):
            custom_styles.append((style_id, name, 'customStyle=1' if custom == '1' else 'DC_'))
    
    print(f"\nStyles CUSTOM ou DC_ détectés ({len(custom_styles)}):\n")
    for sid, name, reason in sorted(custom_styles):
        print(f"  {sid:35} | {name:35} | {reason}")
