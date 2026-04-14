#!/usr/bin/env python3
"""
Détecteur V2 - Amélioré
Détecte correctement les styles custom et offre les 2 formats.
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import json


class WordStyleDetectorV2:
    """Détecte les styles custom d'un document Word - Version 2 améliorée."""
    
    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'w15': 'http://schemas.microsoft.com/office/word/2012/wordml'
    }
    
    def __init__(self, docx_path: str):
        self.docx_path = Path(docx_path)
        self.styles_root = None
        self.all_styles = {}
        self.custom_styles = {}
        self.dc_styles = {}
        
    def extract_all_styles(self) -> Dict:
        """Extrait styles.xml."""
        try:
            with zipfile.ZipFile(self.docx_path, 'r') as z:
                styles_xml = z.read('word/styles.xml')
                self.styles_root = ET.fromstring(styles_xml)
                self._parse_styles()
        except Exception as e:
            print(f"❌ Erreur: {e}")
        return self.all_styles
    
    def _parse_styles(self):
        """Parse le XML des styles."""
        for style_elem in self.styles_root.findall('.//w:style', self.NAMESPACES):
            style_id = style_elem.get(f'{{{self.NAMESPACES["w"]}}}styleId')
            style_type = style_elem.get(f'{{{self.NAMESPACES["w"]}}}type')
            custom_attr = style_elem.get(f'{{{self.NAMESPACES["w"]}}}customStyle')
            
            if not style_id:
                continue
            
            name_elem = style_elem.find(f'{{{self.NAMESPACES["w"]}}}name')
            style_name = name_elem.get(f'{{{self.NAMESPACES["w"]}}}val') if name_elem is not None else style_id
            
            is_custom = custom_attr == '1'
            
            style_info = {
                'id': style_id,
                'name': style_name,
                'type': style_type,
                'customStyle': is_custom,
                'properties': self._extract_properties(style_elem),
                'xml_element': style_elem  # Garder l'élément pour export XML
            }
            
            self.all_styles[style_id] = style_info
            
            # Classer dans custom_styles si custom="1"
            if is_custom:
                self.custom_styles[style_id] = style_info
                
                # Classer dans dc_styles si contient DC_
                if style_name.startswith('DC_') or style_id.startswith('DC'):
                    self.dc_styles[style_id] = style_info
    
    def _extract_properties(self, style_elem: ET.Element) -> Dict:
        """Extrait les propriétés formatées."""
        props = {
            'based_on': None,
            'paragraph': {},
            'character': {}
        }
        
        based_on = style_elem.find(f'{{{self.NAMESPACES["w"]}}}basedOn')
        if based_on is not None:
            props['based_on'] = based_on.get(f'{{{self.NAMESPACES["w"]}}}val')
        
        pPr = style_elem.find(f'{{{self.NAMESPACES["w"]}}}pPr')
        if pPr is not None:
            props['paragraph'] = self._extract_paragraph_props(pPr)
        
        rPr = style_elem.find(f'{{{self.NAMESPACES["w"]}}}rPr')
        if rPr is not None:
            props['character'] = self._extract_character_props(rPr)
        
        return props
    
    def _extract_paragraph_props(self, pPr: ET.Element) -> Dict:
        """Extrait les propriétés de paragraphe."""
        props = {}
        
        jc = pPr.find(f'{{{self.NAMESPACES["w"]}}}jc')
        if jc is not None:
            props['alignment'] = jc.get(f'{{{self.NAMESPACES["w"]}}}val')
        
        spacing = pPr.find(f'{{{self.NAMESPACES["w"]}}}spacing')
        if spacing is not None:
            props['spacing_before'] = self._twip_to_pt(spacing.get(f'{{{self.NAMESPACES["w"]}}}before'))
            props['spacing_after'] = self._twip_to_pt(spacing.get(f'{{{self.NAMESPACES["w"]}}}after'))
        
        ind = pPr.find(f'{{{self.NAMESPACES["w"]}}}ind')
        if ind is not None:
            props['left_indent'] = self._twip_to_pt(ind.get(f'{{{self.NAMESPACES["w"]}}}left'))
        
        # Extraire les tabulations
        tabs_elem = pPr.find(f'{{{self.NAMESPACES["w"]}}}tabs')
        if tabs_elem is not None:
            tabs = []
            for tab in tabs_elem.findall(f'{{{self.NAMESPACES["w"]}}}tab'):
                tab_info = {
                    'val': tab.get(f'{{{self.NAMESPACES["w"]}}}val'),  # left, center, right, decimal
                    'leader': tab.get(f'{{{self.NAMESPACES["w"]}}}leader'),  # none, dot, hyphen, heavy, middleDot, underscore
                    'pos': self._twip_to_pt(tab.get(f'{{{self.NAMESPACES["w"]}}}pos'))  # position en twips
                }
                tabs.append(tab_info)
            if tabs:
                props['tabs'] = tabs
        
        return props
    
    def _extract_character_props(self, rPr: ET.Element) -> Dict:
        """Extrait les propriétés de caractère."""
        props = {}
        
        rFonts = rPr.find(f'{{{self.NAMESPACES["w"]}}}rFonts')
        if rFonts is not None:
            props['font_name'] = rFonts.get(f'{{{self.NAMESPACES["w"]}}}ascii')
        
        sz = rPr.find(f'{{{self.NAMESPACES["w"]}}}sz')
        if sz is not None:
            size_str = sz.get(f'{{{self.NAMESPACES["w"]}}}val')
            props['font_size'] = int(size_str) // 2 if size_str else None
        
        if rPr.find(f'{{{self.NAMESPACES["w"]}}}b') is not None:
            props['bold'] = True
        
        if rPr.find(f'{{{self.NAMESPACES["w"]}}}i') is not None:
            props['italic'] = True
        
        color = rPr.find(f'{{{self.NAMESPACES["w"]}}}color')
        if color is not None:
            color_val = color.get(f'{{{self.NAMESPACES["w"]}}}val')
            props['color_hex'] = color_val
        
        return props
    
    @staticmethod
    def _twip_to_pt(twip_str: Optional[str]) -> Optional[float]:
        if twip_str is None:
            return None
        try:
            return int(twip_str) / 20
        except ValueError:
            return None
    
    @staticmethod
    def _element_to_string(elem: ET.Element) -> str:
        """Convertit un élément XML en chaîne avec namespaces."""
        return ET.tostring(elem, encoding='unicode', method='xml')
    
    def export_to_json(self, output_path: str, styles_dict: Dict = None):
        """Exporte les styles en JSON avec propriétés structurées ET XML brut."""
        if styles_dict is None:
            styles_dict = self.dc_styles
        
        export_data = {}
        for sid, style_info in styles_dict.items():
            # Convertir l'élément XML en chaîne
            xml_raw = self._element_to_string(style_info['xml_element'])
            
            export_data[sid] = {
                'id': style_info['id'],
                'name': style_info['name'],
                'type': style_info['type'],
                'customStyle': style_info['customStyle'],
                'properties': style_info['properties'],
                'xml_raw': xml_raw
            }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(export_data, f, indent=2, ensure_ascii=False)
        
        print(f"✓ JSON exporté: {output_path}")
    
    @staticmethod
    def _indent(elem, level=0):
        """Ajoute l'indentation au XML pour le rendre lisible."""
        i = "\n" + level * "  "
        if len(elem):
            if not elem.text or not elem.text.strip():
                elem.text = i + "  "
            if not elem.tail or not elem.tail.strip():
                elem.tail = i
            for child in elem:
                WordStyleDetectorV2._indent(child, level + 1)
            if not child.tail or not child.tail.strip():
                child.tail = i
        else:
            if level and (not elem.tail or not elem.tail.strip()):
                elem.tail = i

    def export_to_xml(self, output_path: str, styles_dict: Dict = None):
        """Exporte les styles en XML brut (fragment réutilisable) avec indentation."""
        if styles_dict is None:
            styles_dict = self.dc_styles
        
        # Créer un document XML avec les styles
        root = ET.Element('w:styles', {
            'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'xmlns:w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
            'xmlns:w15': 'http://schemas.microsoft.com/office/word/2012/wordml'
        })
        
        # Ajouter chaque style
        for sid, style_info in sorted(styles_dict.items()):
            # Cloner l'élément original
            original_elem = style_info['xml_element']
            cloned_elem = ET.Element(f'{{{self.NAMESPACES["w"]}}}style')
            
            # Copier tous les attributs
            for key, value in original_elem.attrib.items():
                cloned_elem.set(key, value)
            
            # Copier tous les enfants
            for child in original_elem:
                cloned_elem.append(child)
            
            root.append(cloned_elem)
        
        # Ajouter l'indentation
        self._indent(root)
        
        # Formatter et sauvegarder
        tree = ET.ElementTree(root)
        tree.write(output_path, encoding='utf-8', xml_declaration=True)
        
        print(f"✓ XML exporté: {output_path}")


def analyze_and_export(docx_path: str, filter_type: str = 'dc'):
    """Analyse et exporte les styles."""
    
    docx_file = Path(docx_path)
    if not docx_file.exists():
        print(f"❌ Fichier non trouvé: {docx_path}")
        return
    
    print(f"\n{'='*70}")
    print(f"🔍 ANALYSE V2 - {docx_file.name}")
    print(f"{'='*70}")
    
    detector = WordStyleDetectorV2(str(docx_file))
    detector.extract_all_styles()
    
    # Afficher le résumé
    print(f"\n📊 Résumé:")
    print(f"   Total styles: {len(detector.all_styles)}")
    print(f"   Styles custom (customStyle=1): {len(detector.custom_styles)}")
    print(f"   Styles DC_*: {len(detector.dc_styles)}")
    
    # Déterminer quels styles exporter
    if filter_type == 'dc':
        export_styles = detector.dc_styles
        prefix = 'DC_'
    else:
        export_styles = detector.custom_styles
        prefix = 'custom_'
    
    print(f"\n🎨 Styles {prefix}à exporter ({len(export_styles)}):")
    for sid, style in sorted(export_styles.items(), key=lambda x: x[1]['name']):
        char_props = style['properties']['character']
        bold = " **" if char_props.get('bold') else ""
        size = f" {char_props.get('font_size')}pt" if char_props.get('font_size') else ""
        color = f" #{char_props.get('color_hex')}" if char_props.get('color_hex') else ""
        print(f"   • {style['name']:35} | {style['type']:12} |{bold}{size}{color}")
    
    # Exporter dans les 2 formats
    base_name = docx_file.stem + f'_{prefix}'
    json_path = base_name + 'styles.json'
    xml_path = base_name + 'styles.xml'
    
    detector.export_to_json(json_path, export_styles)
    detector.export_to_xml(xml_path, export_styles)
    
    print(f"\n✓ Fichiers générés:")
    print(f"  1. {json_path} - Facile à modifier/parser")
    print(f"  2. {xml_path} - XML réutilisable dans Word")


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python3 detect_styles_v2.py <docx> [dc|custom]")
        print("\nExemples:")
        print("  python3 detect_styles_v2.py assets/TEMPLATE_DC_new.docx dc")
        print("  python3 detect_styles_v2.py assets/TEMPLATE_DC_new.docx custom")
        sys.exit(1)
    
    docx_file = sys.argv[1]
    filter_type = sys.argv[2] if len(sys.argv) > 2 else 'dc'
    
    analyze_and_export(docx_file, filter_type)
