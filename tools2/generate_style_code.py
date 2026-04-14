#!/usr/bin/env python3
"""
Génère du code Python (python-docx) à partir des styles détectés.
Crée:
1. Un script qui recréé tous les styles
2. Un template DOCX avec les styles appliqués
3. Un guide d'utilisation
"""

import json
from pathlib import Path
from typing import Dict, Any, List


class StyleCodeGenerator:
    """Génère du code Python pour recréer les styles Word."""
    
    def __init__(self, styles_json: str):
        """Initialise avec un fichier JSON de styles."""
        self.styles_file = Path(styles_json)
        self.styles: Dict = {}
        self._load_styles()
    
    def _load_styles(self):
        """Charge les styles depuis le JSON."""
        with open(self.styles_file, 'r', encoding='utf-8') as f:
            self.styles = json.load(f)
    
    def generate_full_script(self) -> str:
        """Génère un script Python complet pour recréer les styles."""
        
        script = '''#!/usr/bin/env python3
"""
Script généré automatiquement pour recréer les styles personnalisés.
Crée un template DOCX avec tous les styles détectés du document source.
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from pathlib import Path


def create_styles_template(output_path: str = "TEMPLATE_STYLES.docx") -> Document:
    """Crée un template avec les styles personnalisés."""
    
    doc = Document()
    
    print("\\n=== CRÉATION DU TEMPLATE DE STYLES ===\\n")
'''
        
        # Séparer les styles par type
        paragraph_styles = {}
        character_styles = {}
        
        for style_id, style_info in self.styles.items():
            if style_info['type'] == 'paragraph':
                paragraph_styles[style_id] = style_info
            elif style_info['type'] == 'character':
                character_styles[style_id] = style_info
        
        # Générer le code pour les styles de paragraphe
        if paragraph_styles:
            script += "\n    # ========== STYLES PARAGRAPHE ==========\n"
            for style_id, style_info in paragraph_styles.items():
                script += self._generate_paragraph_style_code(style_id, style_info)
        
        # Générer le code pour les styles de caractère
        if character_styles:
            script += "\n    # ========== STYLES CARACTÈRE ==========\n"
            for style_id, style_info in character_styles.items():
                script += self._generate_character_style_code(style_id, style_info)
        
        # Fin du script
        script += '''
    doc.save(output_path)
    print(f"✓ Template créé: {output_path}")
    return doc


def apply_styles_demo(doc: Document):
    """Démontre l'utilisation des styles."""
    
    print("\\n=== DÉMONSTRATION DES STYLES ===\\n")
    
    # Ajouter quelques paragraphes de démonstration
    # À compléter selon les styles disponibles
    
    return doc


if __name__ == "__main__":
    import sys
    
    output = sys.argv[1] if len(sys.argv) > 1 else "TEMPLATE_STYLES.docx"
    doc = create_styles_template(output)
    apply_styles_demo(doc)
    doc.save(output)
'''
        
        return script
    
    def _generate_paragraph_style_code(self, style_id: str, style_info: Dict) -> str:
        """Génère le code pour créer un style de paragraphe."""
        
        name = style_info['name']
        code = f"\n    # Style: {name}\n"
        code += f"    try:\n"
        code += f"        style = doc.styles.add_style('{style_id}', WD_STYLE_TYPE.PARAGRAPH)\n"
        code += f"        style.name = '{name}'\n"
        
        props = style_info['properties']
        
        # Style parent
        if props.get('based_on'):
            code += f"        # Basé sur: {props['based_on']}\n"
        
        # Propriétés de caractère
        char_props = props.get('character', {})
        if char_props:
            if char_props.get('font_name'):
                code += f"        style.font.name = '{char_props['font_name']}'\n"
            
            if char_props.get('font_size'):
                code += f"        style.font.size = Pt({char_props['font_size']})\n"
            
            if char_props.get('bold'):
                code += f"        style.font.bold = True\n"
            
            if char_props.get('italic'):
                code += f"        style.font.italic = True\n"
            
            if char_props.get('color_rgb'):
                rgb = char_props['color_rgb']
                code += f"        style.font.color.rgb = RGBColor({rgb[0]}, {rgb[1]}, {rgb[2]})\n"
        
        # Propriétés de paragraphe
        para_props = props.get('paragraph', {})
        if para_props:
            if para_props.get('alignment'):
                alignment_map = {
                    'left': 'WD_ALIGN_PARAGRAPH.LEFT',
                    'center': 'WD_ALIGN_PARAGRAPH.CENTER',
                    'right': 'WD_ALIGN_PARAGRAPH.RIGHT',
                    'both': 'WD_ALIGN_PARAGRAPH.JUSTIFY'
                }
                alignment = alignment_map.get(para_props['alignment'], 'WD_ALIGN_PARAGRAPH.LEFT')
                code += f"        style.paragraph_format.alignment = {alignment}\n"
            
            if para_props.get('left_indent'):
                code += f"        style.paragraph_format.left_indent = Inches({para_props['left_indent']/72:.2f})\n"
            
            if para_props.get('spacing_before'):
                code += f"        style.paragraph_format.space_before = Pt({para_props['spacing_before']})\n"
            
            if para_props.get('spacing_after'):
                code += f"        style.paragraph_format.space_after = Pt({para_props['spacing_after']})\n"
        
        code += f"        print(\"✓ Style '{name}' créé\")\n"
        code += f"    except Exception as e:\n"
        code += f"        print(f\"✗ Erreur {name}: {{e}}\")\n"
        
        return code
    
    def _generate_character_style_code(self, style_id: str, style_info: Dict) -> str:
        """Génère le code pour créer un style de caractère."""
        
        name = style_info['name']
        code = f"\n    # Style Caractère: {name}\n"
        code += f"    try:\n"
        code += f"        style = doc.styles.add_style('{style_id}', WD_STYLE_TYPE.CHARACTER)\n"
        code += f"        style.name = '{name}'\n"
        
        props = style_info['properties']
        char_props = props.get('character', {})
        
        if char_props:
            if char_props.get('font_name'):
                code += f"        style.font.name = '{char_props['font_name']}'\n"
            
            if char_props.get('font_size'):
                code += f"        style.font.size = Pt({char_props['font_size']})\n"
            
            if char_props.get('bold'):
                code += f"        style.font.bold = True\n"
            
            if char_props.get('italic'):
                code += f"        style.font.italic = True\n"
            
            if char_props.get('color_rgb'):
                rgb = char_props['color_rgb']
                code += f"        style.font.color.rgb = RGBColor({rgb[0]}, {rgb[1]}, {rgb[2]})\n"
        
        code += f"        print(\"✓ Style Caractère '{name}' créé\")\n"
        code += f"    except Exception as e:\n"
        code += f"        print(f\"✗ Erreur {name}: {{e}}\")\n"
        
        return code
    
    def generate_markdown_guide(self) -> str:
        """Génère un guide Markdown des styles."""
        
        guide = "# 📚 Guide des Styles Personnalisés\n\n"
        guide += f"*Généré automatiquement - {len(self.styles)} style(s) trouvé(s)*\n\n"
        
        guide += "## 📋 Tableau récapitulatif\n\n"
        guide += "| Style | Type | Police | Taille | Gras | Couleur |\n"
        guide += "|-------|------|--------|--------|------|----------|\n"
        
        for style_id, style_info in sorted(self.styles.items(), key=lambda x: x[1]['name']):
            name = style_info['name']
            style_type = style_info['type']
            char_props = style_info['properties'].get('character', {})
            
            font = char_props.get('font_name', '-')
            size = f"{char_props.get('font_size', '-')}pt"
            bold = '✓' if char_props.get('bold') else '-'
            color = f"#{char_props.get('color_hex', '-')}" if char_props.get('color_hex') else '-'
            
            guide += f"| {name} | {style_type} | {font} | {size} | {bold} | {color} |\n"
        
        guide += "\n## 🎨 Détail des styles\n\n"
        
        for style_id, style_info in sorted(self.styles.items(), key=lambda x: x[1]['name']):
            guide += f"### {style_info['name']}\n\n"
            guide += f"- **ID**: `{style_id}`\n"
            guide += f"- **Type**: {style_info['type']}\n"
            
            props = style_info['properties']
            if props.get('based_on'):
                guide += f"- **Basé sur**: {props['based_on']}\n"
            
            char_props = props.get('character', {})
            if char_props.get('font_name'):
                guide += f"- **Police**: {char_props['font_name']}\n"
            if char_props.get('font_size'):
                guide += f"- **Taille**: {char_props['font_size']}pt\n"
            if char_props.get('bold'):
                guide += f"- **Gras**: Oui\n"
            if char_props.get('italic'):
                guide += f"- **Italique**: Oui\n"
            if char_props.get('color_hex'):
                rgb = char_props.get('color_rgb', ())
                guide += f"- **Couleur**: #{char_props['color_hex']} (RGB: {rgb})\n"
            
            guide += "\n"
        
        return guide
    
    def export_generated_script(self, output_path: str):
        """Exporte le script Python généré."""
        script = self.generate_full_script()
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(script)
        print(f"✓ Script généré: {output_path}")
    
    def export_markdown_guide(self, output_path: str):
        """Exporte le guide Markdown."""
        guide = self.generate_markdown_guide()
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(guide)
        print(f"✓ Guide Markdown: {output_path}")


def generate_from_json(json_file: str, output_script: str = None, output_guide: str = None):
    """Génère les fichiers de sortie à partir du JSON des styles."""
    
    if output_script is None:
        output_script = Path(json_file).stem + "_generator.py"
    if output_guide is None:
        output_guide = Path(json_file).stem + "_guide.md"
    
    print(f"\n📝 Génération du code à partir de: {json_file}")
    print("="*60)
    
    generator = StyleCodeGenerator(json_file)
    
    generator.export_generated_script(output_script)
    generator.export_markdown_guide(output_guide)
    
    print("\n✓ Génération terminée!")
    print(f"  - Script Python: {output_script}")
    print(f"  - Guide Markdown: {output_guide}")


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python generate_style_code.py <styles.json>")
        sys.exit(1)
    
    json_file = sys.argv[1]
    generate_from_json(json_file)
