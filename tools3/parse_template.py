"""
Script simplifié pour extraire un fichier XML global en JSON avec tous les détails.
Entrée: fichier _GLOBAL.xml
Sortie: fichier _GLOBAL_raw.json (xml brut traduit en json)
Sortie: fichier _GLOBAL_transformed.json (après taggings et transformations)
"""

from argparse import ArgumentParser
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

def extract_page_dimensions_from_template(template_path: str = None) -> dict:
    """
    Extrait les dimensions de page du fichier TEMPLATE.docx.

    Récupère:
    - Largeur de page (w:w)
    - Hauteur de page (w:h)
    - Marges: top, bottom, left, right

    Args:
        template_path (str): Chemin du TEMPLATE.docx

    Returns:
        dict avec keys: page_width, page_height, usable_width, top_margin, bottom_margin,
                        left_margin, right_margin (en twips)
    """
    if template_path is None:
        template_path = 'assets/TEMPLATE.docx'

    template_path = Path(template_path)

    # Valeurs par défaut (A4 exact: 210×297 mm = 21cm×29.7cm)
    # Conversion précise: 1 cm = 1440/2.54 ≈ 566.9291 twips
    # A4 width: 21 cm = 11905.51... twips ≈ 11906 twips
    # A4 height: 29.7 cm = 16838.58... twips ≈ 16838 twips
    A4_WIDTH_TWIPS = round(21 * 1440 / 2.54)  # ≈ 11906 twips
    A4_HEIGHT_TWIPS = round(29.7 * 1440 / 2.54)  # ≈ 16838 twips
    TWO_CM_TWIPS = round(2 * 1440 / 2.54)  # ≈ 1134 twips
    THREE_CM_TWIPS = round(3 * 1440 / 2.54)  # ≈ 1701 twips
    FIVE_CM_TWIPS = round(5 * 1440 / 2.54)  # ≈ 2835 twips

    defaults = {
        'page_width': A4_WIDTH_TWIPS,
        'page_height': A4_HEIGHT_TWIPS,
        'usable_width': A4_WIDTH_TWIPS - 2 * TWO_CM_TWIPS,  # 21 - 2*2 = 17 cm
        'top_margin': TWO_CM_TWIPS,  # 2 cm en twips
        'bottom_margin': TWO_CM_TWIPS,  # 2 cm en twips
        'left_margin': TWO_CM_TWIPS,  # 2 cm en twips
        'right_margin': TWO_CM_TWIPS,  # 2 cm en twips
        'col_fixed_width_3': THREE_CM_TWIPS,  # 3 cm en twips
        'col_fixed_width_5': FIVE_CM_TWIPS  # 5 cm en twips
    }

    if not template_path.exists():
        print(f"⚠️ Template non trouvé: {template_path}, utilisation des valeurs par défaut")
        return defaults

    try:
        with zipfile.ZipFile(template_path, 'r') as zip_ref:
            with zip_ref.open('word/document.xml') as f:
                content = f.read().decode('utf-8')

                dimensions = {}

                # Extraire pgSz (page size)
                import re
                if 'pgSz' in content:
                    start = content.find('<w:pgSz')
                    end = content.find('/>', start)
                    if start != -1:
                        pgSz_line = content[start:end+2]
                        width_match = re.search(r'w:w="(\d+)"', pgSz_line)
                        height_match = re.search(r'w:h="(\d+)"', pgSz_line)
                        if width_match:
                            dimensions['page_width'] = int(width_match.group(1))
                        if height_match:
                            dimensions['page_height'] = int(height_match.group(1))

                # Extraire pgMar (page margins)
                if 'pgMar' in content:
                    start = content.find('<w:pgMar')
                    end = content.find('/>', start)
                    if start != -1:
                        pgMar_line = content[start:end+2]
                        top_match = re.search(r'w:top="(\d+)"', pgMar_line)
                        bottom_match = re.search(r'w:bottom="(\d+)"', pgMar_line)
                        left_match = re.search(r'w:left="(\d+)"', pgMar_line)
                        right_match = re.search(r'w:right="(\d+)"', pgMar_line)

                        if top_match:
                            dimensions['top_margin'] = int(top_match.group(1))
                        if bottom_match:
                            dimensions['bottom_margin'] = int(bottom_match.group(1))
                        if left_match:
                            dimensions['left_margin'] = int(left_match.group(1))
                        if right_match:
                            dimensions['right_margin'] = int(right_match.group(1))

                # Ajouter largeur colonne fixe
                dimensions['col_fixed_width_3'] = THREE_CM_TWIPS
                dimensions['col_fixed_width_5'] = FIVE_CM_TWIPS

                # Fusionner avec les valeurs par défaut
                result = defaults.copy()
                result.update(dimensions)
                result['usable_width'] = result['page_width'] - result['left_margin'] - result['right_margin']

                print(f"✅ Dimensions extraites du TEMPLATE:")
                return result

    except Exception as e:
        print(f"❌ Erreur lors de la lecture du template: {e}")
        return defaults

def main():
    parser = ArgumentParser(description="Extrait les dimensions de page du TEMPLATE.docx")
    parser.add_argument('--template', type=str, default='assets/TEMPLATE.docx', help='Chemin du TEMPLATE.docx')
    args = parser.parse_args()

    extract_page_dimensions_from_template(args.template)
    print('Template extracted!')

if __name__ == "__main__":
    main()