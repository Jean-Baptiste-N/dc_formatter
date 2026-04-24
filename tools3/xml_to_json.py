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

KEYWORDS_HEADER_DOCUMENT = ["dossier de compétences", "dossier de competence", "dossier de competences", "dossier de competences"]
KEYWORDS_MAIN_SKILLS = ["domaine de compétence", "domaine de competence", "domaines de compétence", "domaines de competence", "compétences principales", "competences principales", "compétence", "competence"]
KEYWORDS_EDUCATION = ["formation", "formations", "certifications", "certification", "langue", "langues", "diplôme", "diplome", "diplômes", "diplomes"]
KEYWORDS_LANGUAGES = ["langue", "langues", "français", "anglais", "espagnol", "allemand", "italien", "chinois", "japonais", "russe"]
KEYWORDS_HEADER_EXPERIENCE = ["expérience", "experience"]
KEYWORDS_PROFESSIONAL_EXPERIENCE = ["expérience professionnelle", "experience professionnelle", "expériences professionnelles", "experience professionnelles"]
KEYWORDS_TECHNICAL_SKILLS = ["techniques", "technique", "informatiques", "informatique", "numériques", "numeriques", "numérique", "numerique"]

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
    FIVE_CM_TWIPS = round(5 * 1440 / 2.54)  # ≈ 2835 twips

    defaults = {
        'page_width': A4_WIDTH_TWIPS,
        'page_height': A4_HEIGHT_TWIPS,
        'usable_width': A4_WIDTH_TWIPS - 2 * TWO_CM_TWIPS,  # 21 - 2*2 = 17 cm
        'top_margin': TWO_CM_TWIPS,  # 2 cm en twips
        'bottom_margin': TWO_CM_TWIPS,  # 2 cm en twips
        'left_margin': TWO_CM_TWIPS,  # 2 cm en twips
        'right_margin': TWO_CM_TWIPS,  # 2 cm en twips
        'col_fixed_width': FIVE_CM_TWIPS  # 5 cm en twips
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
                dimensions['col_fixed_width'] = 5 * 567

                # Fusionner avec les valeurs par défaut
                result = defaults.copy()
                result.update(dimensions)
                result['usable_width'] = result['page_width'] - result['left_margin'] - result['right_margin']

                print(f"✅ Dimensions extraites du TEMPLATE:")
                return result

    except Exception as e:
        print(f"❌ Erreur lors de la lecture du template: {e}")
        return defaults

def calculate_table_widths(page_dims: dict = None) -> tuple:
    """
    Calcule les largeurs des colonnes pour les 2 types de tables.

    Type 1 (education): Col1 = 4cm (étroite), Col2 = reste (large)
    Type 2 (professional): Col1 = reste (large), Col2 = 4cm (étroite)

    Args:
        page_dims (dict): Dimensions de page avec left_margin, right_margin, page_width

    Returns:
        tuple: (col_fixed_width, usable_width, type1_widths, type2_widths)
               où type1_widths et type2_widths sont des tuples (col1, col2)
    """
    if page_dims is None:
        page_dims = extract_page_dimensions_from_template()

    col_fixed_width = page_dims['col_fixed_width']
    usable_width = page_dims['usable_width']

    # Type 1 (Education): Col1 = 4cm, Col2 = reste
    type1_col1 = col_fixed_width
    type1_col2 = usable_width - col_fixed_width

    # Type 2 (Professional): Col1 = reste, Col2 = 4cm
    type2_col1 = usable_width - col_fixed_width
    type2_col2 = col_fixed_width

    return (col_fixed_width, usable_width,
            (type1_col1, type1_col2),
            (type2_col1, type2_col2))

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

def get_text_from_element(element: Dict[str, Any]) -> str:
    """Extrait tout le texte d'un élément (paragraphe ou table)"""
    texts = []

    # Texte direct
    if 'text' in element:
        texts.append(element['text'])

    # Texte depuis les runs
    if 'runs' in element:
        for run in element['runs']:
            if 'text' in run:
                texts.append(run['text'])

    # Texte depuis les tableaux
    if element.get('type') == 'Table':
        for row in element.get('rows', []):
            for cell in row.get('cells', []):
                for para in cell.get('paragraphs', []):
                    texts.append(get_text_from_element(para))

    return ' '.join(texts).lower()

def detect_section_by_keyword(text: str) -> str:
    """Détecte le type de section basé sur les mots-clés (en ordre de priorité)"""
    if any(keyword in text for keyword in KEYWORDS_HEADER_DOCUMENT):
        return 'header'
    elif any(keyword in text for keyword in KEYWORDS_MAIN_SKILLS):
        return 'main_skills'
    elif any(keyword in text for keyword in KEYWORDS_EDUCATION):
        return 'education'
    elif any(keyword in text for keyword in KEYWORDS_PROFESSIONAL_EXPERIENCE):
        return 'professional_experience'
    return None

def apply_section_tags(data: Dict[str, Any]) -> None:
    """
    Applique des tags de section à tous les éléments.
    - Commence par 'header'
    - Un tag s'applique à l'élément détecté et tous les suivants
    - Empêche les retours en arrière aux sections antérieures
    - Une section ne peut être visitée qu'une seule fois
    """
    content = data.get('document', {}).get('content', [])

    # Ordre des sections (pour éviter les retours en arrière)
    SECTION_ORDER = ['header', 'main_skills', 'education', 'professional_experience']
    section_indices = {sec: idx for idx, sec in enumerate(SECTION_ORDER)}

    current_section = 'header'  # Commence toujours par header
    current_section_idx = 0

    for element in content:
        # Récupérer le texte de l'élément
        element_text = get_text_from_element(element)

        # Détecter si cet élément déclenche un changement de section
        detected_section = detect_section_by_keyword(element_text)

        if detected_section and detected_section != current_section:
            # Vérifier que ce n'est pas un retour en arrière
            detected_idx = section_indices.get(detected_section, -1)
            if detected_idx >= current_section_idx:
                # Nouvelle section valide (pas un retour en arrière)
                current_section = detected_section
                current_section_idx = detected_idx
            # Sinon, ignorer le changement de section et continuer avec current_section

        # Appliquer le tag current à l'élément
        if 'tags' not in element:
            element['tags'] = []
        if current_section not in element['tags']:
            element['tags'].append(current_section)

def get_table_widths_for_section(section: str = None, page_dims: Dict[str, int] = None) -> tuple:
    """
    Retourne les largeurs des colonnes selon la section.
    Calcule automatiquement basé sur les dimensions du template.

    Args:
        section: 'education', 'professional_experience', ou None (défaut)
        page_dims: Dimensions de page (si None, utilise extract_page_dimensions_from_template())

    Returns:
        Tuple (col1_width, col2_width) en twips
    """
    if page_dims is None:
        page_dims = extract_page_dimensions_from_template()

    _, usable_width, type1_widths, type2_widths = calculate_table_widths(page_dims)

    if section == 'education':
        return type1_widths
    elif section == 'professional_experience':
        return type2_widths
    else:
        return (usable_width // 2, usable_width // 2)

def create_empty_table_2x2(index: int, row_height: int = 360,
                           col1_width: int = None, col2_width: int = None,
                           section: str = None, page_dims: Dict[str, int] = None,
                           auto_generated: bool = False) -> Dict[str, Any]:
    """
    Crée une table 2x2 vide (sans paragraphes de remplissage) avec dimensions spécifiques par colonne.
    Les dimensions sont calculées selon la section et les marges du template.

    Args:
        index: Index dans le document
        row_height: Hauteur de la ligne en twips
        col1_width: Largeur colonne 1 (twips) - si None, utilisé section+page_dims
        col2_width: Largeur colonne 2 (twips) - si None, utilisé section+page_dims
        section: 'education' ou 'professional_experience'
        page_dims: Dimensions de page extraites du template
        auto_generated: Flag pour indiquer que la table a été créée automatiquement

    Returns:
        Table 2x2 structurée avec dimensions (sans paragraphes vides)
    """
    if page_dims is None:
        page_dims = extract_page_dimensions_from_template()

    if col1_width is None or col2_width is None:
        col1_width, col2_width = get_table_widths_for_section(section, page_dims)

    table_total_width = col1_width + col2_width

    # Définir les bordures selon la section
    if section == 'education':
        # Tables éducation : toutes les bordures
        borders = {
            'top': {'size': '10', 'color': '000000'},
            'bottom': {'size': '10', 'color': '000000'},
            'left': {'size': '10', 'color': '000000'},
            'right': {'size': '10', 'color': '000000'},
            'insideH': {'size': '10', 'color': '000000'},
            'insideV': {'size': '10', 'color': '000000'}
        }
    else:
        # Tables expériences professionelles : seulement bottom border
        borders = {
            'top': None,
            'bottom': {'size': '10', 'color': '000000'},
            'left': None,
            'right': None,
            'insideH': None,
            'insideV': None
        }

    return {
        'index': index,
        'type': 'Table',
        'auto_generated': auto_generated,
        'properties': {
            'table_width': str(table_total_width),
            'table_width_type': 'dxa',
            'section': section,
            'borders': borders,
            'style': "DC_Table_Content"
        },
        'row_count': 2,
        'col_count': 2,
        'rows': [
            {
                'row_index': 0,
                'height': row_height,
                'cells': [
                    {
                        'col_index': 0,
                        'width': col1_width,
                        'properties': {
                            'hAlign': 'left',
                            'vAlign': 'center'
                        },
                        'paragraphs': []
                    },
                    {
                        'col_index': 1,
                        'width': col2_width,
                        'properties': {
                            'hAlign': 'right',
                            'vAlign': 'center'
                        },
                        'paragraphs': []
                    }
                ]
            },
            {
                'row_index': 1,
                'height': row_height,
                'cells': [
                    {
                        'col_index': 0,
                        'width': col1_width,
                        'properties': {
                            'hAlign': 'left',
                            'vAlign': 'center'
                        },
                        'paragraphs': []
                    },
                    {
                        'col_index': 1,
                        'width': col2_width,
                        'properties': {
                            'hAlign': 'right',
                            'vAlign': 'center'
                        },
                        'paragraphs': []
                    }
                ]
            }
        ]
    }

def clone_paragraph_clean(para: Dict[str, Any]) -> Dict[str, Any]:
    """
    Clone et nettoie un paragraphe en créant une NOUVELLE structure propre (pas de réutilisation).
    Cela résout le problème de métadonnées XML Word.

    Fait :
    - Crée un nouveau paragraphe JSON (structure indépendante)
    - Supprime les propriétés indésirables (ilvl, numId, size, alignment, color, font)
    - Préserve le style (important pour Word navigation)
    - Clone les runs avec uniquement bold/italic

    Args:
        para: Paragraphe JSON source

    Returns:
        Nouveau paragraphe JSON propre, sans pollution de contexte, prêt pour injecter dans tables
    """
    new_para = {
        "type": "Paragraph",
        "properties": {}
    }

    # Copier et nettoyer les propriétés
    if 'properties' in para:
        source_props = para['properties']

        # Copier le style s'il existe
        if 'style' in source_props:
            new_para['properties']['style'] = source_props['style']

    # Cloner les runs avec nettoyage
    new_para['runs'] = []
    if 'runs' in para:
        for run in para.get('runs', []):
            new_run = {
                "text": run.get('text', ''),
                "properties": {}
            }

            # Copier UNIQUEMENT bold et italic (filtrer les autres propriétés)
            run_props = run.get('properties', {})
            if run_props.get('bold'):
                new_run['properties']['bold'] = True
            if run_props.get('italic'):
                new_run['properties']['italic'] = True

            new_para['runs'].append(new_run)

    # Copier les tags si présents
    if 'tags' in para:
        new_para['tags'] = para['tags'].copy()

    return new_para

def create_language_header(data: Dict[str, Any]) -> None:
    """
    Crée un header "Langues" juste avant le premier élément contenant KEYWORDS_LANGUAGES,
    si ce header n'existe pas déjà.

    Args:
        data: Structure du document JSON
    """
    content = data.get('document', {}).get('content', [])

    # Chercher le premier élément contenant KEYWORDS_LANGUAGES
    first_language_idx = None
    for i, element in enumerate(content):
        if element.get('type') == 'Paragraph':
            text = get_text_from_element(element)
            if any(keyword in text for keyword in KEYWORDS_LANGUAGES):
                first_language_idx = i
                break

    if first_language_idx is None:
        return  # Aucun keyword détecté, rien à faire

    # Vérifier si l'élément précédent est déjà un header "Langues"
    if first_language_idx > 0:
        prev_element = content[first_language_idx - 1]
        if prev_element.get('type') == 'Paragraph':
            prev_text = get_text_from_element(prev_element)
            # Vérifier si c'est un header contenant des keywords de langues
            if any(keyword in prev_text for keyword in KEYWORDS_LANGUAGES):
                return  # Header existe déjà

    # Créer et insérer le header "Langues" juste avant le premier keyword
    new_header = {
        'type': 'Paragraph',
        'runs': [{'text': 'Langues', 'properties': {}}],
        'properties': {},
        'tags': 'education',
        'section': 'education',
        'auto_generated': True
    }
    content.insert(first_language_idx, new_header)

def split_paragraph_at_language(para: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Scinde un paragraphe au premier keyword de langue détecté.

    Crée 2 paragraphes:
    - Avant: le mot-clé de langue détecté (col 0)
    - Après: la description nettoyée (col 1)

    Nettoie le début de la description: supprime " : ", espaces, jusqu'à la première lettre.

    ⚠️ IMPORTANT: Les runs sont normalisés au parsing, donc les keywords
    sont maintenant directement accessibles sans fragmentation.

    Args:
        para: Paragraphe JSON source

    Returns:
        List[Dict]: Liste de 1 ou 2 paragraphes
    """
    text = get_text_from_element(para)

    # Trouver le premier keyword de langue
    lang_keyword = None
    lang_pos = len(text)

    for keyword in KEYWORDS_LANGUAGES:
        pos = text.find(keyword)
        if pos != -1 and pos < lang_pos:
            lang_keyword = keyword
            lang_pos = pos

    if lang_keyword is None:
        return [para]

    # Scinder les runs selon la position du keyword
    lang_runs = []
    desc_runs = []
    current_pos = 0
    keyword_found = False
    keyword_end_pos = lang_pos + len(lang_keyword)

    for run in para.get('runs', []):
        run_text = run.get('text', '')
        run_len = len(run_text)
        run_end = current_pos + run_len

        if not keyword_found:
            if run_end <= lang_pos:
                # Run entièrement avant le keyword, ignorer
                pass
            elif current_pos >= keyword_end_pos:
                # Run entièrement après le keyword
                desc_runs.append(run)
                keyword_found = True
            else:
                # Run contient le keyword
                # Extraire le keyword lui-même
                keyword_start_in_run = lang_pos - current_pos
                keyword_end_in_run = keyword_end_pos - current_pos

                lang_text = run_text[keyword_start_in_run:keyword_end_in_run]
                lang_runs.append({
                    "text": lang_text,
                    "properties": run.get('properties', {})
                })

                # Récupérer le reste du run après le keyword
                rest_text = run_text[keyword_end_in_run:]
                if rest_text:
                    desc_runs.append({
                        "text": rest_text,
                        "properties": run.get('properties', {})
                    })
                keyword_found = True
        else:
            # Après le keyword
            desc_runs.append(run)

        current_pos = run_end

    # Créer 2 paragraphes
    result = []

    # Paragraphe 1 : le mot-clé de langue
    if lang_runs:
        lang_para = clone_paragraph_clean(para)
        lang_para['runs'] = lang_runs
        result.append(lang_para)

    # Paragraphe 2 : la description nettoyée
    if desc_runs:
        desc_para = clone_paragraph_clean(para)

        # Nettoyer le premier run: supprimer tous les caractères jusqu'à la première lettre
        if desc_runs:
            first_run = desc_runs[0]
            first_text = first_run.get('text', '')

            # Enlever tous les caractères jusqu'à la première lettre
            cleaned_text = ''
            for char in first_text:
                if char.isalpha():
                    cleaned_text = first_text[first_text.index(char):]
                    break

            if cleaned_text:
                first_run['text'] = cleaned_text
                desc_para['runs'] = desc_runs
                result.append(desc_para)
            elif len(desc_runs) > 1:
                # Si le premier run est vide après nettoyage, utiliser les runs suivants
                desc_para['runs'] = desc_runs[1:]
                result.append(desc_para)

    return result if result else [para]

def group_education_paragraphs(paragraphs: List[Dict[str, Any]]) -> List[List[Dict[str, Any]]]:
    """
    Groupe les paragraphes éducation en blocs basés sur les dates.

    Logique:
    - Un bloc commence avec une date (contient '20', ' 20', '/20', '-20')
    - Les paragraphes suivants (non-dates) font partie du même bloc
    - Le prochain bloc commence quand une nouvelle date est détectée
    - Les paragraphes vides sont ignorés à la création des blocs

    ⚠️ IMPORTANT: Les runs sont normalisés au parsing, donc les dates
    sont maintenant directement accessibles sans fragmentation.

    Args:
        paragraphs: Liste de paragraphes

    Returns:
        Liste de blocs, où chaque bloc est une liste de paragraphes (sans vides)
        Exemple: [[date_para, desc1, desc2], [date_para, desc1], ...]
    """
    if not paragraphs:
        return []

    blocks = []
    current_block = []

    for para in paragraphs:
        para_text = get_text_from_element(para)

        # Ignorer les paragraphes vides
        if not para_text.strip():
            continue

        is_date = '20' in para_text or ' 20' in para_text or '/20' in para_text or '-20' in para_text

        if is_date:
            # Nouvelle date = nouveau bloc
            if current_block:
                blocks.append(current_block)
            current_block = [para]  # Start new block with this date
        else:
            # Non-date: ajouter au bloc courant
            if current_block:
                current_block.append(para)
            else:
                # Pas de bloc courant, créer un bloc pour ce paragraphe
                current_block = [para]

    # Ajouter le dernier bloc
    if current_block:
        blocks.append(current_block)

    return blocks

def create_edu_table(data: Dict[str, Any], row_height: int = 360, page_dims: Dict[str, int] = None) -> Dict[str, List[int]]:
    """
    Crée les structures des tables éducation (Formation et Langues).

    Responsabilité: Créer les tables vides et les insérer dans le contenu.

    1. Détecte le header "Formation" et crée une table structurée 2xN avec blocs
    2. Détecte le header "Langues" et crée une table structurée 2xN avec paires langue/description
    3. Marque les indices des sources à supprimer

    Args:
        data: Structure du document JSON
        row_height: Hauteur des lignes en twips
        page_dims: Dimensions de page (si None, utilise extract_page_dimensions_from_template())

    Returns:
        Dict contenant:
        - 'formation_paras': Paragraphes de Formation à traiter
        - 'formation_blocks': Blocs Formation groupés
        - 'lang_paras': Paragraphes Langues à traiter
        - 'lang_blocks': Blocs Langues groupés (paires)
        - 'indices_to_delete': Indices à supprimer
    """
    if page_dims is None:
        page_dims = extract_page_dimensions_from_template()

    content = data.get('document', {}).get('content', [])
    indices_to_delete = []
    formation_paras = []
    formation_blocks = []
    lang_paras = []
    lang_blocks = []

    # ===== CRÉATION FORMATION =====
    for i, element in enumerate(content):
        if element.get('type') == 'Paragraph':
            text = get_text_from_element(element)

            # Chercher un vrai header Formation
            is_title = element.get('properties', {}).get('style', '').startswith('Titre')
            is_only_formation = text.strip() in ['formation', 'formations']

            if (any(keyword in text for keyword in KEYWORDS_EDUCATION) and 'formation' in text and
                (is_title or is_only_formation)):

                # Collecter paragraphes/table après Formation
                j = i + 1

                # Cas 1 : table existante
                if j < len(content) and content[j].get('type') == 'Table' and not content[j].get('auto_generated'):
                    existing_table = content[j]
                    for row in existing_table.get('rows', []):
                        for cell in row.get('cells', []):
                            formation_paras.extend(cell.get('paragraphs', []))
                    indices_to_delete.append(j)
                    j += 1

                # Cas 2 : paragraphes directs
                else:
                    while j < len(content):
                        next_elem = content[j]
                        if next_elem.get('type') == 'Paragraph':
                            elem_text = get_text_from_element(next_elem)

                            if not elem_text.strip():
                                j += 1
                                continue

                            if any(keyword in elem_text for keyword in KEYWORDS_PROFESSIONAL_EXPERIENCE):
                                break

                            is_only_langues = elem_text.strip() in ['langues', 'langue']
                            is_title_check = next_elem.get('properties', {}).get('style', '').startswith('Titre')
                            if is_only_langues and (is_title_check or next_elem.get('auto_generated', False)):
                                break

                            formation_paras.append(next_elem)
                            indices_to_delete.append(j)
                            j += 1
                        elif next_elem.get('type') == 'Table':
                            if not next_elem.get('auto_generated'):
                                for row in next_elem.get('rows', []):
                                    for cell in row.get('cells', []):
                                        formation_paras.extend(cell.get('paragraphs', []))
                                indices_to_delete.append(j)
                            break
                        else:
                            j += 1

                # Grouper en blocs
                formation_blocks = group_education_paragraphs(formation_paras)

                # Créer la table vide
                if formation_paras and len(formation_blocks) > 0:
                    new_table = create_empty_table_2x2(
                        i + 1,
                        row_height,
                        section='education',
                        page_dims=page_dims,
                        auto_generated=True
                    )

                    col1_width, col2_width = get_table_widths_for_section('education', page_dims)
                    new_table['rows'] = []
                    for row_idx in range(len(formation_blocks)):
                        new_table['rows'].append({
                            'row_index': row_idx,
                            'height': row_height,
                            'cells': [
                                {
                                    'col_index': 0,
                                    'width': col1_width,
                                    'properties': {'hAlign': 'left', 'vAlign': 'center'},
                                    'paragraphs': []
                                },
                                {
                                    'col_index': 1,
                                    'width': col2_width,
                                    'properties': {'hAlign': 'left', 'vAlign': 'center'},
                                    'paragraphs': []
                                }
                            ]
                        })
                    new_table['row_count'] = len(formation_blocks)

                    # Insérer la table
                    content.insert(i + 1, new_table)
                    indices_to_delete = [idx + 1 if idx > i else idx for idx in indices_to_delete]
                break

    # ===== CRÉATION LANGUES =====
    lang_header_idx = None
    lang_indices = []
    for i, element in enumerate(content):
        if element.get('type') == 'Paragraph':
            text = get_text_from_element(element)
            is_only_langues = text.strip() in ['langues', 'langue']
            is_title = element.get('properties', {}).get('style', '').startswith('Titre')
            is_auto_gen = element.get('auto_generated', False)

            if is_only_langues and (is_title or is_auto_gen):
                lang_header_idx = i
                break

    if lang_header_idx is not None:
        j = lang_header_idx + 1

        # Cas 1 : table existante
        if j < len(content) and content[j].get('type') == 'Table' and not content[j].get('auto_generated'):
            existing_table = content[j]
            for row in existing_table.get('rows', []):
                for cell in row.get('cells', []):
                    lang_paras.extend(cell.get('paragraphs', []))
            indices_to_delete.append(j)
        else:
            # Cas 2 : collecter paragraphes
            while j < len(content):
                next_elem = content[j]
                if next_elem.get('type') == 'Paragraph':
                    elem_text = get_text_from_element(next_elem)

                    if not elem_text.strip():
                        j += 1
                        continue

                    if any(keyword in elem_text for keyword in KEYWORDS_PROFESSIONAL_EXPERIENCE):
                        break

                    is_only_formation = elem_text.strip() in ['formation', 'formations']
                    is_title_check = next_elem.get('properties', {}).get('style', '').startswith('Titre')
                    if is_only_formation and is_title_check:
                        break

                    if any(keyword in elem_text for keyword in KEYWORDS_LANGUAGES):
                        lang_paras.append(next_elem)
                        lang_indices.append(j)
                    j += 1
                elif next_elem.get('type') == 'Table':
                    break
                else:
                    j += 1

        # Créer table Langues
        if lang_paras:
            # Scinder au keyword de langue et regrouper
            split_paras = []
            for para in lang_paras:
                split_paras.extend(split_paragraph_at_language(para))

            for i in range(0, len(split_paras), 2):
                if i + 1 < len(split_paras):
                    lang_blocks.append((split_paras[i], split_paras[i+1]))
                else:
                    lang_blocks.append((split_paras[i], None))

            new_lang_table = create_empty_table_2x2(
                lang_header_idx + 1,
                row_height,
                section='education',
                page_dims=page_dims,
                auto_generated=True
            )

            col1_width, col2_width = get_table_widths_for_section('education', page_dims)
            new_lang_table['rows'] = []
            for row_idx in range(len(lang_blocks)):
                new_lang_table['rows'].append({
                    'row_index': row_idx,
                    'height': row_height,
                    'cells': [
                        {
                            'col_index': 0,
                            'width': col1_width,
                            'properties': {'hAlign': 'left', 'vAlign': 'center'},
                            'paragraphs': []
                        },
                        {
                            'col_index': 1,
                            'width': col2_width,
                            'properties': {'hAlign': 'left', 'vAlign': 'center'},
                            'paragraphs': []
                        }
                    ]
                })
            new_lang_table['row_count'] = len(lang_blocks)

            # Insérer la table
            content.insert(lang_header_idx + 1, new_lang_table)

            # Marquer sources
            for idx in lang_indices:
                if idx > lang_header_idx:
                    indices_to_delete.append(idx + 1)

    data['document']['content'] = content

    return {
        'formation_paras': formation_paras,
        'formation_blocks': formation_blocks,
        'lang_paras': lang_paras,
        'lang_blocks': lang_blocks,
        'indices_to_delete': indices_to_delete
    }

def insert_text_edu_table(data: Dict[str, Any], creation_result: Dict[str, Any]) -> None:
    """
    Remplit le contenu des tables éducation (Formation et Langues) et supprime les sources.

    Responsabilité: Insérer le texte dans les cellules des tables créées et nettoyer les sources.

    Args:
        data: Structure du document JSON
        creation_result: Résultat de create_edu_table() contenant:
            - 'formation_blocks': Blocs Formation groupés
            - 'lang_blocks': Blocs Langues groupés (paires)
            - 'indices_to_delete': Indices à supprimer
    """
    content = data.get('document', {}).get('content', [])
    formation_blocks = creation_result.get('formation_blocks', [])
    lang_blocks = creation_result.get('lang_blocks', [])
    indices_to_delete = creation_result.get('indices_to_delete', [])

    # ===== REMPLIR FORMATION =====
    for elem in content:
        if elem.get('type') == 'Table' and elem.get('auto_generated'):
            section = elem.get('properties', {}).get('section')
            if section == 'education':
                # Identifier si c'est Formation ou Langues
                rows = elem.get('rows', [])

                # Si on trouve une table éducation, on la remplit
                # Vérifier si elle est déjà remplie (a du contenu)
                if all(not cell['paragraphs'] for row in rows for cell in row['cells']):
                    # Table vide, c'est notre table à remplir

                    # Vérifier si c'est Formation (a des blocs de dates) ou Langues
                    if formation_blocks and len(rows) == len(formation_blocks):
                        # C'est Formation
                        for row_idx, block in enumerate(formation_blocks):
                            if row_idx < len(rows):
                                for para_idx, para in enumerate(block):
                                    cloned = clone_paragraph_clean(para)
                                    if para_idx == 0:
                                        rows[row_idx]['cells'][0]['paragraphs'].append(cloned)
                                    else:
                                        rows[row_idx]['cells'][1]['paragraphs'].append(cloned)
                        formation_blocks = []  # Marqué comme traité

                    elif lang_blocks and len(rows) == len(lang_blocks):
                        # C'est Langues
                        for row_idx, (lang_para, desc_para) in enumerate(lang_blocks):
                            if row_idx < len(rows):
                                lang_cloned = clone_paragraph_clean(lang_para)
                                rows[row_idx]['cells'][0]['paragraphs'] = [lang_cloned]

                                if desc_para:
                                    desc_cloned = clone_paragraph_clean(desc_para)
                                    rows[row_idx]['cells'][1]['paragraphs'] = [desc_cloned]
                        lang_blocks = []  # Marqué comme traité

    # ===== SUPPRIMER SOURCES =====
    for idx in sorted(set(indices_to_delete), reverse=True):
        if 0 <= idx < len(content):
            del content[idx]

    data['document']['content'] = content

def create_xp_tables(data: Dict[str, Any], row_height: int = 360, page_dims: Dict[str, int] = None) -> Dict[str, Any]:
    """
    Crée les structures des tables professionnelles (Expériences Professionnelles).

    Responsabilité: Créer les tables vides et les insérer dans le contenu.

    Crée une table 2x2 pour:
    1. Après le header "Expériences Professionnelles" (pour le job entry)
    2. Quand on détecte KEYWORDS_TECHNICAL_SKILLS
    3. SAUF pour les paragraphes avec ilvl (listes/bullets)
    4. SAUF pour les paragraphes commençant par "Contexte"

    Args:
        data: Structure du document JSON
        row_height: Hauteur des lignes en twips
        page_dims: Dimensions de page (si None, utilise extract_page_dimensions_from_template())

    Returns:
        Dict contenant:
        - 'indices_to_delete': Indices à supprimer (vide pour XP)
    """
    if page_dims is None:
        page_dims = extract_page_dimensions_from_template()

    content = data.get('document', {}).get('content', [])

    # Créer des tables selon les conditions
    new_content = []
    just_after_prof_exp_header = False
    current_section = None

    i = 0
    while i < len(content):
        element = content[i]
        new_content.append(element)

        if element.get('type') == 'Paragraph':
            text = get_text_from_element(element)
            has_ilvl = element.get('properties', {}).get('ilvl') is not None

            # Détecter le header "Expériences Professionnelles"
            if any(keyword in text.lower() for keyword in KEYWORDS_PROFESSIONAL_EXPERIENCE):
                current_section = 'professional_experience'
                just_after_prof_exp_header = True
                i += 1
                continue

            # Condition pour créer une table
            should_create_table = False

            # Cas 1 : First paragraph after "Expériences Professionnelles" header
            if just_after_prof_exp_header and not has_ilvl:
                should_create_table = True
                just_after_prof_exp_header = False

            # Cas 2 : Paragraph with KEYWORDS_TECHNICAL_SKILLS and no ilvl
            # SAUF si le paragraphe commence par "Contexte"
            elif current_section == 'professional_experience' and not has_ilvl:
                if not text.startswith('contexte') or not len(text) > 100:
                    if any(keyword in text.lower() for keyword in KEYWORDS_TECHNICAL_SKILLS):
                        should_create_table = True

            if should_create_table and i + 1 < len(content):
                next_element = content[i + 1]
                # Créer table si l'élément suivant n'est pas déjà une table
                if next_element.get('type') != 'Table':
                    new_table = create_empty_table_2x2(
                        len(new_content),
                        row_height,
                        section=current_section,
                        page_dims=page_dims,
                        auto_generated=True
                    )
                    new_content.append(new_table)
                    i += 1
                    continue

        i += 1

    data['document']['content'] = new_content

    return {'indices_to_delete': []}

def insert_text_xp_tables(data: Dict[str, Any], creation_result: Dict[str, Any]) -> None:
    """
    Remplit le contenu des tables professionnelles et supprime les sources.

    Responsabilité: Insérer le texte dans les cellules des tables créées et nettoyer les sources.

    Logique:
    1. Pour chaque table AUTO professional_experience:
    2. Lire jusqu'à 3 paragraphes après (chercher table EXISTING ou paragraphes)
    3. Si on rencontre une table EXISTING: l'extraire et la marquer pour suppression
    4. Fusionner contenu extrait + paragraphes lus
    5. Distribuer dans les cellules

    Args:
        data: Structure du document JSON
        creation_result: Résultat de create_xp_tables() (pour uniformité, même si vide)
    """
    content = data.get('document', {}).get('content', [])
    indices_to_remove = []

    i = 0
    while i < len(content):
        element = content[i]

        if element.get('type') == 'Table' and element.get('auto_generated'):
            section = element.get('properties', {}).get('section')

            if section == 'professional_experience':
                all_paragraphs = []

                # Lire les éléments après la table AUTO (chercher table EXISTING ou paragraphes)
                j = i + 1
                para_count = 0

                while j < len(content) and para_count < 3:
                    next_elem = content[j]
                    elem_type = next_elem.get('type')

                    # Si on rencontre une table EXISTING: l'extraire
                    if elem_type == 'Table' and not next_elem.get('auto_generated'):
                        # Extraire TOUS les paragraphes de cette table EXISTING
                        for row in next_elem.get('rows', []):
                            for cell in row.get('cells', []):
                                all_paragraphs.extend(cell.get('paragraphs', []))
                        # Marquer cette table pour suppression
                        indices_to_remove.append(j)
                        # IMPORTANT: Continuer à lire les paragraphes APRÈS cette table
                        j += 1
                        continue

                    # Si c'est un paragraphe: l'ajouter
                    if elem_type == 'Paragraph':
                        props = next_elem.get('properties', {})
                        text = get_text_from_element(next_elem)

                        # ARRÊTER si ilvl (c'est une liste/puce - style différent)
                        if props.get('ilvl') is not None:
                            break

                        # SKIP si c'est un titre (KEYWORDS_TECHNICAL_SKILLS) - le laisser en place
                        if any(keyword in text for keyword in KEYWORDS_TECHNICAL_SKILLS):
                            break

                        # ARRÊTER si le paragraphe est long (> 100 caractères)
                        if len(text) > 100:
                            break

                        # Ajouter le paragraphe (même s'il est vide)
                        all_paragraphs.append(next_elem)
                        indices_to_remove.append(j)

                        # Compter seulement les paragraphes NON VIDES
                        if text.strip():
                            para_count += 1
                    elif elem_type == 'Table':
                        # Table AUTO ou autre: arrêter
                        break

                    j += 1

                # Étape 2 : Distribuer dans les cellules
                if all_paragraphs:
                    # Trouver max size
                    max_size_para = None
                    max_size = 0
                    remaining = list(all_paragraphs)

                    for para in all_paragraphs:
                        if para.get('runs'):
                            for run in para['runs']:
                                size_str = run.get('properties', {}).get('size')
                                if size_str:
                                    try:
                                        size = int(size_str)
                                        if size > max_size:
                                            max_size = size
                                            max_size_para = para
                                    except ValueError:
                                        pass

                    if max_size_para and max_size_para in remaining:
                        remaining.remove(max_size_para)
                        element['rows'][0]['cells'][0]['paragraphs'] = [clone_paragraph_clean(max_size_para)]

                    # Trouver date (contient "20")
                    date_para = None
                    for para in remaining:
                        text = get_text_from_element(para)
                        if ' 20' in text or '/20' in text or '-20' in text:
                            date_para = para
                            break

                    if date_para and date_para in remaining:
                        remaining.remove(date_para)
                        element['rows'][0]['cells'][1]['paragraphs'] = [clone_paragraph_clean(date_para)]

                    # Placer le reste dans cell[1][0]
                    # Filtrer: garder seulement les paragraphes avec du texte (exclure vides + page_break-only)
                    if remaining:
                        filtered_paras = []
                        for para in remaining:
                            # Un paragraphe est utile s'il a au moins un run avec du texte
                            runs = para.get('runs', [])
                            has_meaningful_content = any('text' in run for run in runs)

                            if has_meaningful_content:
                                # Cloner et nettoyer le paragraphe
                                filtered_paras.append(clone_paragraph_clean(para))

                        # Placer seulement les paragraphes significatifs
                        if filtered_paras:
                            element['rows'][1]['cells'][0]['paragraphs'] = filtered_paras

        i += 1

    # Supprimer en allant de la fin vers le début pour préserver les indices
    for idx in sorted(indices_to_remove, reverse=True):
        if idx < len(content):
            del content[idx]

    # Ajouter 1 paragraphe vide avant et 1 après chaque table AUTO (education et professional_experience)
    # (après les traitements/suppressions, pour qu'ils ne soient pas relus)
    new_content = []
    for element in content:
        if element.get('type') == 'Table' and element.get('auto_generated'):
            section = element.get('properties', {}).get('section')
            if section in ('education', 'professional_experience'):
                # Ajouter 1 paragraphe vide juste avant la table
                new_content.append({'type': 'Paragraph', 'properties': {}, 'runs': []})
        new_content.append(element)
        if element.get('type') == 'Table' and element.get('auto_generated'):
            section = element.get('properties', {}).get('section')
            if section in ('education', 'professional_experience'):
                # Ajouter 1 paragraphe vide juste après la table
                new_content.append({'type': 'Paragraph', 'properties': {}, 'runs': []})

    data['document']['content'] = new_content

def remove_double_paras_and_spaces (data: Dict[str, Any]) -> None:
    """
    Supprime les paragraphes vides doublons et nettoie les doubles espaces.
    Modifie in-place.

    Logique:
    - Parcourir le contenu du document
    - Garder une trace du dernier paragraphe ajouté
    - Supprimer les paragraphes vides doublons (garder max 1 paragraphe vide consécutif)
    - Remplacer les doubles espaces ("  ") par un simple espace (" ") dans les runs
    """
    content = data.get('document', {}).get('content', [])
    new_content = []
    last_para_was_empty = False

    for element in content:
        if element.get('type') == 'Paragraph':
            text = get_text_from_element(element)
            is_empty = not text.strip()

            if is_empty:
                # Garder seulement 1 paragraphe vide (éviter 2 consécutifs)
                if not last_para_was_empty:
                    new_content.append(element)
                last_para_was_empty = True
            else:
                # Paragraphe non-vide : nettoyer les doubles espaces dans les runs
                if 'runs' in element:
                    for run in element['runs']:
                        if 'text' in run:
                            # Remplacer les doubles espaces par un simple espace
                            run['text'] = run['text'].replace('  ', ' ')

                new_content.append(element)
                last_para_was_empty = False
        else:
            new_content.append(element)
            last_para_was_empty = False

    data['document']['content'] = new_content

def add_colons_between_list_levels(data: Dict[str, Any]) -> None:
    """
    Ajoute des ":" entre deux niveaux de listes successifs (1→2, 2→3, etc).
    SAUF entre ilvl 0 et 1 (et supprime le ":" s'il existe).
    
    Logique:
    - Parcourir les paragraphes avec ilvl
    - Si transition vers niveau supérieur (sauf 0→1): ajouter " :" si absent
    - Si transition 0→1: SUPPRIMER le ":" s'il existe

    Modifie in-place.
    """
    content = data.get('document', {}).get('content', [])

    for i in range(len(content) - 1):
        element = content[i]
        next_element = content[i + 1]

        if element.get('type') == 'Paragraph' and next_element.get('type') == 'Paragraph':
            curr_ilvl = element.get('properties', {}).get('ilvl')
            next_ilvl = next_element.get('properties', {}).get('ilvl')

            # Vérifier s'il y a une transition vers un niveau supérieur
            if curr_ilvl is not None and next_ilvl is not None:
                try:
                    curr_ilvl_int = int(curr_ilvl) if isinstance(curr_ilvl, str) else curr_ilvl
                    next_ilvl_int = int(next_ilvl) if isinstance(next_ilvl, str) else next_ilvl

                    if next_ilvl_int > curr_ilvl_int:
                        # Transition 0→1: SUPPRIMER le ":" s'il existe
                        if curr_ilvl_int == 0 and next_ilvl_int == 1:
                            if 'runs' in element and len(element['runs']) > 0:
                                last_run = element['runs'][-1]
                                if 'text' in last_run:
                                    # Supprimer " :" ou ":" à la fin
                                    last_run['text'] = last_run['text'].rstrip()
                                    if last_run['text'].endswith(' :'):
                                        last_run['text'] = last_run['text'][:-2]
                                    elif last_run['text'].endswith(':'):
                                        last_run['text'] = last_run['text'][:-1]
                        # Autres transitions: AJOUTER ":" s'il n'existe pas
                        else:
                            text = get_text_from_element(element)
                            if ':' not in text:
                                # Ajouter " :" à la fin du dernier run du paragraphe courant
                                if 'runs' in element and len(element['runs']) > 0:
                                    last_run = element['runs'][-1]
                                    if 'text' in last_run:
                                        last_run['text'] += ' :'
                except (ValueError, TypeError):
                    # Ignorer les conversions invalides
                    pass

def apply_styles_in_json(data: Dict[str, Any]) -> None:
    """
    Applique les styles par défaut dans les données JSON.
    Ajoute aussi l'outline_level selon le style (pour volet de navigation Word).
    Modifie in-place.

    Mapping style → outline_level :
    - DC_T1_Sections (niveau 1) → outline_level = 0
    - DC_XP_Title (niveau 2) → outline_level = 1
    - DC_1st_bullet (niveau 3) → outline_level = 2

    Args:
        data (Dict): Structure JSON à modifier
    """
    # Mapping style → outline_level pour Word navigation
    STYLE_OUTLINE_MAPPING = {
        'DC_T1_Sections': 0,  # niveau 1
        'DC_XP_Title': 1,     # niveau 2
        'DC_1st_bullet': 2,   # niveau 3
    }

    # Appliquer les styles des titres
    for itag in data.get('document', {}).get('content', []):
        if 'tags' not in itag:
            continue
        tags = itag['tags']
        # Extraire le texte
        text = get_text_from_element(itag) if itag else ""
        props = itag.get('properties', {})

        if 'header' in tags and any(keyword in text.lower() for keyword in KEYWORDS_HEADER_DOCUMENT):
            props['style'] = 'DC_H_DC'
            text = text.upper()
            if 'runs' in itag and itag['runs']:
                itag['runs'][0]['text'] = text
        elif 'main_skills' in tags and any(keyword in text.lower() for keyword in KEYWORDS_MAIN_SKILLS):
            props['style'] = 'DC_T1_Sections'
        elif 'education' in tags and any(keyword in text for keyword in KEYWORDS_EDUCATION):
            props['style'] = 'DC_T1_Sections'
        elif 'professional_experience' in tags and any(keyword in text for keyword in KEYWORDS_PROFESSIONAL_EXPERIENCE):
            props['style'] = 'DC_T1_Sections'

        if 'header' in tags and 'DC_H_DC' not in props.get('style', '') and len(text) > 0 and len(text) <= 5:
            props['style'] = 'DC_H_Trigramme'
            text = text.upper()
            if 'runs' in itag and itag['runs']:
                itag['runs'][0]['text'] = text
        elif 'header' in tags and 'DC_H_DC' not in props.get('style', '') and any(keyword in text.lower() for keyword in KEYWORDS_HEADER_EXPERIENCE) and len(text) > 5:
            props['style'] = 'DC_H_XP'

        if 'header' in tags and 'DC_H_XP' not in props.get('style', '') and 'DC_H_DC' not in props.get('style', '') and len(text) > 5:
            props['style'] = 'DC_H_Poste'

    for itable in data.get('document', {}).get('content', []):
        if itable.get('type') == 'Table' and 'properties' in itable:
            section = itable.get('properties', {}).get('section')
            if section == 'education':
                rows = itable.get('rows', [])
                # Appliquer le style DC_Table_Year aux paragraphes dans cell[x][0] (colonne 0)
                for row in rows:
                    cells = row.get('cells', [])
                    if len(cells) > 0:
                        for para in cells[0].get('paragraphs', []):
                            if 'properties' not in para:
                                para['properties'] = {}
                            para['properties']['style'] = 'DC_Table_Year'
                # Appliquer le style DC_Table_Content aux paragraphes dans cell[x][1] (colonne 1)
                for row in rows:
                    cells = row.get('cells', [])
                    if len(cells) > 1:
                        for para in cells[1].get('paragraphs', []):
                            if 'properties' not in para:
                                para['properties'] = {}
                            para['properties']['style'] = 'DC_Table_Content'
            elif section == 'professional_experience':
                rows = itable.get('rows', [])
                # Appliquer le style DC_XP_Title aux paragraphes dans cell[0][0]
                if len(rows) > 0:
                    cells = rows[0].get('cells', [])
                    if len(cells) > 0:
                        for para in cells[0].get('paragraphs', []):
                            if 'properties' not in para:
                                para['properties'] = {}
                            para['properties']['style'] = 'DC_XP_Title'
                            # Ajouter outline_level pour DC_XP_Title
                            if 'DC_XP_Title' in STYLE_OUTLINE_MAPPING:
                                para['properties']['outline_level'] = STYLE_OUTLINE_MAPPING['DC_XP_Title']
                # Appliquer le style DC_XP_Date aux paragraphes dans cell[0][1]
                if len(rows) > 0:
                    cells = rows[0].get('cells', [])
                    if len(cells) > 1:
                        for para in cells[1].get('paragraphs', []):
                            if 'properties' not in para:
                                para['properties'] = {}
                            para['properties']['style'] = 'DC_XP_Date'
                # Appliquer le style DC_XP_Poste aux lignes suivantes (cell[1][0])
                if len(rows) > 1:
                    cells = rows[1].get('cells', [])
                    if len(cells) > 0:
                        for para in cells[0].get('paragraphs', []):
                            if 'properties' not in para:
                                para['properties'] = {}
                            para['properties']['style'] = 'DC_XP_Poste'
                # Appliquer le style DC_Normal à tous les paragraphes vides de textes restants
                for row in rows:
                    for cell in row.get('cells', []):
                        for para in cell.get('paragraphs', []):
                            if not para.get('text'):
                                if 'properties' not in para:
                                    para['properties'] = {}
                                if 'style' not in para['properties']:
                                    para['properties']['style'] = 'DC_Table_Content'

    # Appliquer le highlight pour les compétences techniques
    for itag in data.get('document', {}).get('content', []):
        if 'tags' not in itag:
            continue
        tags = itag['tags']
        text = get_text_from_element(itag)  # Déjà en minuscules

        if 'professional_experience' in tags and any(keyword in text for keyword in KEYWORDS_TECHNICAL_SKILLS) and 'contexte' not in text:
            if 'properties' not in itag:
                itag['properties'] = {}
            itag['properties']['style'] = 'DC_XP_BlueContent'

    # Appliquer les styles des listes
    for ilist in data.get('document', {}).get('content', []):
        if 'properties' not in ilist:
            continue
        props = ilist['properties']
        if 'ilvl' not in props:
            continue

        ilvl = props.get('ilvl')
        text = get_text_from_element(ilist)
        if not ilvl:
            continue
        elif ilvl == "0":
            props['style'] = 'DC_1st_bullet'
            # Ajouter outline_level pour DC_1st_bullet
            if 'DC_1st_bullet' in STYLE_OUTLINE_MAPPING:
                props['outline_level'] = STYLE_OUTLINE_MAPPING['DC_1st_bullet']
                text = text.capitalize()
                if 'runs' in ilist and ilist['runs']:
                    ilist['runs'][0]['text'] = text
        elif ilvl == "1":
            props['style'] = 'DC_2nd_bullet'
        elif ilvl == "2":
            props['style'] = 'DC_3rd_bullet'
        elif ilvl == "3":
            props['style'] = 'DC_4th_bullet'
        else:
            props['style'] = 'DC_Normal'  # fallback

    # Appliquer le style Normal pour le reste et les éléments sans style
    for element in data.get('document', {}).get('content', []):
        props = element.get('properties', {})

        # Forcer DC_Normal pour les paragraphes vides (peu importe leur style d'origine)
        is_empty_para = (not element.get('runs') or all(not run.get('text', '').strip() for run in element.get('runs', [])))

        if element.get('type') == 'Paragraph' and is_empty_para:
            props['style'] = 'DC_Normal'
        else:
            # Pour les paragraphes avec du texte: si le style n'est pas un style DC_* ou n'existe pas, appliquer DC_Normal
            current_style = props.get('style', '')
            if not current_style.startswith('DC_'):
                props['style'] = 'DC_Normal'

        # Ajouter outline_level si le style le nécessite
        if props.get('style') in STYLE_OUTLINE_MAPPING:
            props['outline_level'] = STYLE_OUTLINE_MAPPING[props['style']]

        # Nettoyer les propriétés qui outrepassent le style si un style a été appliqué
        if props.get('style'):
            props.pop('size', None)
            props.pop('alignment', None)
            props.pop('color', None)
            props.pop('font', None)

        # Nettoyer aussi les runs des paragraphes (garder bold/italic uniquement)
        if element.get('type') == 'Paragraph' and 'runs' in element:
            for run in element['runs']:
                if 'properties' in run:
                    run_props = run['properties']
                    kept_props = {}
                    if 'bold' in run_props:
                        kept_props['bold'] = run_props['bold']
                    if 'italic' in run_props:
                        kept_props['italic'] = run_props['italic']
                    run['properties'] = kept_props

    # Nettoyer aussi les propriétés des paragraphes à l'intérieur des tables
    # (pour les tables EXISTANTES et AUTO-GÉNÉRÉES)
    for element in data.get('document', {}).get('content', []):
        if element.get('type') == 'Table':
            for row in element.get('rows', []):
                for cell in row.get('cells', []):
                    for para in cell.get('paragraphs', []):
                        if 'properties' not in para:
                            continue
                        para_props = para['properties']

                        # Forcer DC_Table_Content pour les paragraphes vides dans les tables
                        runs = para.get('runs', [])
                        is_empty = (not runs or all(not run.get('text', '').strip() for run in runs))

                        if is_empty:
                            para_props['style'] = 'DC_Table_Content'

                        # Ajouter outline_level si le style le nécessite
                        if para_props.get('style') in STYLE_OUTLINE_MAPPING:
                            para_props['outline_level'] = STYLE_OUTLINE_MAPPING[para_props['style']]

                        # Nettoyer size, alignment, color, font si un style a été appliqué
                        if para_props.get('style'):
                            para_props.pop('size', None)
                            para_props.pop('alignment', None)
                            para_props.pop('color', None)
                            para_props.pop('font', None)

                        # Nettoyer aussi les runs (garder bold/italic uniquement)
                        if 'runs' in para:
                            for run in para['runs']:
                                if 'properties' in run:
                                    run_props = run['properties']
                                    kept_props = {}
                                    if 'bold' in run_props:
                                        kept_props['bold'] = run_props['bold']
                                    if 'italic' in run_props:
                                        kept_props['italic'] = run_props['italic']
                                    run['properties'] = kept_props

def apply_tags_and_styles(raw_json_file: str, output_dir: str = None, template_path: str = None) -> str:
    """
    Charge un JSON brut, applique les tags de section et les styles,
    puis enregistre le résultat transformé.

    Les dimensions des tables sont extraites automatiquement du TEMPLATE.docx.

    Args:
        raw_json_file (str): Chemin du fichier JSON RAW
        output_dir (str): Répertoire de sortie (optionnel, défaut: output/)
        template_path (str): Chemin du TEMPLATE.docx (optionnel)

    Returns:
        str: Chemin du fichier créé
    """
    input_path = Path(raw_json_file)

    # Générer le répertoire de sortie
    if output_dir is None:
        output_dir = Path(input_path.parent.parent) / 'output'
    else:
        output_dir = Path(output_dir)

    # Créer le répertoire s'il n'existe pas
    output_dir.mkdir(parents=True, exist_ok=True)

    # Générer le nom de sortie
    output_file = output_dir / (input_path.stem.replace('_raw', '') + "_transformed.json")

    # Charger le JSON RAW
    with open(input_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # Extraire les dimensions depuis le template
    if template_path is None:
        template_path = 'assets/TEMPLATE.docx'
    page_dims = extract_page_dimensions_from_template(template_path)

    # Stocker les dimensions dans le document pour utilisation ultérieure
    if 'page_dimensions' not in data:
        data['page_dimensions'] = page_dims

    # ===== DETECTER LES 4 SECTIONS =====
    # Appliquer les tags de section
    apply_section_tags(data)

    # ===== TABLES ÉDUCATION =====
    # Créer le header "Langues" juste avant le premier keyword détecté
    create_language_header(data)

    # Créer les structures (Formation et Langues)
    edu_creation_result = create_edu_table(data, row_height=360, page_dims=page_dims)

    # Remplir le contenu et supprimer les sources
    insert_text_edu_table(data, edu_creation_result)

    # ===== TABLES EXPÉRIENCES PROFESSIONNELLES =====
    # Créer les structures
    xp_creation_result = create_xp_tables(data, row_height=360, page_dims=page_dims)

    # Remplir le contenu et supprimer les sources
    insert_text_xp_tables(data, xp_creation_result)

    # ===== NETTOYAGE et RENDU FINAL POUR CHAQUE ELEMENT =====
    # Ajouter les ":" entre les niveaux de listes successifs
    add_colons_between_list_levels(data)

    # Nettoyer les paragraphes vides doublons et les doubles espaces
    remove_double_paras_and_spaces(data)

    # Appliquer les styles
    apply_styles_in_json(data)

    # Sauvegarder le JSON transformé
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"✅ {output_file.name} créé (avec tags, styles et dimensions des tables)")

    return str(output_file)

def main():
    """
    Fonction principale: orchestre tout le pipeline
    - Génère le JSON RAW depuis le XML
    - Applique les tags et styles
    - Enregistre les deux versions
    """
    parser = ArgumentParser(description="Extrait et transforme les balises d'un fichier global.xml en JSON")

    parser.add_argument(
        "xml_file",
        help="Chemin du fichier global.xml"
    )

    parser.add_argument(
        "-o", "--output",
        help="Répertoire de sortie pour les fichiers transformés (optionnel, défaut: output/)"
    )

    args = parser.parse_args()

    print(f"🔄 Traitement de {args.xml_file}...")

    # Étape 1: Générer le JSON RAW
    xml_path = Path(args.xml_file)
    raw_json_file = xml_path.parent / (xml_path.stem + "_raw.json")
    xml_to_json(args.xml_file, str(raw_json_file))

    # Étape 2: Appliquer tags et styles
    apply_tags_and_styles(str(raw_json_file), args.output)

    print(f"\n✨ Pipeline complété!")

if __name__ == "__main__":
    main()
