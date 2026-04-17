"""
Script simplifié pour extraire un fichier XML global en JSON avec tous les détails.
Entrée: fichier global.xml
Sortie: fichier _balises.json
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
            document_structure["document"]["content"].append(para_obj)
            content_index += 1

        # Tableaux
        elif element.tag == f'{{{NS["w"]}}}tbl':
            table_obj = parse_table(element, NS, content_index)
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

    return {
        'index': index,
        'type': 'Table',
        'auto_generated': auto_generated,
        'properties': {
            'table_width': str(table_total_width),
            'table_width_type': 'dxa',
            'section': section,
            'borders': {
                'top': None,
                'bottom': {'size': '12', 'color': '000000'},
                'left': None,
                'right': None,
                'insideH': None,
                'insideV': None
            }
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
            text = get_text_from_element(element).lower()
            if any(keyword in text for keyword in KEYWORDS_LANGUAGES):
                first_language_idx = i
                break

    if first_language_idx is None:
        return  # Aucun keyword détecté, rien à faire

    # Vérifier si l'élément précédent est déjà un header "Langues"
    if first_language_idx > 0:
        prev_element = content[first_language_idx - 1]
        if prev_element.get('type') == 'Paragraph':
            prev_text = get_text_from_element(prev_element).lower()
            # Vérifier si c'est un header contenant des keywords de langues
            if any(keyword in prev_text for keyword in KEYWORDS_LANGUAGES):
                return  # Header existe déjà

    # Créer et insérer le header "Langues" juste avant le premier keyword
    new_header = {
        'type': 'Paragraph',
        'text': 'Langues',
        'properties': {},
        'tags': 'education',
        'section': 'education',
        'auto_generated': True
    }
    content.insert(first_language_idx, new_header)


def create_edu_tables(data: Dict[str, Any], row_height: int = 360, page_dims: Dict[str, int] = None) -> None:
    """
    Crée et restructure les tables de la section "education":
    
    1. Après "Formation" : détecte table existante, la remplace par une table auto (2xN)
       Redistribue le contenu : dates en col 0, textes en col 1
    
    2. Après "Langues" : crée une table 2x2 par défaut
       Redistribue les paragraphes de langues : langue en col 0, description en col 1
    
    Args:
        data: Structure du document JSON
        row_height: Hauteur des lignes en twips
        page_dims: Dimensions de page (si None, utilise extract_page_dimensions_from_template())
    """
    if page_dims is None:
        page_dims = extract_page_dimensions_from_template()

    content = data.get('document', {}).get('content', [])
    new_content = []
    indices_to_remove = []
    
    i = 0
    while i < len(content):
        element = content[i]
        new_content.append(element)

        if element.get('type') == 'Paragraph':
            text = get_text_from_element(element).lower()
            
            # ===== CAS 1 : FORMATION =====
            if any(keyword in text for keyword in KEYWORDS_EDUCATION) and 'formation' in text:
                # Chercher la table existante après ce header
                if i + 1 < len(content):
                    next_elem = content[i + 1]
                    if next_elem.get('type') == 'Table' and not next_elem.get('auto_generated'):
                        # Table existante trouvée
                        existing_table = next_elem
                        num_rows = len(existing_table.get('rows', []))
                        
                        # Extraire tout le contenu de la table existante
                        all_paras_from_table = []
                        for row in existing_table.get('rows', []):
                            for cell in row.get('cells', []):
                                all_paras_from_table.extend(cell.get('paragraphs', []))
                        
                        # Créer une nouvelle table auto avec autant de lignes que la table existante
                        new_table = create_empty_table_2x2(
                            len(new_content),
                            row_height,
                            section='education',
                            page_dims=page_dims,
                            auto_generated=True
                        )
                        
                        # Adapter le nombre de lignes de la table auto à celui de la table existante
                        if num_rows != 2:
                            # Recréer les rows manquantes
                            new_table['rows'] = []
                            col1_width, col2_width = get_table_widths_for_section('education', page_dims)
                            for row_idx in range(num_rows):
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
                            new_table['row_count'] = num_rows
                        
                        # Distribuer les paragraphes extraits dans la nouvelle table
                        # Logique : essayer de deviner dates (col 0) vs contenu (col 1)
                        for row_idx, para in enumerate(all_paras_from_table):
                            if row_idx < num_rows:
                                clean_paragraph_for_table(para)
                                
                                # Vérifier si c'est une date (contient "20xx")
                                para_text = get_text_from_element(para)
                                if ' 20' in para_text or '/20' in para_text or '-20' in para_text:
                                    # Date en col 0
                                    new_table['rows'][row_idx]['cells'][0]['paragraphs'] = [para]
                                else:
                                    # Contenu en col 1
                                    new_table['rows'][row_idx]['cells'][1]['paragraphs'] = [para]
                        
                        new_content.append(new_table)
                        indices_to_remove.append(i + 1)  # Marquer la table existante pour suppression
                        i += 2
                        continue
        
        i += 1
    
    # Supprimer les tables existantes qui ont été traitées
    for idx in sorted(indices_to_remove, reverse=True):
        if idx < len(content):
            del content[idx]
    
    # ===== CAS 2 : LANGUES (créer une table et insérer les paragraphes de langues) =====
    # Chercher le header "Langues" créé précédemment
    lang_header_idx = None
    for i, element in enumerate(new_content):
        if element.get('type') == 'Paragraph':
            text = get_text_from_element(element).lower()
            if 'langues' in text and 'header' in element.get('tags', []):
                lang_header_idx = i
                break
    
    if lang_header_idx is not None:
        # Vérifier s'il y a une table existante juste après le header "Langues"
        existing_lang_table = None
        existing_lang_table_idx = None
        if lang_header_idx + 1 < len(new_content):
            elem = new_content[lang_header_idx + 1]
            if elem.get('type') == 'Table' and not elem.get('auto_generated'):
                existing_lang_table = elem
                existing_lang_table_idx = lang_header_idx + 1
        
        # Déterminer le nombre de lignes de la table
        if existing_lang_table:
            # Utiliser le nombre de lignes de la table existante
            num_rows = len(existing_lang_table.get('rows', []))
            
            # Extraire tout le contenu de la table existante
            all_paras_from_lang_table = []
            for row in existing_lang_table.get('rows', []):
                for cell in row.get('cells', []):
                    all_paras_from_lang_table.extend(cell.get('paragraphs', []))
        else:
            # Collecter tous les paragraphes de langues après le header jusqu'au prochain header/table
            language_paras = []
            for i in range(lang_header_idx + 1, len(new_content)):
                elem = new_content[i]
                
                # Arrêter si on rencontre un autre header ou une section différente
                if elem.get('type') == 'Paragraph':
                    elem_text = get_text_from_element(elem).lower()
                    # Arrêter si c'est un autre header
                    if any(keyword in elem_text for keyword in KEYWORDS_EDUCATION) and 'formation' in elem_text:
                        break
                    if any(keyword in elem_text for keyword in KEYWORDS_PROFESSIONAL_EXPERIENCE):
                        break
                    
                    # Si c'est un keyword de langue, le collecter
                    if any(keyword in elem_text for keyword in KEYWORDS_LANGUAGES):
                        language_paras.append(elem)
                elif elem.get('type') == 'Table':
                    break
            
            num_rows = len(language_paras)
            all_paras_from_lang_table = language_paras
        
        # Créer la table seulement s'il y a du contenu
        if num_rows > 0:
            # Créer une nouvelle table auto avec le nombre de lignes détecté
            new_lang_table = create_empty_table_2x2(
                lang_header_idx + 1,
                row_height,
                section='education',
                page_dims=page_dims,
                auto_generated=True
            )
            
            # Recréer autant de lignes que nécessaire
            col1_width, col2_width = get_table_widths_for_section('education', page_dims)
            new_lang_table['rows'] = []
            for row_idx in range(num_rows):
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
            new_lang_table['row_count'] = num_rows
            
            # Distribuer les paragraphes dans la table
            # Format : langue en col 0, description/reste en col 1
            for row_idx, para in enumerate(all_paras_from_lang_table):
                if row_idx < num_rows:
                    clean_paragraph_for_table(para)
                    # La langue va en col 0
                    new_lang_table['rows'][row_idx]['cells'][0]['paragraphs'] = [para]
            
            # Insérer la table juste après le header "Langues"
            new_content.insert(lang_header_idx + 1, new_lang_table)
            
            # Si une table existante était présente, la supprimer
            if existing_lang_table_idx is not None:
                # L'index a changé après l'insertion, donc ajouter 1
                new_content.pop(existing_lang_table_idx + 1)
    
    data['document']['content'] = new_content


def create_xp_tables(data: Dict[str, Any], row_height: int = 360, page_dims: Dict[str, int] = None) -> None:
    """
    Crée des tables 2x2 :
    1. Après le header "Expériences Professionnelles" (pour le job entry)
    2. Quand on détecte KEYWORDS_TECHNICAL_SKILLS
    3. SAUF pour les paragraphes avec ilvl (listes/bullets)
    4. SAUF pour les paragraphes commençant par "Contexte"

    Args:
        data: Structure du document JSON
        row_height: Hauteur des lignes en twips
        page_dims: Dimensions de page (si None, utilise extract_page_dimensions_from_template())
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
            text = get_text_from_element(element).lower()
            has_ilvl = element.get('properties', {}).get('ilvl') is not None

            # Détecter le header "Expériences Professionnelles"
            if any(keyword in text for keyword in KEYWORDS_PROFESSIONAL_EXPERIENCE):
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
                if not text.startswith('contexte'):  # Exclure les paragraphes commençant par "Contexte"
                    if any(keyword in text for keyword in KEYWORDS_TECHNICAL_SKILLS):
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

def clean_paragraph_for_table(para: Dict[str, Any]) -> None:
    """
    Nettoie les propriétés d'un paragraphe avant de le placer dans une table.
    Supprime :
    - Les propriétés de liste (ilvl, numId) pour éviter les numérotations indésirées
    - Les propriétés de formatage directives (size, alignment, color, font) qui outrepassent le style
    Conserve :
    - Bold et Italic (propriétés de runs)
    - Le style du paragraphe
    """
    if 'properties' in para:
        props = para['properties']
        # Supprimer les propriétés de liste
        props.pop('ilvl', None)
        props.pop('numId', None)
        # Supprimer les propriétés de formatage qui outrepassent le style
        props.pop('size', None)
        props.pop('alignment', None)
        props.pop('color', None)
        props.pop('font', None)
        # Conserver le style et les autres propriétés essentielles

    # Nettoyer aussi les runs : garder uniquement bold et italic
    if 'runs' in para:
        for run in para['runs']:
            if 'properties' in run:
                run_props = run['properties']
                # Garder bold et italic
                kept_props = {}
                if 'bold' in run_props:
                    kept_props['bold'] = run_props['bold']
                if 'italic' in run_props:
                    kept_props['italic'] = run_props['italic']
                # Remplacer les propriétés du run par les propriétés conservées
                run['properties'] = kept_props

def insert_text_xp_tables(data: Dict[str, Any]) -> None:
    """
    Insère le texte dans les cellules des tableaux.

    Logique:
    1. Pour chaque table AUTO professional_experience:
    2. Lire jusqu'à 3 paragraphes après (chercher table EXISTING ou paragraphes)
    3. Si on rencontre une table EXISTING: l'extraire et la marquer pour suppression
    4. Fusionner contenu extrait + paragraphes lus
    5. Distribuer dans les cellules
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
                        if any(keyword in text.lower() for keyword in KEYWORDS_TECHNICAL_SKILLS):
                            break

                        # ARRÊTER si contexte
                        if 'contexte' in text.lower():
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
                        clean_paragraph_for_table(max_size_para)
                        element['rows'][0]['cells'][0]['paragraphs'] = [max_size_para]

                    # Trouver date (contient "20")
                    date_para = None
                    for para in remaining:
                        text = get_text_from_element(para)
                        if ' 20' in text or '/20' in text or '-20' in text:
                            date_para = para
                            break

                    if date_para and date_para in remaining:
                        remaining.remove(date_para)
                        clean_paragraph_for_table(date_para)
                        element['rows'][0]['cells'][1]['paragraphs'] = [date_para]

                    # Placer le reste dans cell[1][0]
                    # Filtrer: garder seulement les paragraphes avec du texte (exclure vides + page_break-only)
                    if remaining:
                        filtered_paras = []
                        for para in remaining:
                            clean_paragraph_for_table(para)

                            # Un paragraphe est utile s'il a au moins un run avec du texte
                            runs = para.get('runs', [])
                            has_meaningful_content = any('text' in run for run in runs)

                            if has_meaningful_content:
                                filtered_paras.append(para)

                        # Placer seulement les paragraphes significatifs
                        if filtered_paras:
                            element['rows'][1]['cells'][0]['paragraphs'] = filtered_paras

        i += 1

    # Supprimer en allant de la fin vers le début pour préserver les indices
    for idx in sorted(indices_to_remove, reverse=True):
        if idx < len(content):
            del content[idx]

    # Ajouter 1 paragraphe vide avant et 1 après chaque table AUTO professional_experience
    # (après les traitements/suppressions, pour qu'ils ne soient pas relus)
    new_content = []
    for element in content:
        if element.get('type') == 'Table' and element.get('auto_generated'):
            section = element.get('properties', {}).get('section')
            if section == 'professional_experience':
                # Ajouter 1 paragraphe vide juste avant la table
                new_content.append({'type': 'Paragraph', 'properties': {}, 'runs': []})
        new_content.append(element)
        if element.get('type') == 'Table' and element.get('auto_generated'):
            section = element.get('properties', {}).get('section')
            if section == 'professional_experience':
                # Ajouter 1 paragraphe vide juste après la table
                new_content.append({'type': 'Paragraph', 'properties': {}, 'runs': []})

    data['document']['content'] = new_content


def apply_styles_in_json(data: Dict[str, Any]) -> None:
    """
    Applique les styles par défaut dans les données JSON.
    Modifie in-place.

    Args:
        data (Dict): Structure JSON à modifier
    """
    # Appliquer les styles des titres
    for itag in data.get('document', {}).get('content', []):
        if 'tags' not in itag:
            continue
        tags = itag['tags']
        text = get_text_from_element(itag).lower() if itag else ""
        props = itag.get('properties', {})

        if 'header' in tags and any(keyword in text for keyword in KEYWORDS_HEADER_DOCUMENT):
            props['style'] = 'DC_H_DC'
        elif 'main_skills' in tags and any(keyword in text for keyword in KEYWORDS_MAIN_SKILLS):
            props['style'] = 'DC_T1_Sections'
        elif 'education' in tags and any(keyword in text for keyword in KEYWORDS_EDUCATION):
            props['style'] = 'DC_T1_Sections'
        elif 'professional_experience' in tags and any(keyword in text for keyword in KEYWORDS_PROFESSIONAL_EXPERIENCE):
            props['style'] = 'DC_T1_Sections'

        if 'header' in tags and 'DC_H_DC' not in props.get('style', '') and len(text) > 0 and len(text) <= 5:
            props['style'] = 'DC_H_Trigramme'
        elif 'header' in tags and 'DC_H_DC' not in props.get('style', '') and any(keyword in text for keyword in KEYWORDS_HEADER_EXPERIENCE) and len(text) > 5:
            props['style'] = 'DC_H_XP'

        if 'header' in tags and 'DC_H_XP' not in props.get('style', '') and 'DC_H_DC' not in props.get('style', '') and len(text) > 5:
            props['style'] = 'DC_H_Poste'

    for itable in data.get('document', {}).get('content', []):
        if itable.get('type') == 'Table' and 'properties' in itable:
            section = itable.get('properties', {}).get('section')
            if section == 'education':
                itable['properties']['style'] = 'DC_Table_Education'
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
                                    para['properties']['style'] = 'DC_Normal'

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
        if not ilvl:
            continue
        elif ilvl == "0":
            props['style'] = 'DC_1st_bullet'
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
        if 'style' not in props or 'style' in props and props['style'] == 'Normal':
            props['style'] = 'DC_Normal'

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

    # Appliquer les tags de section
    apply_section_tags(data)

    # Créer le header "Langues" juste avant le premier keyword détecté
    create_language_header(data)

    # Créer et restructurer les tables de la section education
    create_edu_tables(data, page_dims=page_dims)

    # Appliquer les tables et dimensions
    create_xp_tables(data, page_dims=page_dims)

    # Insérer le texte dans les tables créées
    insert_text_xp_tables(data)

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
