"""
Script simplifié pour extraire un fichier XML global en JSON avec tous les détails.
Entrée: fichier global.xml
Sortie: fichier _balises.json
"""

from argparse import ArgumentParser
import xml.etree.ElementTree as ET
import json
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
KEYWORDS_PROFESSIONAL_EXPERIENCE = ["expérience professionnelle", "experience professionnelle", "expériences professionnelles", "experience professionnelles"]

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


def parse_table(table, ns: Dict, index: int) -> Dict[str, Any]:
    """Parse un tableau"""
    table_obj = {
        'index': index,
        'type': 'Table',
        'rows': []
    }
    
    rows = table.findall('w:tr', ns)
    table_obj['row_count'] = len(rows)
    
    for row_idx, row in enumerate(rows):
        row_obj = {
            'row_index': row_idx,
            'cells': []
        }
        
        cells = row.findall('w:tc', ns)
        for col_idx, cell in enumerate(cells):
            cell_obj = {
                'col_index': col_idx,
                'paragraphs': []
            }
            
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

def apply_styles_in_json(data: Dict[str, Any]) -> None:
    """
    Applique les styles par défaut dans les données JSON.
    Modifie in-place.
    
    Args:
        data (Dict): Structure JSON à modifier
    """
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
        elif 'header' in tags and 'DC_H_DC' not in props.get('style', '') and 'experience' in text and len(text) > 5:
            props['style'] = 'DC_H_XP'
        
        if 'header' in tags and 'DC_H_XP' not in props.get('style', '') and 'DC_H_DC' not in props.get('style', '') and len(text) > 5:
            props['style'] = 'DC_H_Poste'

def apply_tags_and_styles(raw_json_file: str, output_dir: str = None) -> str:
    """
    Charge un JSON brut, applique les tags de section et les styles,
    puis enregistre le résultat transformé.
    
    Args:
        raw_json_file (str): Chemin du fichier JSON RAW
        output_dir (str): Répertoire de sortie (optionnel, défaut: output/)
        
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
    
    # Appliquer les tags de section
    apply_section_tags(data)
    
    # Appliquer les styles
    apply_styles_in_json(data)
    
    # Sauvegarder le JSON transformé
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"✅ {output_file.name} créé (avec tags et styles)")
    
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
