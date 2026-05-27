"""
Script simplifié pour extraire un fichier XML global en JSON avec tous les détails.
Entrée: fichier _GLOBAL.xml
Sortie: fichier _GLOBAL_raw.json (xml brut traduit en json)
Sortie: fichier _GLOBAL_transformed.json (après taggings et transformations)

ORGANISATION DES FONCTIONS:
1. Imports et Constantes
2. Fonctions Utilitaires (text extraction, regex, helpers)
3. Détection de Sections & Core Tagging
4. Section HEADER
5. Section MAIN SKILLS
6. Section EDUCATION
7. Section PROFESSIONAL EXPERIENCE
8. Final Processing & Rendering (nettoyage, styles, indices)
9. Main Entry Point (apply_tags_and_styles, main)
"""

from argparse import ArgumentParser
import json
import re
import sys
from pathlib import Path
from typing import Dict, Any, List, Optional

# Ajouter le répertoire parent à sys.path pour les imports
parent_dir = str(Path(__file__).parent.parent)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

# Imports locaux
try:
    from .parse_template import extract_page_dimensions_from_template
except (ImportError, ValueError):
    # Fallback pour exécution directe (python3 script.py)
    from tools.parse_template import extract_page_dimensions_from_template

# MARK: CONFIGURATION & CONSTANTES
# ===== CONSTANTES =====
TEMPLATE_PATH = 'TEMPLATE/TEMPLATE.docx'

# ===== NAMESPACES =====
NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
}

KEYWORDS_HEADER_DOCUMENT = ["dossier de compétences", "dossier de competence", "dossier de competences", "dossier de competences"]
KEYWORDS_HEADER_EXPERIENCE = ["expérience", "experience"]
KEYWORDS_MAIN_SKILLS = ["domaine de compétence", "domaine de competence", "domaines de compétence", "domaines de competence", "compétences principales", "competences principales", "compétence", "competence"]
KEYWORDS_EDUCATION = ["formation", "formations", "certifications", "certification", "langue", "langues", "diplôme", "diplome", "diplômes", "diplomes"]
KEYWORDS_LANGUAGES = ["langue", "langues", "français", "anglais", "espagnol", "allemand", "italien", "chinois", "japonais", "russe"]
KEYWORDS_PROFESSIONAL_EXPERIENCE = ["expérience professionnelle", "experience professionnelle", "expériences professionnelles", "experience professionnelles"]
KEYWORDS_TECHNICAL_SKILLS = ["techniques", "technique", "informatiques", "informatique", "numériques", "numeriques", "numérique", "numerique"]
XP_DATE_PATTERN = r'(?:depuis\s+|du\s+|de\s+|à\s+partir\s+de\s+)?(?:\d{1,2}(?:\s*[/–-]\s*\d{1,2})?\s*[/–-]\s*\d{2,4}|\d{4}\s*[–-]\s*\d{4})'
MAX_XP_DESCRIPTION_LENGTH = 60  # Limite de caractères pour la description d'une expérience professionnelle
SINGLE_XP_DATE_PATTERN = r'^\d{1,2}(?:\s*[/–-]\s*\d{1,2})?\s*[/–-]\s*\d{2,4}$'

# MARK: FONCTIONS UTILITAIRES
# ===== 2. FONCTIONS UTILITAIRES =====

def get_table_widths_for_section(section: str = None, page_dims: dict = None) -> tuple:
    """
    Calcule les largeurs des colonnes pour les 2 types de tables.

    Type 1 (education): Col1 = 3cm (étroite), Col2 = reste (large)
    Type 2 (professional): Col1 = reste (large), Col2 = 5cm (étroite)

    Args:
        section (str): 'education', 'professional_experience' ou None
        page_dims (dict): Dimensions de page avec left_margin, right_margin, page_width

    Returns:
        tuple: (col_fixed_width, usable_width, type1_widths, type2_widths)
               où type1_widths et type2_widths sont des tuples (col1, col2)
    """
    if page_dims is None:
        raise ValueError("page_dims doit être fourni")

    col_fixed_width_3 = page_dims['col_fixed_width_3']
    col_fixed_width_5 = page_dims['col_fixed_width_5']
    usable_width = page_dims['usable_width']

    # Type 1 (Education): Col1 = 3cm, Col2 = reste
    edu_table_col1 = col_fixed_width_3
    edu_table_col2 = usable_width - col_fixed_width_3

    # Type 2 (Professional): Col1 = reste, Col2 = 5cm
    xp_table_col1 = usable_width - col_fixed_width_5
    xp_table_col2 = col_fixed_width_5

    # Type default: Col1 = Col2 = moitié de l'usable width
    default_table_col1 = usable_width // 2
    default_table_col2 = usable_width - default_table_col1

    # Retourner les largeurs appropriées selon la section
    if section == 'education':
        return (edu_table_col1, edu_table_col2)
    elif section == 'professional_experience':
        return (xp_table_col1, xp_table_col2)
    else:
        return (default_table_col1, default_table_col2)

def get_text_from_element(element: Dict[str, Any], lower: bool = True) -> str:
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

    text = ' '.join(texts)
    return text.lower() if lower else text

def get_raw_text_from_paragraph(para: Dict[str, Any]) -> str:
    """Retourne le texte brut d'un paragraphe sans forcer la casse."""
    return ''.join(run.get('text', '') for run in para.get('runs', [])) or para.get('text', '') or get_text_from_element(para, lower=False)

def match_xp_date(text: str) -> Optional[re.Match]:
    """Retourne un match de date XP n'importe où dans le texte (avec formats élargis)."""
    return re.search(XP_DATE_PATTERN, text, flags=re.IGNORECASE)

def is_single_xp_date(text: str) -> bool:
    """Retourne True si le texte ressemble a une date XP simple (sans prefixe)."""
    return re.match(SINGLE_XP_DATE_PATTERN, text.strip()) is not None

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

    # Copier xp_metadata si présent (important pour préserver les détections XP)
    if 'xp_metadata' in para:
        new_para['xp_metadata'] = para['xp_metadata'].copy()

    return new_para

# MARK: TABLE CREATION
def create_empty_table_2x2(index: int, row_height: int = 360,
                           col1_width: int = None, col2_width: int = None,
                           section: str = None,
                           auto_generated: bool = False,
                           num_rows: int = 2,
                           page_dims: dict = None) -> Dict[str, Any]:
    """
    Crée une table 2xN vide (sans paragraphes de remplissage) avec dimensions spécifiques par colonne.
    Les dimensions sont calculées selon la section et les marges du template, elles sont récupérées de la fonction get_table_widths_for_section.

    Args:
        index: Index dans le document
        row_height: Hauteur de la ligne en twips
        col1_width: Largeur colonne 1 (twips) - si None, utilisé section+page_dims
        col2_width: Largeur colonne 2 (twips) - si None, utilisé section+page_dims
        section: 'education' ou 'professional_experience' ou 'autre'
        auto_generated: Flag pour indiquer que la table a été créée automatiquement
        num_rows: Nombre de rows à créer (défaut 2)
        page_dims: Dictionnaire avec dimensions de page (col_fixed_width_3, col_fixed_width_5, usable_width, etc.)

    Returns:
        Table 2xN structurée avec dimensions (sans paragraphes vides)
    """

    if col1_width is None or col2_width is None:
        if section == 'education':
            col1_width, col2_width = get_table_widths_for_section('education', page_dims)
        elif section == 'professional_experience':
            col1_width, col2_width = get_table_widths_for_section('professional_experience', page_dims)
        else:
            col1_width, col2_width = get_table_widths_for_section(None, page_dims)

    table_total_width = col1_width + col2_width

    # Définir les bordures selon la section
    if section == 'professional_experience':
        # Tables expériences professionelles : seulement bottom border
        borders = {
            'top': None,
            'bottom': {'size': '10', 'color': '000000'},
            'left': None,
            'right': None,
            'insideH': None,
            'insideV': None
        }
    else:
        # Par défaut, pas de bordures
        borders = {
            'top': None,
            'bottom': None,
            'left': None,
            'right': None,
            'insideH': None,
            'insideV': None
        }

    # Créer les rows dynamiquement
    rows = []
    for row_idx in range(num_rows):
        row = {
            'row_index': row_idx,
            'height': row_height,
            'cells': [
                {
                    'col_index': 0,
                    'width': col1_width,
                    'properties': {
                        'hAlign': 'center' if section == 'main_skills' else 'left',
                        'vAlign': 'center'
                    },
                    'paragraphs': []
                },
                {
                    'col_index': 1,
                    'width': col2_width,
                    'properties': {
                        'hAlign': 'right' if section == 'professional_experience' else 'left',
                        'vAlign': 'center'
                    },
                    'paragraphs': []
                }
            ]
        }
        rows.append(row)

    # Retourner juste la table (dictionnaire), pas un tuple
    return {
        'index': index + 1,
        'type': 'Table',
        'auto_generated': auto_generated,
        'properties': {
            'table_width': str(table_total_width),
            'table_width_type': 'dxa',
            'section': section,
            'borders': borders,
            'style': "DC_Table_Content"
        },
        'tags': [section] if section else [],
        'row_count': num_rows,
        'col_count': 2,
        'rows': rows
    }

# MARK: DÉTECTION DE SECTIONS & CORE TAGGING
# ===== 3. DÉTECTION DE SECTIONS & CORE TAGGING =====

def is_promotable_section_title(element: Dict[str, Any], keywords: List[str]) -> bool:
    """Retourne True si l'élément ressemble à un vrai titre de section.

    On exclut tous les paragraphes de liste (`ilvl`) pour éviter qu'un mot-clé
    présent dans une puce soit reclassé en `T1_Sections`.

    ⚠️ EXCLUSION: Ne pas considérer les XP entries (paragraphes commençant par une DATE)
    comme des titres - ce sont des données à splitter.
    """
    if not element or element.get('type') != 'Paragraph':
        return False

    props = element.get('properties', {})
    if props.get('ilvl') is not None:
        return False

    text = get_text_from_element(element)
    if not text.strip():
        return False

    tags = element.get('tags', [])
    if isinstance(tags, str):
        tags = [tags]

    if 'professional_experience' in tags and not any(keyword in text for keyword in KEYWORDS_PROFESSIONAL_EXPERIENCE):
        return False

    if 'professional_experience' in tags and text.strip().startswith(('projet', 'projets')):
        return False

    style = props.get('style', '')

    # Exclure les XP entries: paragraphes commençant par une DATE (regex)
    if match_xp_date(text):
        return False

    if element.get('auto_generated'):
        return True

    if style.startswith('Titre') or style.startswith('Heading'):
        return True

    if style in {'DC_T1_Sections', 'DC_XP_Title', 'DC_H_DC', 'DC_H_XP', 'DC_H_Poste'}:
        return True

    return any(keyword in text for keyword in keywords)

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

        # EXCLUSION: Ne pas déterminer une nouvelle section si le paragraphe a un ilvl
        # (c'est un bullet point, pas un titre de section)
        props = element.get('properties', {})
        has_ilvl = props.get('ilvl') is not None

        # Détecter si cet élément déclenche un changement de section
        # Mais UNIQUEMENT si ce n'est pas un bullet point (ilvl)
        detected_section = None
        if not has_ilvl:
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

def apply_section_header_styles(data: Dict[str, Any]) -> None:
    """
    Applique le style DC_T1_Sections aux vrais headers de section.

    Utilise is_promotable_section_title() pour la détection, garantissant une seule
    source de vérité pour identifier les vrais titres (vs paragraphes ordinaires).

    EXCLUSION DURE:
    - Les paragraphes contenant "langue maternelle" ne sont JAMAIS traités comme des headers
    - Les paragraphes avec xp_split_part (DATE, COMPANY, POSTE) gardent leurs styles propres
    """
    content = data.get('document', {}).get('content', [])

    # Tous les keywords de section à chercher
    section_keywords = (KEYWORDS_EDUCATION + KEYWORDS_PROFESSIONAL_EXPERIENCE +
                       KEYWORDS_MAIN_SKILLS + KEYWORDS_HEADER_DOCUMENT + KEYWORDS_HEADER_EXPERIENCE)

    for element in content:
        # Check dur: exclure "langue maternelle" absolument
        text = get_text_from_element(element)
        if 'langue maternelle' in text.lower():
            continue

        # Ne pas toucher aux paragraphes avec xp_split_part (ils ont leurs propres styles)
        if element.get('xp_split_part'):
            continue

        # Utiliser la fonction existante pour vérifier si c'est un vrai titre
        if is_promotable_section_title(element, section_keywords):
            if 'properties' not in element:
                element['properties'] = {}
            element['properties']['style'] = 'DC_T1_Sections'

# MARK: SECTION HEADER
# ===== 4. SECTION HEADER =====

# Les fonctions spécifiques au header iront ici si nécessaire
# (actuellement gérées par apply_section_tags et apply_section_header_styles)

# MARK: SECTION MAIN SKILLS
# ===== 5. SECTION MAIN SKILLS =====

def create_main_skills_table (data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Crée les structures des tables main skills uniquement quand une table est présente au départ dans la source

    Responsabilité:
    - Itérer sur le titre de main skills contenant KEYWORDS_MAIN_SKILLS
    - la section devrait être unique, créer une table vide
    - Marquer les sources pour suppression

    Args:
        data: Structure du document JSON contenant page_dimensions

    Returns:
        Dict contenant les métadonnées sur les tables créées
    """

    # Récupérer les dimensions de page depuis data
    page_dims = data.get('page_dimensions')
    if page_dims is None:
        raise ValueError("page_dimensions non trouvées dans data")

    content = data.get('document', {}).get('content', [])
    result = {'tables_created': []}

    # Trouver le header main skills
    main_skills_header_idx = None
    for i, elem in enumerate(content):
        if elem.get('type') == 'Paragraph':
            text = get_text_from_element(elem)
            style = elem.get('properties', {}).get('style', '')
            is_section_header = style == 'DC_T1_Sections'

            if any(keyword in text.lower() for keyword in KEYWORDS_MAIN_SKILLS) and is_section_header:
                main_skills_header_idx = i
                break

    if main_skills_header_idx is not None:
        # Chercher la première table source après le header main skills
        source_table_idx = None
        j = main_skills_header_idx + 1

        while j < len(content):
            elem = content[j]

            # Stop si on rencontre un autre header ou une autre section avant une table source
            if elem.get('type') == 'Paragraph':
                text = get_text_from_element(elem)
                style = elem.get('properties', {}).get('style', '')
                is_section_header = style == 'DC_T1_Sections'

                if is_section_header and any(keyword in text.lower() for keyword in KEYWORDS_EDUCATION + KEYWORDS_PROFESSIONAL_EXPERIENCE):
                    break

            if elem.get('type') == 'Table' and not elem.get('auto_generated'):
                source_table_idx = j
                break

            j += 1

        # Créer la table uniquement si une vraie table source existe
        if source_table_idx is not None:
            source_table = content[source_table_idx]
            source_rows = source_table.get('rows', [])

            # Créer une table vide avec les largeurs par défaut de create_empty_table_2x2
            new_table = create_empty_table_2x2(
                source_table_idx,
                section='main_skills',
                auto_generated=True,
                num_rows=len(source_rows),
                page_dims=page_dims
            )

            new_table['properties']['table_width'] = str(page_dims['usable_width'])
            new_table['properties']['table_width_type'] = 'dxa'

            content.insert(source_table_idx, new_table)
            result['tables_created'].append({
                'index': source_table_idx,
                'source_table_index': source_table_idx + 1,
                'row_count': len(source_rows)
            })

    return result

def insert_text_main_skills_table(data: Dict[str, Any], creation_result: Dict[str, Any], page_dims: dict = None) -> None:
    """
    Remplit le contenu de la table main skills et supprime les sources.

    Responsabilité:
    - Pour la table main skills créée
    - Remplir avec le contenu (paragraphes ou tables existantes)
    - Supprimer les sources après migration

    Args:
        data: Structure du document JSON
        creation_result: Résultat de create_skills_table() contenant les indices des tables créées
    """
    content = data.get('document', {}).get('content', [])
    indices_to_remove = []

    # Itérer sur toutes les tables auto main skills créées
    for i, elem in enumerate(content):
        if elem.get('type') == 'Table' and elem.get('auto_generated') and elem.get('properties', {}).get('section') == 'main_skills':
            rows = elem.get('rows', [])

            # La table source suit immédiatement la table auto créée
            source_idx = None
            for j in range(i + 1, len(content)):
                next_elem = content[j]

                if next_elem.get('type') == 'Table' and not next_elem.get('auto_generated'):
                    source_idx = j
                    break

                if next_elem.get('type') == 'Paragraph':
                    text = get_text_from_element(next_elem)
                    style = next_elem.get('properties', {}).get('style', '')
                    is_section_header = style == 'DC_T1_Sections'

                    if is_section_header and any(keyword in text.lower() for keyword in KEYWORDS_EDUCATION + KEYWORDS_PROFESSIONAL_EXPERIENCE):
                        break

            if source_idx is None:
                continue

            existing_table = content[source_idx]
            source_rows = existing_table.get('rows', [])

            # Copier chaque bloc source dans la ligne correspondante de la nouvelle table
            for row_idx, source_row in enumerate(source_rows):
                if row_idx >= len(rows):
                    break

                target_row = rows[row_idx]

                source_cells = source_row.get('cells', [])
                target_cells = target_row.get('cells', [])

                for cell_idx, target_cell in enumerate(target_cells):
                    source_cell = source_cells[cell_idx] if cell_idx < len(source_cells) else None
                    if source_cell is None:
                        target_cell['paragraphs'] = []
                        continue

                    if 'properties' in source_cell:
                        target_cell['properties'] = source_cell['properties'].copy()

                    cloned_paragraphs = []
                    for para in source_cell.get('paragraphs', []):
                        cloned_paragraphs.append(clone_paragraph_clean(para))

                    target_cell['paragraphs'] = cloned_paragraphs

            elem['row_count'] = len(rows)
            if page_dims:
                elem.setdefault('properties', {})['table_width'] = str(page_dims['usable_width'])
            indices_to_remove.append(source_idx)

    # Supprimer les tables source en ordre inverse
    for idx in sorted(indices_to_remove, reverse=True):
        if idx < len(content):
            del content[idx]

    data['document']['content'] = content

# MARK: SECTION EDUCATION
# ===== 6. SECTION EDUCATION =====

def create_language_header(data: Dict[str, Any]) -> None:
    """
    Crée un header "Langues" juste avant le premier élément contenant KEYWORDS_LANGUAGES,
    si ce header n'existe pas déjà.

    Args:
        data: Structure du document JSON
    """
    content = data.get('document', {}).get('content', [])

    # D'abord, vérifier si un header "Langues" existe déjà dans le document
    # Cherche dans les styles Heading/Titre (avant apply_section_header_styles) ET DC_T1_Sections (après apply_section_header_styles)
    header_langues_exists = False
    for element in content:
        if element.get('type') == 'Paragraph':
            text = get_text_from_element(element).strip()
            style = element.get('properties', {}).get('style', '')
            # Check: c'est un vrai header "Langues" (pas "Français langue maternelle")
            if style.startswith('Heading') or style.startswith('Titre') or style == 'DC_T1_Sections':
                if text.lower() == 'langues':
                    header_langues_exists = True
                    break
                if text.lower() == 'langue':
                    element['runs'] = [{'text': 'Langues', 'properties': {}}]
                    header_langues_exists = True
                    break

    if header_langues_exists:
        return  # Header "Langues" existe déjà, rien à faire

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

    # Créer et insérer le header "Langues" juste avant le premier keyword
    new_header = {
        'type': 'Paragraph',
        'runs': [{'text': 'Langues', 'properties': {}}],
        'properties': {},
        'tags': ['education'],
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
    text = ''.join(run.get('text', '') for run in para.get('runs', [])) or para.get('text', '') or get_text_from_element(para)
    normalized_text = text.strip().lower()

    # Ne pas splitter sur les mots-clés génériques si le paragraphe ne porte pas
    # réellement une langue: ils servent surtout au header "Langues".
    split_keywords = [kw for kw in KEYWORDS_LANGUAGES if kw not in {'langue', 'langues'}]

    lang_keyword = None
    lang_pos = len(text)
    for keyword in split_keywords:
        match = re.search(rf'(?<!\w){re.escape(keyword)}(?!\w)', normalized_text)
        if match and match.start() < lang_pos:
            lang_keyword = keyword
            lang_pos = match.start()

    if lang_keyword is None:
        return [para]

    # Séparer le label de langue et sa description au premier séparateur utile.
    split_end = lang_pos + len(lang_keyword)
    colon_pos = text.find(':', split_end)
    if colon_pos != -1:
        split_end = colon_pos

    lang_text = text[:split_end].rstrip(' :\u00a0\t').strip()
    desc_text = text[split_end:].lstrip(' :\u00a0\t').strip()

    if not lang_text:
        return [para]

    first_run_props = para.get('runs', [{}])[0].get('properties', {}) if para.get('runs') else {}
    result = []

    lang_para = clone_paragraph_clean(para)
    lang_para['runs'] = [{"text": lang_text, "properties": first_run_props}]
    result.append(lang_para)

    if desc_text:
        desc_para = clone_paragraph_clean(para)
        desc_para['runs'] = [{"text": desc_text, "properties": first_run_props}]
        result.append(desc_para)

    return result

def group_education_paragraphs(paragraphs: List[Dict[str, Any]]) -> List[List[Dict[str, Any]]]:
    """
    Groupe les paragraphes éducation en blocs basés sur les dates.

    Logique:
    - Un bloc commence avec une année ou date détectée (par exemple 1996, 2002, 2002-2003)
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

        is_date = re.search(r'(?<!\d)(?:19|20)\d{2}(?!\d)', para_text) is not None

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

def create_edu_table(data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Crée les structures des tables éducation de manière générique.

    Cherche toutes les sections éducation dans le document (formations et langues)
    et crée une table pour chaque section qui contient du contenu.

    Responsabilité:
    - Itérer sur tous les titres contenant KEYWORDS_EDUCATION
    - Pour chaque section trouvée (formations, langues, etc.), créer une table vide
    - Marquer les sources pour suppression

    Args:
        data: Structure du document JSON contenant page_dimensions

    Returns:
        Dict contenant les métadonnées sur les tables créées
    """
    # Récupérer les dimensions de page depuis data
    page_dims = data.get('page_dimensions')
    if page_dims is None:
        raise ValueError("page_dimensions non trouvées dans data")

    content = data.get('document', {}).get('content', [])
    result = {'tables_created': []}

    # Trouver tous les headers éducation
    edu_headers = []
    for i, elem in enumerate(content):
        if elem.get('type') == 'Paragraph':
            text = get_text_from_element(elem)
            style = elem.get('properties', {}).get('style', '')
            is_section_header = style == 'DC_T1_Sections'
            is_auto_language_header = elem.get('auto_generated') and text.strip().lower() == 'langues'

            # Chercher si c'est un header éducation
            if ((any(keyword in text.lower() for keyword in KEYWORDS_EDUCATION) and is_section_header)
                    or is_auto_language_header):
                # Déterminer le type: formations ou langues
                if any(kw in text.lower() for kw in ['formation', 'formations', 'diplôme', 'diplômes', 'certification', 'certifications']):
                    edu_type = 'formations'
                elif any(kw in text.lower() for kw in ['langue', 'langues']):
                    edu_type = 'langues'
                else:
                    edu_type = 'education'  # Type générique

                edu_headers.append({'index': i, 'text': text, 'type': edu_type})

    # Créer une table pour chaque section éducation trouvée
    # Traiter en ORDRE INVERSE pour que les insertions ne modifient pas les indices des headers suivants
    for header_info in reversed(edu_headers):
        header_idx = header_info['index']
        edu_type = header_info['type']

        # Vérifier qu'il y a du contenu après le header
        has_content = False
        content_indices = []
        j = header_idx + 1

        while j < len(content):
            elem = content[j]

            # Stop si on rencontre un autre header ou une autre section
            if elem.get('type') == 'Paragraph':
                text = get_text_from_element(elem)
                style = elem.get('properties', {}).get('style', '')
                is_section_header = style == 'DC_T1_Sections'

                if is_section_header and any(keyword in text.lower() for keyword in KEYWORDS_EDUCATION + KEYWORDS_PROFESSIONAL_EXPERIENCE):
                    break
                if not text.strip():
                    j += 1
                    continue
                has_content = True
                content_indices.append(j)
            elif elem.get('type') == 'Table':
                has_content = True
                content_indices.append(j)
                break

            j += 1

        # Créer la table uniquement s'il y a du contenu
        if has_content and content_indices:
            insert_idx = header_idx + 1

            # Créer une table vide pour cette section
            new_table = create_empty_table_2x2(
                insert_idx,
                section='education',
                auto_generated=True,
                num_rows=0,  # Commencer avec 0 rows, sera remplies par insert_text_edu_table
                page_dims=page_dims
            )
            new_table['edu_type'] = edu_type

            content.insert(insert_idx, new_table)
            result['tables_created'].append({'index': insert_idx, 'type': edu_type})

    data['document']['content'] = content
    return result

def insert_text_edu_table(data: Dict[str, Any], creation_result: Dict[str, Any], page_dims: dict = None) -> None:
    """
    Remplit le contenu des tables éducation (formations et langues) et supprime les sources.

    Responsabilité:
    - Pour chaque table auto éducation créée
    - Remplir avec le contenu (paragraphes ou tables existantes)
    - Supprimer les sources après migration

    Args:
        data: Structure du document JSON
        creation_result: Résultat de create_edu_table() contenant les indices des tables créées
        page_dims: Dimensions de page (requis pour calculer largeurs de colonnes)
    """
    content = data.get('document', {}).get('content', [])
    indices_to_remove = []

    # Itérer sur toutes les tables auto education créées
    for i, elem in enumerate(content):
        if elem.get('type') == 'Table' and elem.get('auto_generated'):
            section = elem.get('properties', {}).get('section')
            edu_type = elem.get('edu_type')

            if section == 'education' and edu_type:
                rows = elem.get('rows', [])

                # ===== FORMATIONS (diplômes, certifications, etc.) =====
                if edu_type in ['formations', 'diplomes', 'certifications']:
                    # Chercher la table source suivante, même si elle n'est pas juste après le header
                    j = i + 1
                    existing_table = None
                    while j < len(content):
                        next_elem = content[j]

                        if next_elem.get('type') == 'Table' and not next_elem.get('auto_generated'):
                            existing_table = next_elem
                            break

                        if next_elem.get('type') == 'Paragraph':
                            text = get_text_from_element(next_elem)
                            style = next_elem.get('properties', {}).get('style', '')
                            is_section_header = style == 'DC_T1_Sections'

                            if is_section_header and any(keyword in text.lower() for keyword in KEYWORDS_EDUCATION + KEYWORDS_PROFESSIONAL_EXPERIENCE):
                                break

                        j += 1

                    if existing_table is not None:
                        source_rows = existing_table.get('rows', [])
                        source_col_count = existing_table.get('col_count', 0)

                        if source_rows and source_col_count == 2:
                            temp_table = create_empty_table_2x2(
                                0,  # index fictif
                                section='education',
                                auto_generated=True,
                                num_rows=len(source_rows),
                                page_dims=page_dims
                            )
                            rows = temp_table['rows']

                            for row_idx, source_row in enumerate(source_rows):
                                if row_idx >= len(rows):
                                    break

                                source_cells = source_row.get('cells', [])
                                target_cells = rows[row_idx].get('cells', [])

                                for cell_idx, target_cell in enumerate(target_cells):
                                    source_cell = source_cells[cell_idx] if cell_idx < len(source_cells) else None
                                    if source_cell is None:
                                        target_cell['paragraphs'] = []
                                        continue

                                    cloned_paragraphs = [clone_paragraph_clean(para) for para in source_cell.get('paragraphs', [])]
                                    target_cell['paragraphs'] = cloned_paragraphs

                            elem['rows'] = rows
                            elem['row_count'] = len(rows)
                        else:
                            all_paras = []

                            # Extraire tous les paragraphes
                            for row in existing_table.get('rows', []):
                                for cell in row.get('cells', []):
                                    all_paras.extend(cell.get('paragraphs', []))

                            # Grouper par blocs
                            blocks = group_education_paragraphs(all_paras)

                            # Créer les rows avec le nombre exact requis
                            if blocks:
                                # Utiliser create_empty_table_2x2 pour générer les rows avec le bon nombre
                                temp_table = create_empty_table_2x2(
                                    0,  # index fictif
                                    section='education',
                                    auto_generated=True,
                                    num_rows=len(blocks),
                                    page_dims=page_dims
                                )
                                rows = temp_table['rows']

                                # Remplir chaque row avec les blocs
                                for row_idx, block in enumerate(blocks):
                                    for para_idx, para in enumerate(block):
                                        cloned = clone_paragraph_clean(para)
                                        if para_idx == 0:
                                            rows[row_idx]['cells'][0]['paragraphs'].append(cloned)
                                        else:
                                            rows[row_idx]['cells'][1]['paragraphs'].append(cloned)

                                elem['rows'] = rows
                                elem['row_count'] = len(rows)

                        # Marquer pour suppression
                        indices_to_remove.append(j)

                # ===== LANGUES =====
                elif edu_type == 'langues':
                    j = i + 1
                    lang_table = elem
                    existing_table = None
                    source_paragraphs = []
                    source_indices = []

                    while j < len(content):
                        next_elem = content[j]

                        if next_elem.get('type') == 'Table' and not next_elem.get('auto_generated'):
                            existing_table = next_elem
                            break

                        if next_elem.get('type') == 'Paragraph':
                            text = get_text_from_element(next_elem)

                            # Stop si autre section
                            if any(keyword in text.lower() for keyword in KEYWORDS_PROFESSIONAL_EXPERIENCE):
                                break
                            if text.strip() and any(kw in text.lower() for kw in ['formation', 'formations', 'diplôme', 'certification']):
                                break

                            if text.strip():
                                source_paragraphs.append(next_elem)
                                source_indices.append(j)

                        elif next_elem.get('type') == 'Table':
                            break

                        j += 1

                    if existing_table is not None:
                        source_rows = existing_table.get('rows', [])

                        # Utiliser create_empty_table_2x2 pour générer les rows avec le bon nombre
                        temp_table = create_empty_table_2x2(
                            0,  # index fictif
                            section='education',
                            auto_generated=True,
                            num_rows=len(source_rows),
                            page_dims=page_dims
                        )
                        rows = temp_table['rows']

                        # Remplir chaque row en recopiant exactement la table source
                        for row_idx, source_row in enumerate(source_rows):
                            if row_idx >= len(rows):
                                break

                            source_cells = source_row.get('cells', [])
                            target_cells = rows[row_idx].get('cells', [])

                            source_cell_0 = source_cells[0] if len(source_cells) > 0 else None
                            source_cell_1 = source_cells[1] if len(source_cells) > 1 else None

                            for cell_idx, target_cell in enumerate(target_cells):
                                if cell_idx >= len(source_cells):
                                    target_cell['paragraphs'] = []
                                    continue

                                source_cell = source_cells[cell_idx]
                                cloned_paragraphs = [clone_paragraph_clean(para) for para in source_cell.get('paragraphs', [])]
                                target_cell['paragraphs'] = cloned_paragraphs

                            # Si la description est collee au keyword en col0, la splitter vers col1
                            if source_cell_0 and source_cell_1:
                                cell0_paras = source_cell_0.get('paragraphs', [])
                                cell1_paras = source_cell_1.get('paragraphs', [])
                                if cell0_paras and not cell1_paras:
                                    split_parts = split_paragraph_at_language(cell0_paras[0])
                                    if split_parts and len(split_parts) > 1:
                                        target_cells[0]['paragraphs'] = [clone_paragraph_clean(split_parts[0])]
                                        target_cells[1]['paragraphs'] = [clone_paragraph_clean(split_parts[1])]

                        lang_table['rows'] = rows
                        lang_table['row_count'] = len(rows)

                        # Marquer la source pour suppression
                        indices_to_remove.append(j)
                    elif source_paragraphs:
                        split_rows = []

                        for para in source_paragraphs:
                            split_parts = split_paragraph_at_language(para)

                            if not split_parts:
                                continue

                            lang_para = split_parts[0]
                            desc_para = split_parts[1] if len(split_parts) > 1 else None
                            split_rows.append((lang_para, desc_para))

                        if split_rows:
                            temp_table = create_empty_table_2x2(
                                0,
                                section='education',
                                auto_generated=True,
                                num_rows=len(split_rows),
                                page_dims=page_dims
                            )
                            rows = temp_table['rows']

                            for row_idx, (lang_para, desc_para) in enumerate(split_rows):
                                rows[row_idx]['cells'][0]['paragraphs'] = [clone_paragraph_clean(lang_para)]
                                rows[row_idx]['cells'][1]['paragraphs'] = [clone_paragraph_clean(desc_para)] if desc_para else []

                            lang_table['rows'] = rows
                            lang_table['row_count'] = len(rows)

                            # Marquer les sources pour suppression
                            indices_to_remove.extend(source_indices)

    # Supprimer les sources (en ordre inverse pour éviter les décalages d'indices)
    for idx in sorted(indices_to_remove, reverse=True):
        if idx < len(content):
            del content[idx]

    data['document']['content'] = content

# MARK: SECTION PROFESSIONAL EXPERIENCE
# ===== 7. SECTION PROFESSIONAL EXPERIENCE =====

# ## XP Pattern Detection
def _detect_and_tag_xp_content(para: Dict[str, Any]) -> None:
    """
    Détecte et tague le contenu XP basé sur des patterns.
    Tagge le paragraphe avec les champs xp_* appropriés même s'il n'a pas été splité.

    Détection:
    - xp_date: contient une date (regex XP_DATE_PATTERN)
    - xp_entry_start: marqué True si c'est une xp_date (début d'une entry)
    - xp_company: NOM MAJUSCULE SEUL ou très court (pas de verbes, pas de conjonctions)
    - xp_poste: titre de poste (pattern spécifique: Verbe+Nom, titres métier)
    - xp_description: texte long (> MAX_XP_DESCRIPTION_LENGTH) contenant "contexte"

    Args:
        para: Paragraphe JSON à tagger (modifié in-place)
    """
    if not para.get('xp_split_part'):  # Ne pas retagger les paragraphes déjà splittés
        # Ne pas tagger les titres de section
        style = para.get('properties', {}).get('style', '')
        if style in ['DC_T1_Sections']:
            return

        text = get_raw_text_from_paragraph(para).strip()

        if not text:
            return

        # Initialiser xp_metadata s'il n'existe pas
        if 'xp_metadata' not in para:
            para['xp_metadata'] = {}

        # Détection xp_date (contient une date)
        if match_xp_date(text):
            para['xp_metadata']['detected_xp_date'] = True
            # Marquer aussi que c'est le début d'une entry - AJOUTER À xp_metadata directement
            para['xp_metadata']['xp_entry_start'] = True
            para['xp_entry_start'] = True  # Aussi en propriété directe pour traitement immédiat
            return  # Ne pas faire d'autres détections si c'est une date

        # Ne pas détecter company/poste dans les bullets/sous-points (ilvl défini)
        # car company/poste ne doivent être détectés qu'au niveau principal de l'entrée XP
        has_ilvl = para.get('properties', {}).get('ilvl') is not None

        if not has_ilvl:

            # Détection xp_description (texte long avec "contexte" ou "projet" ou "mission")
            if len(text) > MAX_XP_DESCRIPTION_LENGTH and any(keyword in text.lower() for keyword in ['contexte', 'projet', 'projets', 'mission', 'missions', 'développement', 'developpement', 'développements', 'developpements', 'objectif', 'objectifs', 'réalisation', 'realisation', 'réalisations', 'realisations', 'conception', 'montage', 'montages']):
                para['xp_metadata']['detected_xp_description'] = True
                return  # Ne pas faire d'autres détections si c'est une description

            # Détection xp_poste (KEYWORDS STRICTS SEULEMENT - pas de pattern générique)
            poste_keywords = ['Développeur', 'Développeuse', 'Developpeur', 'Developpeuse', 'Ingénieur', 'Ingénieure', 'Ingenieur', 'Ingenieure', 'Manager', 'Responsable', 'Chef', 'Cheffe', 'Lead', 'Tech Lead', 'Data Analyst', 'Data Engineer', 'Scientist', 'Technicien', 'Technicienne', 'Consultant', 'Consultante', 'Architecte', 'Directeur', 'Directrice', 'Senior', 'Product Owner', 'Scrum', 'DevOps', 'Administrateur', 'Administratrice', 'Alternance', 'Thèse', 'Doctorat', 'Stagiaire', 'Apprenti']
            # Vérifier si un mot-clé apparaît au début OU après un préfixe comme "Data"
            is_poste_keyword = False
            text_lower = text.lower()
            for kw in poste_keywords:
                kw_lower = kw.lower()
                # Check if text starts with keyword OR contains it as a word (not substring)
                if text.startswith(kw) or f' {kw_lower}' in f' {text_lower}' or text_lower.startswith(f'data {kw_lower}'):
                    is_poste_keyword = True
                    break
            # Si c'est un poste connu, le marquer
            if is_poste_keyword and len(text) < 50:
                para['xp_metadata']['detected_xp_poste'] = True
                return  # Ne pas faire d'autres détections si c'est un poste

            # Détection xp_company (nom propre court qui n'a pas été marqué comme poste)
            # Critères: court, commence par majuscule, pas de verbes d'action courants
            is_very_short = len(text) < 50
            has_no_common_verbs = not any(word in text.lower() for word in ['recueil', 'etude', 'étude', 'communication', 'rédaction', 'redaction', 'construction', 'évolutions', 'évolution', 'evolutions', 'evolution', 'système', 'systeme', 'systèmes', 'systemes', 'gestion', 'traitement', 'traitements', 'stockage', 'sauvegarde', 'parsing', 'dashboard', 'thèse', 'these'])
            starts_with_capital = text[0].isupper()
            if is_very_short and starts_with_capital and has_no_common_verbs:
                para['xp_metadata']['detected_xp_company'] = True
                return  # Ne pas faire d'autres détections si c'est une compagnie

def split_xp_entry(para: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Scinde une entrée d'expérience pro détectée par une DATE présente dans le texte.

    Utilise la détection de DATE (regex) pour identifier une XP entry valide.
    Avec lazy matching, le POSTE peut être absent.

    Format attendu: DATE : COMPANY - POSTE (POSTE optionnel)
    ou COMPANY - POSTE - DATE (dans ce cas, la DATE est détectée à l'intérieur du texte)
    ou COMPANY - DATE
        POSTE en paragraphe suivant (pas obligatoire)
    ou bien la table habituelle | COMPANY | DATE |
                                | POSTE   |      |
    Exemples:
    - "01/2021- 02/2024 : CALTOPO(USA)- Développeur Full Stack" → DATE | COMPANY | POSTE
    - "09/2017 – 09/2020 : MICROSOFT(USA)" → DATE | COMPANY | "" (poste vide)
    - "THALES - Ingénieur - 02-2022 à 05-2023" → COMPANY | POSTE | DATE (DATE détectée à l'intérieur du texte)
    - "CAPGEMINI - 01/2020 à 12/2021" → COMPANY | "" (poste vide ici mais dans le paragraphe suivant) | DATE (DATE détectée à l'intérieur du texte)

    Crée 2 ou 3 paragraphes:
    1. DATE (avec style 'xp_date')
    2. COMPANY (avec style 'xp_title')
    3. POSTE (avec style 'xp_poste') - optionnel

    Détection:
    - Cherche une DATE avec regex
    - Puis cherche `: ` ou `- ` qui sépare DATE de COMPANY
    - Puis cherche `- ` ou fin du texte (lazy) pour séparer COMPANY du POSTE

    Args:
        para: Paragraphe JSON source

    Returns:
        List[Dict]: Liste de 1 (pas XP entry) ou 2-3 paragraphes (XP entry splittée)
    """
    text = get_raw_text_from_paragraph(para)

    # Chercher `: ` qui sépare DATE de COMPANY
    colon_match = re.search(r':\s+', text)
    if not colon_match:
        return [para]  # Pas de `: ` trouvé

    # Extraire la portion DATE avant le `:` pour eviter les tronquages sur les plages
    date_candidate = text[:colon_match.start()].strip()

    prefix = ""
    date_body = date_candidate
    prefix_match = re.match(r'^\s*((?:depuis|du|de|à\s+partir\s+de)\s+)(.+)$', date_candidate, flags=re.IGNORECASE)
    if prefix_match:
        prefix = prefix_match.group(1)
        date_body = prefix_match.group(2).strip()

    range_sep = re.search(r'\s+[–-]\s+', date_body)
    if range_sep:
        left, right = re.split(r'\s+[–-]\s+', date_body, maxsplit=1)
        if not is_single_xp_date(left) or not is_single_xp_date(right):
            return [para]
        left_norm = re.sub(r'\s*([/–-])\s*', r'\1', left)
        right_norm = re.sub(r'\s*([/–-])\s*', r'\1', right)
        date_text = f"{prefix}{left_norm}-{right_norm}"
    else:
        date_tokens = list(re.finditer(r'\d{1,2}(?:\s*[/–-]\s*\d{1,2})?\s*[/–-]\s*\d{2,4}', date_body))
        if len(date_tokens) >= 2:
            left = date_tokens[0].group(0)
            right = date_tokens[1].group(0)
            between = date_body[date_tokens[0].end():date_tokens[1].start()]
            if not re.search(r'[–-]', between):
                return [para]
            if not is_single_xp_date(left) or not is_single_xp_date(right):
                return [para]
            left_norm = re.sub(r'\s*([/–-])\s*', r'\1', left)
            right_norm = re.sub(r'\s*([/–-])\s*', r'\1', right)
            date_text = f"{prefix}{left_norm}-{right_norm}"
        else:
            if not is_single_xp_date(date_body):
                return [para]
            date_text = prefix + re.sub(r'\s*([/–-])\s*', r'\1', date_body)
    remaining_after_date = text[colon_match.end():].strip()

    # Extraire COMPANY (apres `:` et avant le prochain `- ` ou fin du texte)
    after_colon = remaining_after_date
    dash_pattern = r'^(.+?)\s*[-–]\s+(.+)$'  # Lazy match pour COMPANY, greedy pour le reste
    dash_match = re.match(dash_pattern, after_colon)
    if dash_match:
        company_text = dash_match.group(1).strip()
        poste_text = dash_match.group(2).strip()
    else:
        company_text = after_colon.strip()
        poste_text = ""

    # Vérifier que DATE et COMPANY ont du contenu
    if not date_text or not company_text:
        return [para]

    # Créer les paragraphes
    first_run_props = para.get('runs', [{}])[0].get('properties', {}) if para.get('runs') else {}
    result = []

    # 1. Paragraphe DATE
    date_para = clone_paragraph_clean(para)
    date_para['runs'] = [{"text": date_text, "properties": first_run_props}]
    date_para['xp_split_part'] = 'xp_date'
    date_para['xp_entry_start'] = True  # Marquer que c'est le début d'une entry
    # Nettoyer les metadata du clone (ne pas propager celles du source)
    if 'xp_metadata' in date_para:
        del date_para['xp_metadata']
    result.append(date_para)

    # 2. Paragraphe COMPANY
    company_para = clone_paragraph_clean(para)
    company_para['runs'] = [{"text": company_text, "properties": first_run_props}]
    company_para['xp_split_part'] = 'xp_company'
    # Nettoyer les metadata du clone (ne pas propager celles du source)
    if 'xp_metadata' in company_para:
        del company_para['xp_metadata']
    result.append(company_para)

    # 3. Paragraphe POSTE
    poste_para = clone_paragraph_clean(para)
    poste_para['runs'] = [{"text": poste_text, "properties": first_run_props}]
    poste_para['xp_split_part'] = 'xp_poste'
    # Nettoyer les metadata du clone (ne pas propager celles du source)
    if 'xp_metadata' in poste_para:
        del poste_para['xp_metadata']
    result.append(poste_para)

    return result

def apply_xp_entry_splits(data: Dict[str, Any]) -> None:
    """
    Splitte toutes les entrées d'expérience pro détectées par DATE.

    Détection: Cherche les paragraphes qui:
    1. Sont marqués avec le tag 'professional_experience' (via apply_section_tags)
    2. N'ont pas de ilvl (ne sont pas des bullets)
    3. contiennent une DATE (regex)

    Format attendu: DATE : COMPANY - POSTE (POSTE optionnel)
    ou COMPANY - POSTE - DATE (dans ce cas, la DATE est détectée à l'intérieur du texte)
    ou COMPANY - DATE
        POSTE en paragraphe suivant (pas obligatoire)
    ou bien la table habituelle | COMPANY | DATE |
                                | POSTE   |      |

    Cette fonction doit être appelée APRÈS apply_section_tags() pour que les tags
    soient disponibles, et AVANT apply_section_header_styles() pour éviter que les
    XP entries soient marquées comme des headers.
    Elle flag tous les éléments qui permettent la distinction des blocs d'XP avec un champ `xp_split_part` (values: 'xp_date', 'xp_company', 'xp_poste', 'xp_description') pour les différencier des autres paragraphes.
    Elle flag les listes bullets classiques et les list bullets ou paragraphes appartenant à environnement technique, pour signifier une fin de xp entry (ex: compétences techniques listées à la fin d'une expérience pro).
    """
    content = data.get('document', {}).get('content', [])
    new_content = []

    for element in content:
        # Chercher les XP entries marquées avec le tag 'professional_experience'
        if (element.get('type') == 'Paragraph'):
            tags = element.get('tags', [])
            if isinstance(tags, str):
                tags = [tags]

            if 'professional_experience' in tags:
                props = element.get('properties', {})
                # Exclure les bullets (ilvl est défini)
                if props.get('ilvl') is None:
                    # Laisser split_xp_entry decider si c'est une vraie XP entry
                    split_result = split_xp_entry(element)
                    if len(split_result) > 1:
                        new_content.extend(split_result)
                        continue

        new_content.append(element)

    data['document']['content'] = new_content

def is_professional_section_header(element: Dict[str, Any]) -> bool:
    """Retourne True si l'élément est le header 'Expériences Professionnelles'."""
    if element.get('type') != 'Paragraph':
        return False
    props = element.get('properties', {})
    if props.get('style') != 'DC_T1_Sections':
        return False
    text = get_text_from_element(element)
    return any(keyword in text for keyword in KEYWORDS_PROFESSIONAL_EXPERIENCE)

def is_professional_tagged(element: Dict[str, Any]) -> bool:
    """Retourne True si l'élément appartient à la section expérience pro."""
    tags = element.get('tags', [])
    if isinstance(tags, str):
        tags = [tags]
    return 'professional_experience' in tags or element.get('properties', {}).get('section') == 'professional_experience'

def is_empty_paragraph(element: Dict[str, Any]) -> bool:
    """Retourne True si le paragraphe est vide."""
    if element.get('type') != 'Paragraph':
        return False
    return not get_text_from_element(element).strip()

def is_technical_skills_header(element: Dict[str, Any]) -> bool:
    """Retourne True si l'élément correspond au sous-bloc 'Environnement technique'.

    On cible explicitement les lignes qui démarrent par "environnement(s)" afin d'éviter
    les faux positifs sur des titres de poste (ex: "responsable technique").
    On exclut volontairement tout paragraphe contenant "contexte".
    """
    if element.get('type') != 'Paragraph':
        return False
    text = get_text_from_element(element)
    normalized_text = text.strip().lower()
    # Inclure les fautes courantes (environement/environements) pour rester tolérant aux typos source.
    if not normalized_text.startswith(('environnement', 'environnements', 'environement', 'environements')):
        return False
    return any(keyword in normalized_text for keyword in KEYWORDS_TECHNICAL_SKILLS) and 'contexte' not in normalized_text

def is_xp_description_paragraph(element: Dict[str, Any]) -> bool:
    """Retourne True si le paragraphe ressemble a un bloc de contexte long."""
    if element.get('type') != 'Paragraph':
        return False
    props = element.get('properties', {})
    if props.get('ilvl') is not None:
        return False
    text_raw = get_raw_text_from_paragraph(element).strip()
    if not text_raw:
        return False
    text_lower = text_raw.lower()
    return 'contexte' in text_lower or len(text_raw) > MAX_XP_DESCRIPTION_LENGTH

def has_bullets_after(content: List[Dict[str, Any]], start_idx: int, max_lookhead: int = 5) -> bool:
    """
    Vérifie s'il y a des paragraphes avec ilvl (bullets) dans les prochains éléments.

    Utile pour détecter les sous-sections "Environnement technique" qui sont suivies de listes.

    Args:
        content: Liste du contenu
        start_idx: Index de départ (non inclus)
        max_lookhead: Nombre d'éléments à regarder en avant

    Returns:
        True si au moins un paragraphe avec ilvl est trouvé
    """
    for j in range(start_idx + 1, min(start_idx + 1 + max_lookhead, len(content))):
        elem = content[j]
        if elem.get('type') == 'Paragraph':
            if elem.get('properties', {}).get('ilvl') is not None:
                return True
    return False

def table_contains_xp_entry_start(table: Dict[str, Any]) -> bool:
    """Vérifie si une table contient au moins un paragraphe avec xp_entry_start."""
    for row in table.get('rows', []):
        for cell in row.get('cells', []):
            for para in cell.get('paragraphs', []):
                if para.get('xp_entry_start'):
                    return True
    return False

def apply_xp_bullet_flags_and_levels(data: Dict[str, Any]) -> None:
    """
    Flag tous les paragraphes d'une XP entry en xp_bullet et ajuste les ilvl si nécessaire.

    Règle:
    - Dans chaque bloc XP (début: xp_date, fin: prochain xp_date ou sortie de section),
      tous les paragraphes non-vide entre les headers XP sont marqués xp_bullet.
    - Si le premier xp_bullet n'a pas de ilvl, on décale tous les ilvl d'un niveau:
      sans ilvl -> ilvl 0, ilvl0 -> ilvl1, etc.
    - Si le premier xp_bullet a déjà un ilvl, ne rien faire.
    - NOUVEAU: Dans un bloc "Environnement technique", tous les bullets doivent avoir ilvl = 2
    - Correction : si un paragraphe xp_ bullets n'a finalement pas de ilvl, lui appliquer un ilvl = 0
    """
    content = data.get('document', {}).get('content', [])
    current_block: List[int] = []
    technical_block: List[int] = []
    in_entry = False
    in_technical_block = False

    def finalize_block() -> None:
        if not current_block:
            return
        first_elem = content[current_block[0]]
        first_ilvl = first_elem.get('properties', {}).get('ilvl')
        if first_ilvl is None:
            for idx in current_block:
                elem = content[idx]
                props = elem.setdefault('properties', {})
                ilvl = props.get('ilvl')
                if ilvl is None:
                    props['ilvl'] = "0"
                else:
                    try:
                        props['ilvl'] = str(int(ilvl) + 1)
                    except (ValueError, TypeError):
                        pass

    def finalize_technical_block() -> None:
        """Force ilvl = 2 pour tous les bullets du bloc technique."""
        if not technical_block:
            return
        for idx in technical_block:
            elem = content[idx]
            props = elem.setdefault('properties', {})
            # Si le bloc technique a des bullets (ilvl), les forcer à 2
            if props.get('ilvl') is not None:
                props['ilvl'] = "2"

    def finalize_paras_without_ilvl_in_block() -> None:
        """Applique ilvl=0 à tous les bullets du bloc courant qui n'ont pas d'ilvl."""
        if not current_block:
            return
        for idx in current_block:
            elem = content[idx]
            props = elem.setdefault('properties', {})
            if props.get('ilvl') is None:
                props['ilvl'] = "0"

    for idx, element in enumerate(content):
        if not is_professional_tagged(element):
            if in_entry:
                finalize_block()
                finalize_technical_block()
                finalize_paras_without_ilvl_in_block()
            in_entry = False
            in_technical_block = False
            current_block = []
            technical_block = []
            continue

        # Vérifier si une table contient xp_entry_start (dates XP dans les cellules)
        if element.get('type') == 'Table' and table_contains_xp_entry_start(element):
            if in_entry:
                finalize_block()
                finalize_technical_block()
                finalize_paras_without_ilvl_in_block()
            in_entry = True
            in_technical_block = False
            current_block = []
            technical_block = []
            continue

        # Démarrer une entry sur xp_split_part == 'xp_date' OU xp_entry_start (paragraphes racine)
        if element.get('type') == 'Paragraph' and (element.get('xp_split_part') == 'xp_date' or element.get('xp_entry_start')):
            if in_entry:
                finalize_block()
                finalize_technical_block()
                finalize_paras_without_ilvl_in_block()
            in_entry = True
            in_technical_block = False
            current_block = []
            technical_block = []
            continue

        # Traiter les blocs techniques MÊME si on n'est pas en_entry
        if element.get('type') == 'Paragraph' and is_technical_skills_header(element):
            element['xp_technical'] = True
            finalize_technical_block()  # Finaliser le bloc précédent avant de démarrer un nouveau
            in_technical_block = True
            technical_block = []
            continue

        if in_technical_block:
            if element.get('type') == 'Paragraph':
                if element.get('properties', {}).get('ilvl') is not None:
                    element['xp_technical'] = True
                    technical_block.append(idx)
                    continue
                if is_empty_paragraph(element):
                    continue
                finalize_technical_block()
                in_technical_block = False
            elif element.get('type') == 'Table':
                finalize_technical_block()
                in_technical_block = False
                continue

        if not in_entry:
            continue

        if element.get('type') == 'Paragraph' and is_xp_description_paragraph(element):
            element['xp_split_part'] = 'xp_description'
            continue

        if element.get('type') != 'Paragraph':
            continue
        if element.get('xp_split_part'):
            continue
        if element.get('xp_entry_start'):
            continue  # Ne pas marquer le titre du bloc comme bullet
        if element.get('xp_technical'):
            continue  # Ne pas marquer les bullets techniques comme xp_bullet
        if is_professional_section_header(element):
            continue
        if is_empty_paragraph(element):
            continue  # Paragraphes vides restent vides

        element['xp_bullet'] = True
        current_block.append(idx)

    if in_entry:
        finalize_block()
        finalize_technical_block()
        finalize_paras_without_ilvl_in_block()
    else:
        finalize_technical_block()  # Finaliser le dernier bloc technique même s'il n'y a pas d'entry

def detect_xp_patterns(data: Dict[str, Any]) -> None:
    """
    Applique la détection automatique de patterns XP sur tous les paragraphes, y compris ceux dans les tables.

    Pour chaque paragraphe dans la section professionnelle:
    - Si déjà tagué avec xp_split_part, rien à faire (c'est explicite)
    - Sinon, applique _detect_and_tag_xp_content() pour les détections automatiques

    Les détections automatiques sont sauvegardées dans xp_metadata avec les clés:
    - detected_xp_date
    - detected_xp_company
    - detected_xp_poste
    - detected_xp_description

    Cela permet de vérifier quels paragraphes ont été détectés comme quoi,
    même s'ils n'ont pas été splittés.
    """
    content = data.get('document', {}).get('content', [])

    # Parcourir les paragraphes au niveau racine
    for element in content:
        if element.get('type') == 'Paragraph':
            tags = element.get('tags', [])
            if isinstance(tags, str):
                tags = [tags]

            # Appliquer la détection aux paragraphes professionnels non-splittés
            if 'professional_experience' in tags:
                _detect_and_tag_xp_content(element)

        # Parcourir aussi les paragraphes dans les tables
        elif element.get('type') == 'Table':
            tags = element.get('tags', [])
            if isinstance(tags, str):
                tags = [tags]

            # Si la table est une table professionnelle, détecter les patterns dans ses paragraphes
            if 'professional_experience' in tags:
                for row in element.get('rows', []):
                    for cell in row.get('cells', []):
                        for para in cell.get('paragraphs', []):
                            _detect_and_tag_xp_content(para)

def create_xp_tables(data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Crée les structures des tables professionnelles (Expériences Professionnelles).

    Responsabilité: Créer les tables vides et les insérer dans le contenu.

    Crée une table 2x2 pour:
    1. Après le header "Expériences Professionnelles" (pour le job entry)
    2. À la fin d'une sous-section "Environnement technique" (sortie de bullets)
    3. À la sortie d'une liste (transition ilvl → no ilvl) + paragraphe descriptif

    ⚠️ N'accélère PAS si KEYWORDS_TECHNICAL_SKILLS est suivi de bullets (c'est une sous-section)

    Args:
        data: Structure du document JSON contenant page_dimensions

    Returns:
        Dict contenant:
        - 'indices_to_delete': Indices à supprimer (vide pour XP)
    """
    # Récupérer les dimensions de page depuis data
    page_dims = data.get('page_dimensions')
    if page_dims is None:
        raise ValueError("page_dimensions non trouvées dans data")

    content = data.get('document', {}).get('content', [])

    # Créer des tables selon les conditions
    new_content = []
    expect_entry_start = False
    waiting_for_env_end = False
    last_was_bullet = False

    def mark_env_block(idx: int, content_list: List[Dict[str, Any]]) -> None:
        nonlocal waiting_for_env_end, expect_entry_start
        # Si une liste suit l'en-tête technique, on attend la fin du bloc pour démarrer l'XP suivante.
        if has_bullets_after(content_list, idx):
            waiting_for_env_end = True
        else:
            expect_entry_start = True

    i = 0
    while i < len(content):
        element = content[i]

        if not is_professional_tagged(element):
            expect_entry_start = False
            waiting_for_env_end = False
            last_was_bullet = False
            new_content.append(element)
            i += 1
            continue

        if element.get('type') == 'Paragraph':
            text = get_text_from_element(element)
            has_ilvl = element.get('properties', {}).get('ilvl') is not None
            # Détecter le header "Expériences Professionnelles"
            if is_professional_section_header(element):
                expect_entry_start = True
                waiting_for_env_end = False
                last_was_bullet = False
                new_content.append(element)
                i += 1
                continue

            is_empty = not text.strip()
            is_technical = is_technical_skills_header(element)

            if has_ilvl:
                last_was_bullet = True
            else:
                if last_was_bullet:
                    last_was_bullet = False
                    if is_technical:
                        mark_env_block(i, content)
                    else:
                        expect_entry_start = True

                if waiting_for_env_end and not is_technical and not is_empty:
                    expect_entry_start = True
                    waiting_for_env_end = False

                if is_technical:
                    mark_env_block(i, content)
        elif element.get('type') == 'Table':
            if waiting_for_env_end:
                expect_entry_start = True
                waiting_for_env_end = False

        # Condition pour créer une table AVANT l'élément courant (début d'xp_entry)
        should_create_table = element.get('xp_split_part') == 'xp_date'
        if element.get('type') == 'Table' and not element.get('auto_generated') and is_professional_tagged(element):
            should_create_table = True
        if should_create_table:
            prev_elem_is_table = len(new_content) > 0 and new_content[-1].get('type') == 'Table'
            if prev_elem_is_table:
                should_create_table = False

        if should_create_table:
            new_table = create_empty_table_2x2(
                len(new_content),
                section='professional_experience',
                auto_generated=True,
                page_dims=page_dims
            )
            new_content.append(new_table)

        new_content.append(element)
        i += 1

    data['document']['content'] = new_content

    return {'indices_to_delete': []}

def insert_text_xp_tables(data: Dict[str, Any], creation_result: Dict[str, Any], page_dims: dict = None) -> None:
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
    if page_dims is None:
        page_dims = data.get('page_dimensions')
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

                        # SKIP si c'est un titre d'environnement technique - le laisser en place
                        if is_technical_skills_header(next_elem):
                            break

                        # ARRÊTER si le paragraphe décrit un contexte
                        if 'contexte' in text:
                            break

                        # ARRÊTER si le paragraphe est long (> 75 caractères)
                        if len(text) > MAX_XP_DESCRIPTION_LENGTH:
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
                    temp_table = create_empty_table_2x2(
                        0,
                        section='professional_experience',
                        auto_generated=True,
                        num_rows=2,
                        page_dims=page_dims
                    )
                    element['rows'] = temp_table['rows']
                    element['row_count'] = len(element['rows'])

                    def pop_split_part(paras: List[Dict[str, Any]], part: str) -> Optional[Dict[str, Any]]:
                        """Retourne et retire le premier paragraphe matchant la partie demandée."""
                        for idx, para in enumerate(paras):
                            if para.get('xp_split_part') == part:
                                return paras.pop(idx)
                        return None

                    remaining = list(all_paragraphs)

                    company_para = pop_split_part(remaining, 'xp_company')
                    date_para = pop_split_part(remaining, 'xp_date')
                    poste_para = pop_split_part(remaining, 'xp_poste')

                    if company_para:
                        element['rows'][0]['cells'][0]['paragraphs'] = [clone_paragraph_clean(company_para)]
                    else:
                        # Trouver max size
                        max_size_para = None
                        max_size = 0

                        for para in remaining:
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

                    if date_para:
                        element['rows'][0]['cells'][1]['paragraphs'] = [clone_paragraph_clean(date_para)]
                    else:
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

                    if poste_para:
                        remaining.insert(0, poste_para)

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

    data['document']['content'] = content

# MARK: FINAL PROCESSING & RENDERING
# ===== 8. FINAL PROCESSING & RENDERING =====

def preserve_xp_metadata(data: Dict[str, Any]) -> None:
    """
    Préserve et explicite tous les tags xp_* dans le JSON transformé.

    Crée un champ `xp_metadata` contenant les métadonnées XP pour chaque paragraphe
    dans la section professionnelle, pour permettre les vérifications.

    Métadonnées préservées:
    - xp_split_part: ('xp_date', 'xp_company', 'xp_poste', 'xp_description')
    - xp_bullet: True si paragraphe bullet
    - xp_technical: True si dans bloc technique
    - xp_technical_bullet: True si bullet dans bloc technique
    - xp_entry_start: True si début d'une entry
    - ilvl: Niveau de liste

    Args:
        data: Structure du document JSON (modifiée in-place)
    """
    content = data.get('document', {}).get('content', [])

    for element in content:
        # Parcourir aussi les paragraphes dans les tables
        if element.get('type') == 'Table':
            for row in element.get('rows', []):
                for cell in row.get('cells', []):
                    for para in cell.get('paragraphs', []):
                        _build_xp_metadata(para)
        elif element.get('type') == 'Paragraph':
            _build_xp_metadata(element)

def _build_xp_metadata(para: Dict[str, Any]) -> None:
    """
    Construit le champ xp_metadata pour un paragraphe.

    Les xp_metadata ne sont créées que pour les paragraphes ayant des propriétés XP réelles.
    ilvl n'est PAS inclus (c'est une propriété standard du paragraphe, pas une métadonnée XP).

    Les flags xp_* directs sont supprimés après avoir été copiés dans xp_metadata.
    """
    metadata = {}

    # Récupérer les métadonnées existantes (si présentes) - mais pas ilvl
    if 'xp_metadata' in para:
        for key, value in para['xp_metadata'].items():
            if key != 'ilvl':  # Exclure ilvl des metadata copiées
                metadata[key] = value

    # Récupérer TOUS les champs xp_* du paragraphe (en tant que propriétés directes)
    # Ces champs marquent le paragraphe comme ayant des propriétés XP
    has_xp_property = False

    if 'xp_split_part' in para:
        metadata['xp_split_part'] = para['xp_split_part']
        has_xp_property = True
    if 'xp_bullet' in para:
        metadata['xp_bullet'] = para['xp_bullet']
        has_xp_property = True
    if 'xp_technical' in para:
        metadata['xp_technical'] = para['xp_technical']
        has_xp_property = True
    if 'xp_technical_bullet' in para:
        metadata['xp_technical_bullet'] = para['xp_technical_bullet']
        has_xp_property = True
    if 'xp_entry_start' in para:
        metadata['xp_entry_start'] = para['xp_entry_start']
        has_xp_property = True

    # IMPORTANT: Ajouter le champ xp_metadata au paragraphe SEULEMENT s'il a des propriétés XP
    # Ne pas créer de xp_metadata vide ou seulement basée sur ilvl
    if has_xp_property or metadata:
        para['xp_metadata'] = metadata
    elif 'xp_metadata' in para:
        # Si aucune propriété XP et pas de metadata existante, supprimer le champ vide
        del para['xp_metadata']

    # Nettoyer les flags xp_* directs du paragraphe après les avoir copiés dans xp_metadata
    # Garder SEULEMENT dans xp_metadata, pas en propriété directe
    for key in ['xp_split_part', 'xp_bullet', 'xp_technical', 'xp_technical_bullet', 'xp_entry_start']:
        if key in para:
            del para[key]

def add_empty_paragraphs_around_tables(data: Dict[str, Any]) -> None:
    """
    Ajoute un paragraphe vide AVANT et APRÈS chaque table du document.

    Utile pour:
    - Espace visuel avant et après les tables
    - Permettre à Word de naviguer correctement
    - Faciliter le rendu et l'édition

    Cette fonction est appelée APRÈS que toutes les tables aient été créées
    et remplies (create_edu_table, insert_text_edu_table, create_xp_tables, insert_text_xp_tables),
    mais AVANT le nettoyage des doubles paragraphes.

    Args:
        data: Structure du document JSON
    """
    content = data.get('document', {}).get('content', [])

    if not content:
        return

    # Construire une nouvelle liste avec les paragraphes vides autour des tables
    new_content = []

    for elem in content:
        # Si c'est une table, ajouter un paragraphe vide AVANT
        if elem.get('type') == 'Table':
            # Vérifier si le dernier élément ajouté n'est pas déjà un paragraphe vide
            if new_content and new_content[-1].get('type') == 'Paragraph':
                last_para = new_content[-1]
                # Si le dernier paragraphe n'est pas vide, ajouter un paragraphe vide
                if last_para.get('runs') or last_para.get('text'):
                    new_content.append({
                        'type': 'Paragraph',
                        'properties': {},
                        'runs': []
                    })
            elif not new_content or new_content[-1].get('type') == 'Table':
                # Ajouter un paragraphe vide avant la table
                new_content.append({
                    'type': 'Paragraph',
                    'properties': {},
                    'runs': []
                })

        # Ajouter l'élément lui-même
        new_content.append(elem)

        # Si c'est une table, ajouter un paragraphe vide APRÈS
        if elem.get('type') == 'Table':
            new_content.append({
                'type': 'Paragraph',
                'properties': {},
                'runs': []
            })

    # Remplacer le contenu du document
    data['document']['content'] = new_content

def _clean_paragraphs_list(paragraphs: List[Dict[str, Any]]) -> bool:
    """
    Nettoie une liste de paragraphes en supprimant les doublons vides et les doubles espaces.

    Retourne True si des changements ont été faits, False sinon.
    Modifie la liste in-place.

    Args:
        paragraphs: Liste de paragraphes à nettoyer

    Returns:
        bool: True si des changements ont été faits
    """
    if not paragraphs:
        return False

    new_list = []
    last_was_empty = False
    changes_made = False
    initial_length = len(paragraphs)

    for para in paragraphs:
        text = get_text_from_element(para)
        is_empty = not text.strip()

        if is_empty:
            # Garder seulement 1 paragraphe vide (éviter 2 consécutifs)
            if not last_was_empty:
                new_list.append(para)
            else:
                changes_made = True  # On a supprimé un paragraphe vide
            last_was_empty = True
        else:
            # Paragraphe non-vide : nettoyer les doubles espaces dans les runs
            if 'runs' in para:
                for run in para['runs']:
                    if 'text' in run:
                        original_text = run['text']
                        # Remplacer TOUS les espaces multiples par un simple espace (boucle)
                        while '  ' in run['text']:
                            run['text'] = run['text'].replace('  ', ' ')
                        if original_text != run['text']:
                            changes_made = True

            new_list.append(para)
            last_was_empty = False

    # Remplacer la liste in-place
    paragraphs.clear()
    paragraphs.extend(new_list)

    # Vérifier si la taille a changé
    if len(new_list) != initial_length:
        changes_made = True

    return changes_made


def remove_double_paras_and_spaces(data: Dict[str, Any]) -> None:
    """
    Supprime les paragraphes vides doublons et nettoie les doubles espaces.
    Modifie in-place.

    Boucle jusqu'à ce qu'aucun changement ne soit détecté.

    Logique:
    - Parcourir le contenu du document (paragraphes root + paragraphes dans les tables)
    - Garder une trace du dernier paragraphe ajouté
    - Supprimer les paragraphes vides doublons (garder max 1 paragraphe vide consécutif)
    - Remplacer les doubles espaces ("  ") par un simple espace (" ") dans les runs
    - Répéter jusqu'à stabilité (aucun changement)
    - Nettoyer aussi les paragraphes à l'intérieur des cellules des tables
    """
    import sys
    iteration = 0
    while True:
        iteration += 1
        content = data.get('document', {}).get('content', [])
        changes_made = False
        initial_length = len(content)

        # ÉTAPE 1: Nettoyer les paragraphes root
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
                    else:
                        changes_made = True  # On a supprimé un paragraphe vide
                    last_para_was_empty = True
                else:
                    # Paragraphe non-vide : nettoyer les doubles espaces dans les runs
                    if 'runs' in element:
                        for run in element['runs']:
                            if 'text' in run:
                                original_text = run['text']
                                # Remplacer TOUS les espaces multiples par un simple espace (boucle)
                                while '  ' in run['text']:
                                    run['text'] = run['text'].replace('  ', ' ')
                                if original_text != run['text']:
                                    changes_made = True

                    new_content.append(element)
                    last_para_was_empty = False
            elif element.get('type') == 'Table':
                # ÉTAPE 2: Nettoyer les paragraphes dans les cellules des tables
                rows = element.get('rows', [])
                for row in rows:
                    cells = row.get('cells', [])
                    for cell in cells:
                        paragraphs = cell.get('paragraphs', [])
                        if paragraphs:
                            # Nettoyer cette liste de paragraphes
                            if _clean_paragraphs_list(paragraphs):
                                changes_made = True

                new_content.append(element)
                last_para_was_empty = False
            else:
                new_content.append(element)
                last_para_was_empty = False

        data['document']['content'] = new_content

        # Si la taille a changé ou aucun changement détecté, s'arrêter
        if len(new_content) == initial_length and not changes_made:
            break

def recalculate_indices(data: Dict[str, Any]) -> None:
    """
    Recalcule les indices de tous les éléments du document de manière continue.

    Après les transformations (ajout/suppression d'éléments), les indices peuvent avoir
    des trous. Cette fonction les recalcule de 0 à n de manière séquentielle.

    Args:
        data: Structure du document JSON (modifiée in-place)
    """
    content = data.get('document', {}).get('content', [])

    for i, element in enumerate(content):
        element['index'] = i

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

    # Appliquer les styles des tables main skills, éducation et expérience professionnelle
    for itable in data.get('document', {}).get('content', []):
        if itable.get('type') == 'Table':
            tags = itable.get('tags', [])
            if isinstance(tags, str):
                tags = [tags]

            # Détecter le type de table via tags ET properties
            is_main_skills = 'main_skills' in tags or itable.get('properties', {}).get('section') == 'main_skills'
            is_education = 'education' in tags or itable.get('properties', {}).get('section') == 'education'
            is_professional = 'professional_experience' in tags or itable.get('properties', {}).get('section') == 'professional_experience'

            if is_main_skills:
                rows = itable.get('rows', [])
                # Appliquer le style DC_Table_Skills_Title aux paragraphes dans cell[x][0] (colonne 0)
                for row in rows:
                    cells = row.get('cells', [])
                    if len(cells) > 0:
                        for para in cells[0].get('paragraphs', []):
                            if 'properties' not in para:
                                para['properties'] = {}
                            para['properties']['style'] = 'DC_Table_Skills_Title'
                # Appliquer le style DC_Table_Skills_Content aux paragraphes dans cell[x][1..n] (colonnes 1 à n)
                for row in rows:
                    cells = row.get('cells', [])
                    if len(cells) > 1:
                        for cell in cells[1:]:
                            for para in cell.get('paragraphs', []):
                                if 'properties' not in para:
                                    para['properties'] = {}
                                para['properties']['style'] = 'DC_Table_Skills_Content'

            if is_education:
                rows = itable.get('rows', [])
                # Appliquer le style DC_Table_Year aux paragraphes dans cell[x][0] (colonne 0)
                for row in rows:
                    cells = row.get('cells', [])
                    if len(cells) > 0:
                        for para in cells[0].get('paragraphs', []):
                            if 'properties' not in para:
                                para['properties'] = {}
                            para['properties']['style'] = 'DC_Table_Year'
                # Appliquer le style DC_Table_Content aux paragraphes dans cell[x][1..n] (colonnes 1 à n)
                for row in rows:
                    cells = row.get('cells', [])
                    if len(cells) > 1:
                        for cell in cells[1:]:
                            for para in cell.get('paragraphs', []):
                                if 'properties' not in para:
                                    para['properties'] = {}
                                para['properties']['style'] = 'DC_Table_Content'

            if is_professional:
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
        if 'professional_experience' in tags and is_technical_skills_header(itag):
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
                    # Supprimer les runs supplémentaires qui étaient fusionnés
                    if len(ilist['runs']) > 1:
                        ilist['runs'] = ilist['runs'][:1]
        elif ilvl == "1":
            props['style'] = 'DC_2nd_bullet'
        elif ilvl == "2":
            props['style'] = 'DC_3rd_bullet'
        elif ilvl == "3":
            props['style'] = 'DC_4th_bullet'
        else:
            props['style'] = 'DC_Normal'  # fallback

    # Appliquer les styles des titres après les listes pour corriger les faux positifs liés à ilvl
    for itag in data.get('document', {}).get('content', []):
        if 'tags' not in itag:
            continue
        tags = itag['tags']
        text = get_text_from_element(itag) if itag else ""
        props = itag.setdefault('properties', {})

        if 'header' in tags and any(keyword in text.lower() for keyword in KEYWORDS_HEADER_DOCUMENT):
            props['style'] = 'DC_H_DC'
            text = text.upper()
            if 'runs' in itag and itag['runs']:
                itag['runs'][0]['text'] = text
        elif 'main_skills' in tags and is_promotable_section_title(itag, KEYWORDS_MAIN_SKILLS):
            props['style'] = 'DC_T1_Sections'
            text = text.capitalize()
            if 'runs' in itag and itag['runs']:
                itag['runs'][0]['text'] = text
        elif 'education' in tags and is_promotable_section_title(itag, KEYWORDS_EDUCATION):
            props['style'] = 'DC_T1_Sections'
            text = text.capitalize()
            if 'runs' in itag and itag['runs']:
                itag['runs'][0]['text'] = text
        elif 'professional_experience' in tags and is_promotable_section_title(itag, KEYWORDS_PROFESSIONAL_EXPERIENCE):
            props['style'] = 'DC_T1_Sections'
            text = text.capitalize()
            if 'runs' in itag and itag['runs']:
                itag['runs'][0]['text'] = text

        if 'header' in tags:
            if len(text) > 0 and len(text) <= 5 and props.get('style') != 'DC_H_DC':
                props['style'] = 'DC_H_Trigramme'
                text = text.upper()
                if 'runs' in itag and itag['runs']:
                    itag['runs'][0]['text'] = text
            elif any(keyword in text.lower() for keyword in KEYWORDS_HEADER_EXPERIENCE) and len(text) > 5 and props.get('style') != 'DC_H_DC':
                props['style'] = 'DC_H_XP'
            elif len(text) > 5 and props.get('style') not in ('DC_H_DC', 'DC_H_XP'):
                props['style'] = 'DC_H_Poste'

    # Appliquer le style Normal pour le reste et les éléments sans style
    for element in data.get('document', {}).get('content', []):
        props = element.setdefault('properties', {})

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

# MARK: MAIN ENTRY POINT
# ===== 9. MAIN ENTRY POINT =====

def apply_tags_and_styles(raw_json_file: str, output_dir: str, page_dimensions: dict) -> str:
    """
    Charge un JSON brut, applique les tags de section et les styles,
    puis enregistre le résultat transformé.

    Args:
        raw_json_file (str): Chemin du fichier JSON RAW
        output_dir (str): Répertoire de sortie
        page_dimensions (dict): Dimensions de page (extraites une seule fois du template)

    Returns:
        str: Chemin du fichier créé
    """
    input_path = Path(raw_json_file)

    # Créer le répertoire s'il n'existe pas
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Générer le nom de sortie
    output_file = output_dir / (input_path.stem.replace('_raw', '') + "_transformed.json")

    # Charger le JSON RAW
    with open(input_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # Stocker les dimensions dans le document pour utilisation ultérieure
    if 'page_dimensions' not in data:
        data['page_dimensions'] = page_dimensions

    # ===== DETECTER LES 4 SECTIONS =====
    # Appliquer les tags de section
    apply_section_tags(data)

    # Splitter les entrées d'expérience pro AVANT de marquer les headers
    apply_xp_entry_splits(data)

    # Appliquer le style DC_T1_Sections aux headers de section
    apply_section_header_styles(data)

    # Détecter les patterns XP sur tous les paragraphes (même ceux non-splittés)
    # DOIT ÊTRE AVANT apply_xp_bullet_flags_and_levels car elle se sert de xp_entry_start
    detect_xp_patterns(data)

    # Flagger les bullets des XP entries et ajuster les ilvl si nécessaire
    apply_xp_bullet_flags_and_levels(data)

    # ===== TABLE MAIN SKILLS si existante =====
    # Créer la table Main Skills si on détecte une table source des compétences techniques
    main_skills_creation_result = create_main_skills_table(data)

    # Remplir le contenu de la table Main Skills et supprimer les sources
    insert_text_main_skills_table(data, main_skills_creation_result, page_dims=page_dimensions)

    # ===== TABLES ÉDUCATION =====
    # Créer le header "Langues" juste avant le premier keyword détecté
    create_language_header(data)

    # Créer les structures (Formation et Langues)
    edu_creation_result = create_edu_table(data)

    # Remplir le contenu et supprimer les sources
    insert_text_edu_table(data, edu_creation_result, page_dims=page_dimensions)

    # ===== TABLES EXPÉRIENCES PROFESSIONNELLES =====
    # Créer les structures
    xp_creation_result = create_xp_tables(data)

    # Remplir le contenu et supprimer les sources
    insert_text_xp_tables(data, xp_creation_result, page_dims=page_dimensions)

    # Appliquer la détection XP aux paragraphes dans les tables nouvellement remplies
    # (car detect_xp_patterns() a été appelée avant la création des tables)
    detect_xp_patterns(data)

    # ===== AJOUT DES PARAGRAPHES VIDES AUTOUR DES TABLES =====
    # Ajouter un paragraphe vide avant et après chaque table
    # (après que toutes les tables aient été créées/remplies)
    add_empty_paragraphs_around_tables(data)

    # ===== NETTOYAGE et RENDU FINAL POUR CHAQUE ELEMENT =====
    # Ajouter les ":" entre les niveaux de listes successifs
    add_colons_between_list_levels(data)

    # Nettoyer les paragraphes vides doublons et les doubles espaces
    remove_double_paras_and_spaces(data)

    # Recalculer les indices de manière continue après toutes les transformations
    recalculate_indices(data)

    # Préserver et expliciter tous les tags xp_* pour vérification des détections
    preserve_xp_metadata(data)

    # Appliquer les styles
    apply_styles_in_json(data)

    # DEUXIÈME PASSE DE NETTOYAGE: Après la fusion des runs dans apply_styles_in_json
    # (qui peut reintroduire des doubles espaces lors de text.capitalize())
    remove_double_paras_and_spaces(data)

    # Sauvegarder le JSON transformé
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return str(output_file)

def main():
    """
    Fonction principale: orchestre le pipeline de transformation
    - Génère le JSON RAW depuis le XML
    - Applique les tags et styles
    - Enregistre les deux versions
    """
    parser = ArgumentParser(description="Transforme un JSON RAW (tags + styles)")

    parser.add_argument(
        "-s", "--source_json_raw",
        required=True,
        help="Chemin du fichier JSON RAW"
    )

    parser.add_argument(
        "-t", "--template",
        default=TEMPLATE_PATH,
        help=f"Chemin du template DOCX (défaut: {TEMPLATE_PATH})"
    )

    parser.add_argument(
        "-o", "--output_dir",
        default="OUTPUT3_JSON-TRANSFORMED",
        help="Répertoire de sortie (défaut: OUTPUT3_JSON-TRANSFORMED)"
    )

    args = parser.parse_args()

    # Extraire les dimensions du template
    page_dims = extract_page_dimensions_from_template(args.template)

    # Appliquer les tags et styles
    json_transformed = apply_tags_and_styles(args.source_json_raw, args.output_dir, page_dims)

    if not json_transformed:
        print("❌ Erreur lors de la transformation")
        sys.exit(1)

if __name__ == "__main__":
    main()
