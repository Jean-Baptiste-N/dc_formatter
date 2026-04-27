# Documentation Complète - DC Formatter

## 📋 Table des Matières

1. [Architecture Générale](#architecture-générale)
2. [Pipeline de Transformation](#pipeline-de-transformation)
3. [Modules Python Détaillés](#modules-python-détaillés)
4. [Formats de Données](#formats-de-données)
5. [Hiérarchie et Tags](#hiérarchie-et-tags)
6. [Styles et Mise en Forme](#styles-et-mise-en-forme)
7. [Troubleshooting](#troubleshooting)

---

## Architecture Générale

### Vue d'Ensemble

DC Formatter traite les documents Word selon un pipeline en 4 étapes:

```
┌─────────────────────────────────────────────────────────┐
│           DOCUMENT WORD SOURCE (DOCX)                   │
│  - Structure hiérarchique (Heading1, Heading2, etc.)    │
│  - Paragraphes et tableaux                              │
│  - Styles et formatage                                  │
└────────────────────┬────────────────────────────────────┘
                     │
                     ↓
        ┌────────────────────────────┐
        │  PHASE 1: EXTRACTION       │
        │  - Extraction XML brut     │
        │  - Dimensions template     │
        │  - Création XML global     │
        └────────────────┬───────────┘
                         │
                         ↓
           ┌──────────────────────────┐
           │   OUTPUT1_XML-RAW/       │
           │  *_GLOBAL.xml            │
           └────────────┬─────────────┘
                        │
                        ↓
        ┌────────────────────────────┐
        │  PHASE 2: CONVERSION       │
        │  - Parsing XML en JSON     │
        │  - Extraction structure    │
        │  - Propriétés (style, ...) │
        └────────────────┬───────────┘
                         │
                         ↓
           ┌──────────────────────────┐
           │   OUTPUT2_JSON-RAW/      │
           │  *_GLOBAL_raw.json       │
           └────────────┬─────────────┘
                        │
                        ↓
        ┌────────────────────────────┐
        │  PHASE 3: TRANSFORMATION   │
        │  - Détection hiérarchie    │
        │  - Application tags        │
        │  - Création tableaux       │
        │  - Application styles      │
        └────────────────┬───────────┘
                         │
                         ↓
           ┌──────────────────────────┐
           │ OUTPUT3_JSON-TRANSFORMED/│
           │ *_GLOBAL_transformed.json│
           └────────────┬─────────────┘
                        │
                        ↓
        ┌────────────────────────────┐
        │  PHASE 4: RENDU            │
        │  - Injection dans template │
        │  - Application bordures    │
        │  - Génération DOCX         │
        └────────────────┬───────────┘
                         │
                         ↓
           ┌──────────────────────────┐
           │ OUTPUT4_DOCX-RESULT/     │
           │  *_GLOBAL_formatted.docx │
           └──────────────────────────┘
```

### Composants Clés

- **Template DOCX** (`assets/TEMPLATE.docx`): Référence pour dimensions et styles
- **Namespaces XML**: Standards OOXML (word, drawing, etc.)
- **JSON intermédiaire**: Format de travail pour transformations
- **Configuration**: Détecteurs de hiérarchie, keywords, dimensions

---

## Pipeline de Transformation

### Phase 1: Extraction (extract_xml_raw.py)

**Entrée**: Document DOCX (zip archive)
**Sortie**: XML global + dimensions

#### Fonctions Principales

| Fonction | Description |
|----------|-------------|
| `extract_xml_raw(docx_file)` | Extrait tous les fichiers XML du DOCX (word/document.xml, document.xml.rels, styles.xml, etc.) |
| `create_global_xml(xml_contents)` | Combine tous les XMLs dans un seul document racine |
| `extract_document_xml(docx_file)` | Extrait uniquement word/document.xml formaté |
| `export_all_xml(docx_file, output_dir)` | Orchestre extraction + formatage + sauvegarde |
| `indent_xml_string(xml_string)` | Indente et formate le XML pour lisibilité |

#### Exemple

```python
from tools3.extract_xml_raw import export_all_xml
xml_file = export_all_xml("document.docx", "OUTPUT1_XML-RAW/")
# Crée: OUTPUT1_XML-RAW/document_GLOBAL.xml
```

---

### Phase 2: Conversion XML → JSON RAW (parse_xml_raw_to_json_raw.py)

**Entrée**: XML global brut
**Sortie**: JSON RAW avec structure complète

#### Structure JSON RAW

```json
{
  "document": {
    "paragraphs": [
      {
        "index": 0,
        "type": "Paragraph",
        "properties": {
          "pStyle": "Heading1",
          "alignment": "center",
          "bold": true,
          "size": 40
        },
        "runs": [
          {
            "text": "TITRE DU DOCUMENT",
            "properties": {
              "bold": true,
              "color": "538cd3",
              "size": 40
            }
          }
        ]
      }
    ],
    "tables": [
      {
        "type": "Table",
        "row_count": 3,
        "col_count": 2,
        "rows": [...]
      }
    ]
  }
}
```

#### Fonctions Principales

| Fonction | Description |
|----------|-------------|
| `xml_to_json(xml_file, output_path)` | Convertit XML complet en JSON RAW |
| `extract_paragraph_properties()` | Extrait propriétés (style, alignment, etc.) |
| `extract_run_properties()` | Extrait propriétés de texte (bold, italic, color, etc.) |
| `parse_table()` | Parse tables avec cellules et contenu |

---

### Phase 3: Transformation (process_json_raw_to_json_transformed.py)

**Entrée**: JSON RAW
**Sortie**: JSON enrichi avec structure et tags

#### Transformations Appliquées

1. **Détection de hiérarchie**
   - Identification Heading1 (titres principaux)
   - Identification Heading2 (sous-titres)
   - Classification sections (Experience, Education, Skills, etc.)

2. **Application de tags**
   - `<section>`: Identification sections principales
   - `<subsection>`: Identification sous-sections
   - `<content>`: Contenu régulier
   - `<table>`: Tableaux structurés

3. **Création de structures**
   - Tables d'éducation (Date | Formation)
   - Tables d'expérience (Poste | Période)
   - Listes de compétences

#### Fonctions Clés

| Fonction | Description |
|----------|-------------|
| `apply_tags_and_styles()` | Orchestre toutes transformations |
| `detect_sections()` | Identifie sections (Experience, Education, Skills) |
| `apply_heading_tags()` | Ajoute tags Heading1/H2 |
| `create_education_table()` | Crée table Education (date + formation) |
| `create_professional_table()` | Crée table Expérience (poste + période) |

#### Exemple JSON Transformé

```json
{
  "sections": [
    {
      "type": "Header",
      "tag": "section",
      "content": "DOSSIER DE COMPÉTENCES"
    },
    {
      "type": "Skills",
      "tag": "subsection", 
      "entries": [
        {"tag": "skill", "content": "Python, JavaScript, SQL"}
      ]
    },
    {
      "type": "Education",
      "tag": "subsection",
      "table": {
        "rows": [
          {
            "cells": [
              {"content": "2018-2020"},
              {"content": "Master Informatique"}
            ]
          }
        ]
      }
    }
  ]
}
```

---

### Phase 4: Rendu (render_json_transformed_to_docx.py)

**Entrée**: JSON transformé + template DOCX
**Sortie**: DOCX final formaté

#### Processus

1. **Chargement template** → Document vide avec styles définis
2. **Injection structure** → Paragraphes et tables du JSON
3. **Application formatage** → Styles, couleurs, bordures
4. **Génération** → Sauvegarde DOCX final

#### Fonctions Principales

| Fonction | Description |
|----------|-------------|
| `json_to_docx()` | Orchestre rendu complet |
| `add_paragraph_from_json()` | Ajoute paragraphe avec propriétés |
| `add_table_from_json()` | Ajoute tableau avec contenu et bordures |
| `set_table_borders()` | Configure bordures tableau |
| `parse_alignment()` | Convertit alignment string en enum Word |

---

## Modules Python Détaillés

### extract_xml_raw.py

Responsable de l'extraction du contenu XML depuis le DOCX.

**Classes/Fonctions principales**:
- `extract_xml_raw(docx_file)` - Extrait tous les .xml
- `create_global_xml(xml_contents)` - Combine en XML global
- `export_all_xml(docx_file, output_dir)` - Orchestre tout

**Dépendances**: `zipfile`, `ElementTree`, `minidom`

### parse_template.py

Extrait les dimensions et paramètres du template DOCX.

**Classe principale**:
- `extract_page_dimensions_from_template(template_path)` → dict

**Dimensions extraites**:
```python
{
    'page_width': 11906,        # twips (A4)
    'page_height': 16838,       # twips
    'top_margin': 1440,         # twips
    'bottom_margin': 1440,
    'left_margin': 1440,
    'right_margin': 1440,
    'usable_width': 8226        # page_width - marges
}
```

### parse_xml_raw_to_json_raw.py

Convertit XML brut en structure JSON.

**Fonctions principales**:
- `xml_to_json(xml_file, output_path)` - Conversion complète
- `extract_paragraph_properties()` - Propriétés paragraphe
- `extract_run_properties()` - Propriétés texte (bold, color, etc.)
- `parse_table()` - Parsing des tableaux

### process_json_raw_to_json_transformed.py

Applique tags, détecte hiérarchie et crée structures.

**Keywords détection**:
```python
KEYWORDS_EDUCATION = ["formation", "diplôme", "certification", "langue"]
KEYWORDS_PROFESSIONAL_EXPERIENCE = ["expérience professionnelle"]
KEYWORDS_TECHNICAL_SKILLS = ["techniques", "informatiques"]
```

**Fonctions principales**:
- `apply_tags_and_styles()` - Orchestre transformations
- `detect_sections()` - Identification sections
- `create_education_table()` - Détecte et crée table Education
- `create_professional_table()` - Détecte et crée table Expérience

### render_json_transformed_to_docx.py

Génère DOCX final à partir du JSON transformé.

**Fonctions principales**:
- `json_to_docx(json_file, template_path, output_dir)`
- `add_paragraph_from_json()` - Ajoute paragraphes
- `add_table_from_json()` - Ajoute tableaux
- `set_table_borders()` - Configure bordures

### pipeline.py

Interface CLI qui orchestre le pipeline complet.

**Commandes principales**:
- `full` - Pipeline complète
- `extract` - Phase 1 (extraction)
- `transform-render` - Phase 2 (transformation + rendu)
- `extract-dims` - Extraction dimensions
- `extract-xml` - Extraction XML
- `xml-to-json` - Conversion XML → JSON
- `transform` - Transformation seule
- `render` - Rendu seul

---

## Formats de Données

### XML (Phase 1)

Structure OOXML Word avec tous les namespaces:
- **word/document.xml** - Contenu principal
- **word/styles.xml** - Styles définis
- **word/document.xml.rels** - Relations
- **customXml/** - Données customisées
- **_rels/.rels** - Métadonnées

Namespaces principaux:
```xml
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
```

### JSON RAW (Phase 2)

Représentation fidèle de la structure XML:
- Tous les paragraphes avec propriétés
- Tous les runs (texte formaté)
- Tables avec rows/cells
- Propriétés préservées (style, alignment, couleur, etc.)

### JSON TRANSFORMED (Phase 3)

Structure enrichie avec:
- Sections identifiées et taggées
- Hiérarchie (Heading1 → Heading2 → Content)
- Tables structurées
- Styles appliqués

### DOCX Final (Phase 4)

Document Word standard formaté:
- Basé sur template (preserves styles)
- Contenu injecté et formaté
- Tableaux avec bordures
- Mise en page respectée

---

## Hiérarchie et Tags

### Détection Hiérarchie

**Heading1** (Titres principaux):
- Texte court ET centré OU complètement MAJUSCULE
- Taille grande (40pt+)
- Souvent en couleur spéciale (bleu)

**Heading2** (Sous-titres):
- Texte court-moyen finissant par ":"
- Contient keywords section (Experience, Education, etc.)
- Généralement bold

**Content** (Contenu normal):
- Paragraphes réguliers
- Longueur variable
- Formatage standard

### Sections Identifiées

| Section | Keywords |
|---------|----------|
| Experience | "expérience professionnelle" |
| Education | "formation", "diplôme", "certification" |
| Skills | "compétences", "techniques", "informatiques" |
| Languages | "langue", "français", "anglais" |

### Tags Appliqués

```json
"tags": {
  "type": "Heading1|Heading2|Content|Table",
  "section": "Header|Skills|Education|Experience|Languages|Document",
  "properties": {...}
}
```

---

## Styles et Mise en Forme

### Propriétés Paragraphe

- **pStyle**: Style appliqué (Normal, Heading1, etc.)
- **alignment**: left|center|right|both
- **indent**: Indentation gauche/droite
- **spacing**: Espacement avant/après
- **section_break**: Type de saut de section

### Propriétés Run (Texte)

- **bold**: Texte en gras
- **italic**: Texte en italique
- **size**: Taille en demi-points (20 = 10pt)
- **color**: Couleur en hexadécimal (rrggbb)
- **font**: Nom de la police

### Propriétés Table

- **table_width**: Largeur totale en twips
- **borders**: Définition bordures top/bottom/left/right/insideH/insideV
- **style**: Style table appliqué
- **alignment**: Alignment table

---

## Troubleshooting

### Erreur: "XML parsing failed"
**Cause**: Fichier DOCX endommagé ou format incorrect
**Solution**: Vérifier intégrité DOCX, ouvrir/resauvegarder dans Word

### Erreur: "JSON transformation failed"
**Cause**: Structure JSON invalide après extraction
**Solution**: Vérifier format XML source, relancer extraction

### Résultat DOCX incomplet
**Cause**: Sections non détectées correctement
**Solution**: Vérifier keywords dans process_json_raw_to_json_transformed.py

### Tables mal formatées
**Cause**: Dimensions de colonne incorrectes
**Solution**: Ajuster dans `get_table_widths_for_section()`

---

## Pour Aller Plus Loin

- **Fichier balises**: Voir [DICTIONNAIRE_BALISES_XML.md](DICTIONNAIRE_BALISES_XML.md)
- **Patterns XML**: Voir [analyse/PATTERNS_XML_DECOUVERTES.md](analyse/PATTERNS_XML_DECOUVERTES.md)
- **Détection hiérarchie**: Voir [analyse/GUIDE_DETECTION_HIERARCHIE.md](analyse/GUIDE_DETECTION_HIERARCHIE.md)
- **Validation styles**: Voir [analyse/TEMPLATE_VALIDATION_REPORT.md](analyse/TEMPLATE_VALIDATION_REPORT.md)

---

**Version**: 1.0  
**Dernière mise à jour**: Avril 2026  
**Auteur**: DC Formatter Team
