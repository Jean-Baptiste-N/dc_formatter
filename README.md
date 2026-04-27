# DC Formatter - Pipeline de Traitement de Documents Word

## 📋 Présentation

**DC Formatter** est un outil complet pour extraire, transformer et générer des documents Word (.docx) avec une structure contrôlée. 

Il traite un document DOCX en 4 phases successives:
1. **Extraction** → XML global + dimensions du template
2. **Conversion** → JSON brut (structure complète du document)
3. **Transformation** → JSON enrichi (tags, styles, hiérarchie)
4. **Rendu** → DOCX final formaté selon le template

## 🚀 Démarrage Rapide

### Installation
```bash
pip install -r requirements.txt
```

### Pipeline Complet (recommandé)
```bash
python3 -m tools3.pipeline full -s DC_JNZ_2026.docx
```

### Résultats
Les fichiers générés sont organisés dans 4 dossiers:
- `OUTPUT1_XML-RAW/` → XML global brut
- `OUTPUT2_JSON-RAW/` → Structure JSON complète
- `OUTPUT3_JSON-TRANSFORMED/` → JSON avec tags et styles
- `OUTPUT4_DOCX-RESULT/` → Document Word final

## 📚 Documentation Complète

Voir [Z_DOCUMENTATION/README.md](Z_DOCUMENTATION/README.md) pour:
- Architecture détaillée du pipeline
- Description de chaque module Python
- Analyse des formats (XML, JSON, DOCX)
- Guide de détection des hiérarchies
- Commandes avancées

## 🛠️ Modules Python

| Module | Fonction |
|--------|----------|
| `extract_xml_raw.py` | Extraction XML depuis le DOCX |
| `parse_template.py` | Extraction dimensions et paramètres |
| `parse_xml_raw_to_json_raw.py` | Conversion XML → JSON brut |
| `process_json_raw_to_json_transformed.py` | Application tags/styles |
| `render_json_transformed_to_docx.py` | Génération DOCX final |
| `pipeline.py` | Interface CLI (orchestration) |
| `zip_docx.py` | Archivage et compression |

## 📊 Structure des Données

```
DOCX (zippé)
    ↓ extraction
XML brut (word/document.xml + relations)
    ↓ parsing + structuration
JSON RAW (paragraphes, tables, runs bruts)
    ↓ tagging + détection hiérarchie
JSON TRANSFORMED (avec structures sections + tags)
    ↓ rendu template
DOCX final (formaté + stylisé)
```

## 📝 Exemples Couramment Utilisés

```bash
# Pipeline complet
python3 -m tools3.pipeline full -s document.docx

# Avec output custom
python3 -m tools3.pipeline full -s document.docx -o results/

# Phase 1 seulement (extraction)
python3 -m tools3.pipeline extract -s document.docx

# Phase 2 seulement (transformation)
python3 -m tools3.pipeline transform-render -s document.docx

# Extraire dimensions du template
python3 -m tools3.pipeline extract-dims
```

## 🔧 Configuration

- **Template par défaut**: `assets/TEMPLATE.docx`
- **Sources DOCX**: Placez les fichiers dans `DC_SOURCES/` ou utilisez `-s chemin/complet`
- **Dossier sortie**: Paramètre `-o` (par défaut: OUTPUT*/)

## ⚡ Toutes les Commandes

Voir [Z_DOCUMENTATION/COMMANDES.sh](Z_DOCUMENTATION/COMMANDES.sh) pour la liste exhaustive avec exemples.
