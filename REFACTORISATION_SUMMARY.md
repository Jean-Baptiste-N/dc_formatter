# 🔧 Refactorisation Complète - Résumé des Changements

## 📋 Vue d'ensemble
Le projet a été réorganisé en 5 modules distincts avec une architecture CLI claire via `pipeline.py`.

## 📁 Structure des Fichiers

### **tools3/extract_xml_raw.py**
- ✅ Extrait le XML brut d'un fichier DOCX
- Fonctions principales:
  - `export_all_xml(docx_file, output_folder)` - Crée le fichier `_GLOBAL.xml`
  - `extract_document_xml(docx_file, output_folder)` - Extrait le `word/document.xml`

### **tools3/parse_template.py**
- ✅ Extrait les dimensions du template DOCX une seule fois
- Fonctions principales:
  - `extract_page_dimensions_from_template(template_path)` - Retourne dict avec dimensions
- ✅ **CHANGEMENT**: `template_path` est maintenant obligatoire (pas de valeur par défaut)

### **tools3/parse_xml_raw_to_json_raw.py**
- ✅ Convertit le XML en JSON RAW
- Fonctions principales:
  - `xml_to_json(xml_file, json_output)` - Crée `_GLOBAL_raw.json`
- ✅ **CHANGEMENT**: `json_output` est maintenant obligatoire

### **tools3/process_json_raw_to_json_transformed.py**
- ✅ Applique tags et styles au JSON RAW
- Fonctions principales:
  - `apply_tags_and_styles(raw_json_file, output_dir, page_dimensions)` - Crée `_transformed.json`
- ✅ **CHANGEMENTS**:
  - `page_dimensions` remplace `template_path` (plus besoin d'extraire à chaque fois)
  - `output_dir` est maintenant obligatoire
  - Les page_dimensions sont stockées dans `data['page_dimensions']`
  - Les fonctions `create_edu_table()` et `create_xp_tables()` récupèrent les dimensions depuis data

### **tools3/render_json_transformed_to_docx.py**
- ✅ Rend le JSON transformé en DOCX
- Fonctions principales:
  - `json_to_docx(json_file, template_file, output_dir)` - Crée le DOCX final
- ✅ **CHANGEMENT**: `output_dir` est maintenant obligatoire

### **tools3/pipeline.py**
- ✅ **NOUVEAU**: Architecture CLI complète avec subcommandes

## 🎯 Structure CLI du Pipeline

### **Commandes Individuelles**
```bash
# Extraire les dimensions du template
python -m tools3.pipeline extract-dims template.docx

# Extraire le XML brut
python -m tools3.pipeline extract-xml document.docx [-o output_dir]

# Convertir XML en JSON RAW
python -m tools3.pipeline xml-to-json input.xml output.json

# Transformer le JSON RAW
python -m tools3.pipeline transform template.docx json_raw.json [-o output_dir]

# Rendre le JSON en DOCX
python -m tools3.pipeline render json_transformed.json template.docx output_dir
```

### **Commandes Composées**

#### Phase 1: Extraction (dims + xml + json raw)
```bash
python -m tools3.pipeline extract document.docx template.docx [-o structures]
```

#### Phase 2: Transformation + Rendu (après extraction)
```bash
python -m tools3.pipeline transform-render document.docx template.docx [-o structures]
```

#### Pipeline Complète
```bash
python -m tools3.pipeline full document.docx template.docx [-o structures]
```

## 🔄 Flux d'Exécution

### Approche Progressive
1. **Extraction**: `extract` → génère XML + JSON RAW
2. **Transformation**: `transform-render` → utilise le JSON RAW précédent

### Approche Globale
1. **Pipeline Complète**: `full` → tout d'un coup

## 💡 Points Clés de la Refactorisation

### ✅ Elimination des Paramètres Par Défaut
- Tous les paramètres importants sont maintenant obligatoires
- Plus de confusion sur les chemins par défaut

### ✅ Les Dimensions du Template
- Extraites **une seule fois** au début du pipeline
- Passées comme paramètre `page_dimensions` dict
- Stockées dans `data['page_dimensions']` pour accès ultérieur

### ✅ Harmoniséisation
- Toutes les fonctions suivent la même signature
- Pas de fallback silencieux ou de valeurs par défaut cachées
- Code plus prévisible et testable

### ✅ CLI Cohérente
- Structure claire avec subcommandes
- Messages de progression intuitifs
- Gestion d'erreurs cohérente

## 📝 Utilisation Recommandée

### Cas 1: Pipeline Complète (Simple)
```bash
python -m tools3.pipeline full test/BM2/input.docx assets/TEMPLATE.docx
```

### Cas 2: Traiter Plusieurs Documents (Efficace)
```bash
# Phase 1: Extraction pour tous
python -m tools3.pipeline extract test/BM2/input.docx assets/TEMPLATE.docx
python -m tools3.pipeline extract test/CBD/input.docx assets/TEMPLATE.docx

# Phase 2: Transformation (réutilise les JSON RAW)
python -m tools3.pipeline transform-render test/BM2/input.docx assets/TEMPLATE.docx
python -m tools3.pipeline transform-render test/CBD/input.docx assets/TEMPLATE.docx
```

## 🐛 Migration des Anciens Appels

### Avant:
```python
from tools3.process_json_raw_to_json_transformed import apply_tags_and_styles
apply_tags_and_styles('raw.json', 'output_dir')  # template_path inféré automatiquement
```

### Après:
```python
from tools3.parse_template import extract_page_dimensions_from_template
from tools3.process_json_raw_to_json_transformed import apply_tags_and_styles

dims = extract_page_dimensions_from_template('assets/TEMPLATE.docx')
apply_tags_and_styles('raw.json', 'output_dir', dims)
```
