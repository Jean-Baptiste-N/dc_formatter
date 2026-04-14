# 🎯 Guide Parsing XML Détaillé

## Vue d'ensemble

Le script `xml_to_json.py` peut maintenant extraire un document Word XML avec une hiérarchie complète et récupérer les propriétés CSS/formatage détaillées.

## 📝 Utilisation

### Parsing détaillé (défaut)
```bash
python3 tools3/xml_to_json.py structures/DC_JNZ_2026_RAW.xml
```

### Parsing simple (legacy)
```bash
python3 tools3/xml_to_json.py structures/DC_JNZ_2026_RAW.xml --simple
```

## 📊 Structure JSON générale

```json
{
  "document": {
    "type": "Document",
    "source": "DC_JNZ_2026_RAW.xml",
    "content": [
      {
        "index": 0,
        "type": "Paragraph",
        "properties": { ... },
        "runs": [ ... ]
      },
      {
        "index": 73,
        "type": "Table",
        "row_count": 4,
        "col_count": 2,
        "rows": [ ... ]
      }
    ],
    "stats": {
      "total_elements": 214,
      "paragraphs": 208,
      "tables": 6
    }
  }
}
```

## 🏷️ Propriétés de PARAGRAPHE capturées

Depuis `w:pPr` (Paragraph Properties):

| Propriété | Balise XML | Description |
|-----------|-----------|-------------|
| `paraId` | `w14:paraId` | Identifiant unique du paragraphe |
| `style` | `w:pStyle` | Style appliqué (ex: "Heading1", "Normal") |
| `alignment` | `w:jc` | Justification (left, center, right, both, etc) |
| `ilvl` | `w:ilvl` | Niveau de liste (0-8, ne s'affiche que si en liste) |
| `numId` | `w:numId` | ID instance numérotation (ne s'affiche que si en liste) |

### Exemple : Paragraphe avec liste
```json
{
  "index": 22,
  "type": "Paragraph",
  "properties": {
    "style": "Heading2",
    "ilvl": "0",
    "numId": "3",
    "paraId": "00000017",
    "alignment": "left"
  },
  "runs": [...]
}
```

## 🔤 Propriétés de RUN/TEXTE capturées

Depuis `w:rPr` (Run Properties):

| Propriété | Balise XML | Description |
|-----------|-----------|-------------|
| `text` | `w:t` | Contenu textuel |
| `bold` | `w:b` | True si gras |
| `italic` | `w:i` | True si italique |
| `size` | `w:sz` | Taille en half-points (ex: "40" = 20pt) |
| `color` | `w:color` | Couleur RGB hex (ex: "538cd3") |
| `font` | `w:rFonts` (ascii) | Police utilisée (ex: "Arial") |

### Exemple : Run avec formatage complexe
```json
{
  "properties": {
    "bold": true,
    "size": "40",
    "color": "538cd3",
    "font": "Arial"
  },
  "text": "DOSSIER DE COMPETENCES"
}
```

## 📦 Hiérarchie Paragraphe > Runs > Texte

```
Paragraphe (properties)
  ├── Run 1
  │   ├── properties (bold, italic, size, color, font)
  │   └── text
  ├── Run 2
  │   ├── properties
  │   └── text
  └── Run 3
      ├── properties
      └── text
```

### Exemple complet
```json
{
  "index": 5,
  "type": "Paragraph",
  "properties": {
    "alignment": "center",
    "paraId": "00000006"
  },
  "runs": [
    {
      "properties": {
        "bold": true,
        "size": "40",
        "color": "538cd3",
        "font": "Arial"
      },
      "text": "DOSSIER DE COMPETENCES"
    }
  ]
}
```

## 📋 Structure TABLEAU

```json
{
  "index": 73,
  "type": "Table",
  "row_count": 4,
  "col_count": 2,
  "rows": [
    {
      "row_index": 0,
      "cells": [
        {
          "col_index": 0,
          "paragraphs": [
            {
              "index": 0,
              "type": "Paragraph",
              "properties": {...},
              "runs": [...]
            }
          ]
        }
      ]
    }
  ]
}
```

```
Table
  ├── Row 0
  │   ├── Cell [0,0]
  │   │   └── Paragraph(s)
  │   │       └── Run(s)
  │   └── Cell [0,1]
  │       └── Paragraph(s)
  └── Row 1
      ├── Cell [1,0]
      └── Cell [1,1]
```

## 🔍 Listes et Hiérarchies

Les niveaux de liste sont capturés via `ilvl`:

```json
[
  {
    "type": "Paragraph",
    "properties": { "ilvl": "0", "numId": "1" },
    "runs": [{"text": "Niveau 1"}]
  },
  {
    "type": "Paragraph",
    "properties": { "ilvl": "1", "numId": "1" },
    "runs": [{"text": "  Niveau 2"}]
  },
  {
    "type": "Paragraph",
    "properties": { "ilvl": "0", "numId": "1" },
    "runs": [{"text": "Niveau 1 (retour)"}]
  }
]
```

## 🛠️ Utilisation en Python

### Charger le JSON
```python
import json

with open('structures/DC_JNZ_2026_RAW.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

content = data['document']['content']
```

### Filtrer par type
```python
# Tous les paragraphes
paragraphs = [e for e in content if e['type'] == 'Paragraph']

# Tous les tableaux
tables = [e for e in content if e['type'] == 'Table']
```

### Accéder aux propriétés
```python
para = content[5]

# Propriétés
print(para['properties']['paraId'])    # "00000006"
print(para['properties']['alignment']) # "center"

# Runs
for run in para['runs']:
    text = run['text']
    is_bold = run['properties'].get('bold', False)
    color = run['properties'].get('color')
    print(f"Text: {text}, Bold: {is_bold}, Color: {color}")
```

### Filtrer listes par niveau
```python
# Paragraphes de liste niveau 2
level_2_items = [
    e for e in content 
    if e.get('properties', {}).get('ilvl') == '1'
]
```

## 🎨 Propriétés disponibles

### Couleurs
Les couleurs sont en format RGB hexadécimal (ex: `"538cd3"` = bleu).

Conversions utiles:
```python
def hex_to_rgb(hex_color):
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

rgb = hex_to_rgb("538cd3")  # (83, 205, 211)
```

### Tailles
Les tailles sont en half-points (multiply par 2 pour avoir les points):
```python
size_half_points = 40
size_points = size_half_points / 2  # 20 pt
```

## ✅ Formats supportés

### Styles reconnus
- Normal, Heading1-Heading9
- List Bullet, List Number
- Styles personnalisés (DC_* etc)

### Alignements reconnus
- `left` - Aligné à gauche
- `center` - Centré
- `right` - Aligné à droite
- `both` - Justifié

### Niveaux de liste
- `ilvl`: 0-8 (9 niveaux possibles)
- Plus élevé = plus indentée

## 📌 De quoi vous avez besoin pour extraire

Pour enrichir davantage l'extraction, la fonction peut être modifiée pour capturer:
- `w:tab` - Tabulations
- `w:ind` - Indentation
- `w:spacing` - Espacement
- `w:strike` - Barré
- `w:underline` - Souligné
- Et d'autres propriétés Word

Demandez si vous en avez besoin !

## 🎯 Cas d'usage courants

### 1. Extraire tous les titres
```python
headings = [e for e in content 
    if 'Heading' in e.get('properties', {}).get('style', '')]
```

### 2. Extraire listes hiérarchiques
```python
for elem in content:
    if elem['type'] == 'Paragraph':
        ilvl = elem.get('properties', {}).get('ilvl')
        if ilvl:
            indent = "  " * int(ilvl)
            text = ''.join(r['text'] for r in elem.get('runs', []))
            print(f"{indent}• {text}")
```

### 3. Extraire avec formatage
```python
for elem in content:
    if elem['type'] == 'Paragraph':
        for run in elem.get('runs', []):
            text = run['text']
            prop = run.get('properties', {})
            if prop.get('bold'):
                text = f"**{text}**"
            if prop.get('italic'):
                text = f"*{text}*"
            print(text)
```
