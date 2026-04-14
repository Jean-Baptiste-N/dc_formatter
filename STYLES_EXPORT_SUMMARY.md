# 📋 Résumé Export DC_ Styles

**Date**: 2025-04-10  
**Source**: `assets/TEMPLATE_DC_new.docx`  
**Total DC_ styles extraits**: 23 (13 paragraphes + 10 caractères)

---

## 📂 Fichiers de Reference

### 1. `/styles/DC_styles.json` (20 KB)
**Format hybride**: Propriétés structurées + XML brut

```json
{
  "DCLigne": {
    "id": "DCLigne",
    "name": "DC_Ligne",
    "type": "paragraph",
    "customStyle": true,
    "properties": {
      "based_on": "DCNormal",
      "paragraph": {
        "tabs": [
          {
            "val": "right",
            "leader": "underscore",
            "pos": 496.15
          }
        ]
      },
      "character": {}
    },
    "xml_raw": "<ns0:style xmlns:ns0=\"...\" ... />"
  }
}
```

**Utilisation**:
- Propriétés structurées pour analyses/transformations : `data[styleId]['properties']`
- XML brut complet pour injection Word : `data[styleId]['xml_raw']`

---

### 2. `/styles/DC_styles.xml` (12 KB)
**Format**: XML Fragment injecrable

```xml
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
          xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="DCHDC">
    <w:name w:val="DC_H_DC" />
    <w:pPr>
      <w:spacing w:before="120" w:after="120" ... />
      <w:jc w:val="center" />
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" />
      <w:b />
      <w:color w:val="548DD4" />
      <w:sz w:val="40" />
    </w:rPr>
  </w:style>
  ...
</w:styles>
```

**Utilisation**: Injection directe dans `/word/styles.xml` de fichiers DOCX

---

### 3. `/styles/DC_styles_RAW.json` (14 KB)
**Format**: XML brut uniquement (legacy)

```json
{
  "DCHDC": {
    "id": "DCHDC",
    "name": "DC_H_DC",
    "xml_raw": "<ns0:style xmlns:ns0=\"...\" ... />"
  }
}
```

**Note**: Remplacé par DC_styles.json (plus complet)

---

## 🎨 Styles Disponibles (23)

### Headers (4 styles)
- `DC_H_DC` → `DC_H_DC Car` (Bleu 548DD4, 20pt, **gras**)
- `DC_H_Trigramme` → `DC_H_Trigramme Car` (Orange F69545, **gras**)
- `DC_H_Poste` → `DC_H_Poste Car` (18pt, **gras**)
- `DC_H_XP` → `DC_H_XP Car` (Orange F69545, 18pt, **gras**)

### Bullets/Numérotation (4 styles)
- `DC_1st_bullet` (Plus haut niveau, Orange EC7C2F, 12pt)
- `DC_2nd_bullet` (Niveau 2, 10pt, **gras**)
- `DC_3rd_bullet` (Niveau 3, 9pt)
- `DC_4th_bullet` (Niveau 4, 8pt)

### Sections (1 style)
- `DC_T1_Sections` → `DC_T1_Sections Car` (Bleu 538CD4, 16pt, **gras**)

### Tables (3 styles)
- `DC_Table_BlueTitle` (En-tête bleu, **gras**, 538CD4)
- `DC_Table_Content` (Contenu, noir 000000)
- `DC_Table_Year` (Année, **gras**)

### Experience (4 styles)
- `DC_XP_Title` (Titre, gris 808080, 16pt)
- `DC_XP_Date` (Date, gris 808080, 12pt, **gras**)
- `DC_XP_Poste` (Poste, noir, 12pt, **gras**)
- `DC_XP_BlueContent` (Contenu bleu, 538CD4, 12pt)

### Autres (2 styles)
- `DC_Normal` (Normal, 10pt)
- `DC_Ligne` (Ligne tiretée tabulation)

---

## 🔧 Propriétés Extraites

### Propriétés de Paragraphe
- **alignment**: center, left, right, both
- **spacing_before/after**: En points (TWIPS/20)
- **left_indent**: En points
- **tabs**: Liste de `{val, leader, pos}`
  - `val`: `left|center|right|decimal`
  - `leader`: `none|dot|hyphen|heavy|middleDot|underscore`
  - `pos`: Position en points

### Propriétés de Caractère
- **font_name**: Arial, etc.
- **font_size**: En points
- **bold**: True/False
- **italic**: True/False
- **color_hex**: Code hex (548DD4, F69545, etc.)

---

## 💾 Structure JSON Complète

```typescript
interface StyleExport {
  [styleId: string]: {
    id: string                    // Identifiant Word (ex: "DCHDC")
    name: string                  // Nom affiché (ex: "DC_H_DC")
    type: "paragraph" | "character"
    customStyle: boolean          // Toujours true pour DC_
    properties: {
      based_on: string | null     // Style parent
      paragraph: {                // Seulement pour type="paragraph"
        alignment?: string
        spacing_before?: number
        spacing_after?: number
        left_indent?: number
        tabs?: Array<{
          val: string
          leader: string
          pos: number
        }>
      }
      character: {                // Seulement pour type="character"
        font_name?: string
        font_size?: number
        bold?: boolean
        italic?: boolean
        color_hex?: string
      }
    }
    xml_raw: string               // XML Word brut natif
  }
}
```

---

## 🛠 Usage Patterns

### 1. Accéder aux propriétés parsées
```python
import json

data = json.load(open('styles/DC_styles.json'))
style = data['DCLigne']

# Extraire les tabulations
tabs = style['properties']['paragraph'].get('tabs', [])
for tab in tabs:
    print(f"Tab: {tab['val']} @ {tab['pos']}pt ({tab['leader']})")
```

### 2. Récupérer le XML brut
```python
xml_string = style['xml_raw']
# Peut être parsé/injecté directement dans Word
```

### 3. Régénérer après modifications
```bash
python3 tools/detect_styles_v2.py assets/TEMPLATE_DC_new.docx dc
```

---

## ✅ Checklist Completeness

- [x] Tous les 23 DC_ styles détectés
- [x] Propriétés de paragraphe capturées (spacing, indent, tabs, alignment)
- [x] Propriétés de caractère capturées (font, size, bold, color)
- [x] XML brut préservé pour 100% fidelité
- [x] Export JSON hybride (parsing + XML)
- [x] Export XML fragment injecrable
- [x] Namespaces gérés correctement (ns0: vs w:)
- [x] Conversion TWIPS → Points

---

## 📝 Notes Techniques

### Namespace Prefixes
- `ns0:` et `w:` réfèrent au même namespace
- URI: `http://schemas.openxmlformats.org/wordprocessingml/2006/main`
- Préfixe cosmétique, l'URI est ce qui compte

### Unités
- TWIPS: 1/20 de point (Word interne)
- Points: Unité standard (TWIPS ÷ 20)
- Toutes les exportations en points

### Héritage Styles
- Certains styles basés sur d'autres (ex: `DC_2nd_bullet` basé sur `DC_1st_bullet`)
- Voir `based_on` pour la hiérarchie complète

---

**Prochaines étapes**:
1. Créer un outil d'injection pour appliquer ces styles à d'autres fichiers
2. Intégrer les vérifications de style dans `parse_reformat.py`
3. Générer template de style code via `generate_style_code.py`
