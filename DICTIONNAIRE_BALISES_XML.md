# 📚 Dictionnaire des Balises XML Word (WordprocessingML)

## 📖 Guide d'Utilisation

Ce document explique **toutes les balises** présentes dans le fichier XML brut d'un document Word (.docx).

**Format du dictionnaire**:
```
<w:balise> — CATÉGORIE
Signification: Explication claire
Attributs courants: attr1, attr2
Exemple: <w:balise w:attr="valeur">
```

---

## 🎯 Balises Principales (Structure de Document)

### `<w:document>` — STRUCTURE
Signification: Élément racine du document Word. Contient tout le contenu du fichier.
Attributs courants: `xmlns` (déclarations de namespace)
Exemple: 
```xml
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <!-- contenu -->
</w:document>
```

### `<w:body>` — STRUCTURE
Signification: Corps du document. Contient tous les paragraphes, tableaux et sections.
Attributs courants: aucun
Exemple: 
```xml
<w:body>
  <w:p>...</w:p>
  <w:tbl>...</w:tbl>
</w:body>
```

---

## 📝 Balises de Paragraphes

### `<w:p>` — CONTAINER (Paragraphe)
Signification: **Paragraph** - Unité de texte dans le document. Chaque ligne/bloc est un `<w:p>`.
Attributs courants: `w:rsidR`, `w:rsidDel`, `w:rsidP`, `w14:paraId`
Exemple: 
```xml
<w:p w:rsidR="00000000" w14:paraId="00000001">
  <w:pPr>...</w:pPr>
  <w:r>...</w:r>
</w:p>
```

### `<w:pPr>` — PROPERTIES (Propriétés de Paragraphe)
Signification: **Paragraph Properties** - Contient tous les formatages du paragraphe (alignment, indentation, espacement, etc.)
Attributs courants: aucun (c'est un container)
Exemple: 
```xml
<w:pPr>
  <w:pStyle w:val="Heading1"/>
  <w:jc w:val="center"/>
  <w:ind w:left="1080" w:right="0"/>
</w:pPr>
```

### `<w:pStyle>` — STYLE
Signification: **Paragraph Style** - Nom du style appliqué au paragraphe (Normal, Heading1, Heading2, etc.)
Attributs courants: `w:val="StyleName"`
Exemple: 
```xml
<w:pStyle w:val="Heading1"/>
<w:pStyle w:val="Normal"/>
```

---

## 🔤 Balises de Runs (Texte avec Formatage)

### `<w:r>` — CONTAINER (Run)
Signification: **Run** - Série de caractères avec le même formatage. Combine le texte `<w:t>` et ses propriétés `<w:rPr>`.
Attributs courants: `w:rsidR`, `w:rsidRPr`, `w:rsidDel`
Exemple: 
```xml
<w:r>
  <w:rPr>
    <w:b/>
    <w:sz w:val="40"/>
  </w:rPr>
  <w:t>Texte en gras</w:t>
</w:r>
```

### `<w:t>` — TEXT (Texte)
Signification: **Text** - Contient le texte réel du document. Habituellement dans un `<w:r>`.
Attributs courants: `xml:space="preserve"` (préserve les espaces)
Exemple: 
```xml
<w:t xml:space="preserve">DOSSIER DE COMPETENCES</w:t>
```

### `<w:rPr>` — PROPERTIES (Propriétés de Run)
Signification: **Run Properties** - Formatages du texte (gras, italique, taille, couleur, police, etc.)
Attributs courants: aucun (c'est un container)
Exemple: 
```xml
<w:rPr>
  <w:b w:val="1"/>
  <w:sz w:val="40"/>
  <w:color w:val="538cd3"/>
</w:rPr>
```

### `<w:br>` — BREAK
Signification: **Break** - Saut de ligne dans un paragraphe.
Attributs courants: aucun
Exemple: 
```xml
<w:br/>
<w:br w:type="page"/>  <!-- saut de page -->
```

---

## 🎨 Balises de Formatage (dans `<w:rPr>`)

### `<w:b>` — BOLD
Signification: **Bold** - Texte en gras.
Attributs courants: `w:val="0"` (désactiver), `w:val="1"` (activer), ou absente = activé
Exemple: 
```xml
<w:b/>              <!-- Activé (présence = true) -->
<w:b w:val="1"/>    <!-- Explicitement activé -->
<w:b w:val="0"/>    <!-- Désactivé -->
```

### `<w:bCs>` — BOLD COMPLEX SCRIPT
Signification: **Bold Complex Script** - Gras pour écritures complexes (arabe, hébreu, etc.)
Attributs courants: `w:val="0"` ou `w:val="1"`
Exemple: 
```xml
<w:bCs w:val="1"/>
```

### `<w:i>` — ITALIC
Signification: **Italic** - Texte en italique.
Attributs courants: `w:val="0"` ou `w:val="1"`
Exemple: 
```xml
<w:i/>
<w:i w:val="1"/>
```

### `<w:iCs>` — ITALIC COMPLEX SCRIPT
Signification: **Italic Complex Script** - Italique pour écritures complexes.
Attributs courants: `w:val="0"` ou `w:val="1"`

### `<w:u>` — UNDERLINE
Signification: **Underline** - Texte souligné.
Attributs courants: `w:val="single"`, `w:val="double"`, `w:val="none"`, etc.
Exemple: 
```xml
<w:u w:val="single"/>   <!-- Souligné simple -->
<w:u w:val="double"/>   <!-- Souligné double -->
<w:u w:val="none"/>     <!-- Non souligné -->
```

### `<w:strike>` — STRIKE
Signification: **Strike** - Texte barré / rayé.
Attributs courants: `w:val="0"` ou `w:val="1"`
Exemple: 
```xml
<w:strike w:val="1"/>
```

### `<w:smallCaps>` — SMALL CAPS
Signification: **Small Capitals** - Texte en petites majuscules.
Attributs courants: `w:val="0"` ou `w:val="1"`
Exemple: 
```xml
<w:smallCaps w:val="1"/>
```

### `<w:sz>` — SIZE
Signification: **Size** - Taille de police en **demi-points** (donc 40 = 20pt, 24 = 12pt).
Attributs courants: `w:val="nombre"`
Exemple: 
```xml
<w:sz w:val="40"/>      <!-- 20 points -->
<w:sz w:val="24"/>      <!-- 12 points -->
<w:sz w:val="22"/>      <!-- 11 points -->
```

### `<w:szCs>` — SIZE COMPLEX SCRIPT
Signification: **Size Complex Script** - Taille pour écritures complexes.
Attributs courants: `w:val="nombre"`

### `<w:color>` — COLOR
Signification: **Color** - Couleur du texte en RGB hexadécimal.
Attributs courants: `w:val="XXXXXX"` (hexa RGB)
Exemple: 
```xml
<w:color w:val="000000"/>   <!-- Noir -->
<w:color w:val="FFFFFF"/>   <!-- Blanc -->
<w:color w:val="538cd3"/>   <!-- Bleu -->
<w:color w:val="ec7c2f"/>   <!-- Orange -->
<w:color w:val="FF0000"/>   <!-- Rouge -->
```

### `<w:rFonts>` — FONTS
Signification: **Run Fonts** - Police de caractères à utiliser.
Attributs courants: `w:ascii="Font"`, `w:hAnsi="Font"`, `w:cs="Font"`, `w:eastAsia="Font"`
Exemple: 
```xml
<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
```

### `<w:vertAlign>` — VERTICAL ALIGN
Signification: **Vertical Alignment** - Alignement vertical du texte (baseline, superscript, subscript).
Attributs courants: `w:val="baseline"`, `w:val="superscript"`, `w:val="subscript"`
Exemple: 
```xml
<w:vertAlign w:val="baseline"/>     <!-- Normal -->
<w:vertAlign w:val="superscript"/>  <!-- Exposant -->
<w:vertAlign w:val="subscript"/>    <!-- Indice -->
```

### `<w:shd>` — SHADING
Signification: **Shading** - Couleur de fond / surbrillance du texte.
Attributs courants: `w:val="clear"`, `w:fill="XXXXXX"`
Exemple: 
```xml
<w:shd w:fill="FFFF00"/>    <!-- Fond jaune -->
<w:shd w:val="clear" w:fill="auto"/>
```

### `<w:rtl>` — RIGHT-TO-LEFT
Signification: **Right-to-Left** - Direction de texte droite-à-gauche (pour arabe, hébreu, etc.)
Attributs courants: `w:val="0"` (gauche-à-droite) ou `w:val="1"` (droite-à-gauche)
Exemple: 
```xml
<w:rtl w:val="1"/>
```

---

## 📐 Balises d'Alignement et Indentation (dans `<w:pPr>`)

### `<w:jc>` — JUSTIFICATION (Alignment)
Signification: **Justification / Alignment** - Alignement horizontal du paragraphe.
Attributs courants: `w:val` = "left", "center", "right", "justify"
Exemple: 
```xml
<w:jc w:val="left"/>      <!-- Aligné à gauche -->
<w:jc w:val="center"/>    <!-- Centré -->
<w:jc w:val="right"/>     <!-- Aligné à droite -->
<w:jc w:val="justify"/>   <!-- Justifié -->
```

### `<w:ind>` — INDENTATION
Signification: **Indentation** - Espacements de paragraphe (gauche, droite, première ligne).
Attributs courants: `w:left`, `w:right`, `w:firstLine`, `w:hanging`
Exemple: 
```xml
<w:ind w:left="1080" w:right="0" w:firstLine="0"/>
<!-- left=1080 twips ≈ 1 inches, firstLine=0 = pas d'indentation de première ligne -->

<w:ind w:left="1080" w:hanging="360"/>
<!-- hanging = indentation négative (première ligne moins indentée) -->
```

### `<w:spacing>` — SPACING
Signification: **Spacing** - Espacement avant/après le paragraphe et interligne.
Attributs courants: `w:before`, `w:after`, `w:line`, `w:lineRule`
Exemple: 
```xml
<w:spacing w:before="240" w:after="0" w:line="240" w:lineRule="auto"/>
<!-- before/after/line en twips (1 twip = 1/20 point) -->
```

### `<w:numPr>` — NUMBERING PROPERTIES
Signification: **Numbering Properties** - Propriétés de liste/numérotation du paragraphe.
Attributs courants: aucun (container)
Contient: `<w:ilvl>` et `<w:numId>`
Exemple: 
```xml
<w:numPr>
  <w:ilvl w:val="0"/>    <!-- Niveau 0 = principal list item -->
  <w:numId w:val="3"/>   <!-- ID du style de numérotation -->
</w:numPr>
```

### `<w:ilvl>` — INDENTATION LEVEL
Signification: **Indentation Level** - Niveau de profondeur dans une liste (0 = principal, 1 = sous-item, etc.).
Attributs courants: `w:val="nombre"`
Exemple: 
```xml
<w:ilvl w:val="0"/>   <!-- Niveau principal -->
<w:ilvl w:val="1"/>   <!-- Sous-niveau 1 -->
<w:ilvl w:val="2"/>   <!-- Sous-niveau 2 -->
```

### `<w:numId>` — NUMBER ID
Signification: **Number ID** - Identifiant du style de numérotation/liste à utiliser.
Attributs courants: `w:val="nombre"`
Exemple: 
```xml
<w:numId w:val="3"/>
<w:numId w:val="4"/>
```

---

## 🔲 Balises de Bordures et Ombrage (dans `<w:pPr>`)

### `<w:pBdr>` — PARAGRAPH BORDERS
Signification: **Paragraph Borders** - Bordures autour du paragraphe.
Attributs courants: aucun (container)
Contient: `<w:top>`, `<w:left>`, `<w:bottom>`, `<w:right>`, `<w:between>`
Exemple: 
```xml
<w:pBdr>
  <w:top w:val="single" w:sz="12" w:space="1" w:color="000000"/>
  <w:bottom w:val="single" w:sz="12" w:space="1" w:color="000000"/>
</w:pBdr>
```

### `<w:top>`, `<w:bottom>`, `<w:left>`, `<w:right>`, `<w:between>` — BORDER SIDES
Signification: Bordures individuelles (haut, bas, gauche, droite, entre lignes).
Attributs courants: `w:val` (type: "nil", "single", "double", etc.), `w:sz` (épaisseur), `w:space`, `w:color`, `w:shadow`
Exemple: 
```xml
<w:top w:space="0" w:sz="0" w:val="nil"/>  <!-- Pas de bordure -->
```

### `<w:shd>` — SHADING (Fond du paragraphe)
Signification: **Shading** - Couleur de fond du paragraphe.
Attributs courants: `w:val="clear"`, `w:fill="XXXXXX"`
Exemple: 
```xml
<w:shd w:fill="auto" w:val="clear"/>
```

---

## 📋 Balises de Tableaux

### `<w:tbl>` — TABLE
Signification: **Table** - Conteneur pour un tableau complet.
Attributs courants: aucun
Exemple: 
```xml
<w:tbl>
  <w:tblPr>...</w:tblPr>
  <w:tblGrid>...</w:tblGrid>
  <w:tr>...</w:tr>
</w:tbl>
```

### `<w:tblPr>` — TABLE PROPERTIES
Signification: **Table Properties** - Propriétés du tableau (style, largeur, etc.).
Attributs courants: aucun (container)

### `<w:tblStyle>` — TABLE STYLE
Signification: **Table Style** - Style du tableau.
Attributs courants: `w:val="StyleName"`

### `<w:tblW>` — TABLE WIDTH
Signification: **Table Width** - Largeur du tableau.
Attributs courants: `w:w="nombre"`, `w:type="auto"` ou `"dxa"`

### `<w:tblGrid>` — TABLE GRID
Signification: **Table Grid** - Définition des colonnes du tableau.
Attributs courants: aucun
Contient: `<w:gridCol>`

### `<w:gridCol>` — GRID COLUMN
Signification: **Grid Column** - Définition d'une colonne.
Attributs courants: `w:w="largeur"`

### `<w:tr>` — TABLE ROW
Signification: **Table Row** - Ligne du tableau.
Attributs courants: aucun
Contient: `<w:trPr>` et `<w:tc>` (cellules)

### `<w:trPr>` — TABLE ROW PROPERTIES
Signification: **Table Row Properties** - Propriétés de la ligne.
Attributs courants: aucun

### `<w:trHeight>` — TABLE ROW HEIGHT
Signification: **Table Row Height** - Hauteur de la ligne.
Attributs courants: `w:val="hauteur"`, `w:type="auto"` ou `"atLeast"`

### `<w:tc>` — TABLE CELL
Signification: **Table Cell** - Cellule du tableau.
Attributs courants: aucun
Contient: `<w:tcPr>` et `<w:p>` (paragraphes)

### `<w:tcPr>` — TABLE CELL PROPERTIES
Signification: **Table Cell Properties** - Propriétés de la cellule.
Attributs courants: aucun

### `<w:tcBorders>` — TABLE CELL BORDERS
Signification: **Table Cell Borders** - Bordures de la cellule.
Attributs courants: aucun

### `<w:tblLayout>` — TABLE LAYOUT
Signification: **Table Layout** - Type de layout (autofit ou fixed).
Attributs courants: `w:type="autofit"` ou `"fixed"`

### `<w:tblLook>` — TABLE LOOK
Signification: **Table Look** - Options d'affichage du tableau (en-têtes, etc.).
Attributs courants: `w:val="XXXX"` (hex flags)

### `<w:tblInd>` — TABLE INDENTATION
Signification: **Table Indentation** - Indentation du tableau.
Attributs courants: `w:w="valeur"`, `w:type="dxa"`

### `<w:tblHeader>` — TABLE HEADER
Signification: **Table Header** - Marque une ligne comme en-tête.
Attributs courants: aucun

### `<w:vAlign>` — VERTICAL ALIGNMENT (Cellule)
Signification: **Vertical Alignment** - Alignement vertical du contenu de la cellule.
Attributs courants: `w:val="top"`, `w:val="center"`, `w:val="bottom"`

---

## 📄 Balises de Sections et Pages

### `<w:sectPr>` — SECTION PROPERTIES
Signification: **Section Properties** - Propriétés de section/page (marges, en-têtes, pieds de page, etc.).
Attributs courants: aucun
Contient: `<w:pgSz>`, `<w:pgMar>`, `<w:headerReference>`, etc.
Exemple: 
```xml
<w:sectPr>
  <w:headerReference r:id="rId8" w:type="default"/>
  <w:pgSz w:h="16840" w:w="11920" w:orient="portrait"/>
  <w:pgMar w:top="1760" w:bottom="280" w:left="992" w:right="1133"/>
</w:sectPr>
```

### `<w:pgSz>` — PAGE SIZE
Signification: **Page Size** - Dimensions de la page.
Attributs courants: `w:w="largeur"`, `w:h="hauteur"`, `w:orient="portrait"` ou `"landscape"`
Conversion: 1 inch = 1440 twips, A4 ≈ 11920×16840 twips
Exemple: 
```xml
<w:pgSz w:w="11920" w:h="16840"/>  <!-- A4 en portrait -->
```

### `<w:pgMar>` — PAGE MARGINS
Signification: **Page Margins** - Marges de la page.
Attributs courants: `w:top`, `w:bottom`, `w:left`, `w:right`, `w:header`, `w:footer`
Unité: twips (1 twip = 1/20 point)
Exemple: 
```xml
<w:pgMar w:top="1760" w:bottom="280" w:left="992" w:right="1133" w:header="720" w:footer="720"/>
```

### `<w:headerReference>` — HEADER REFERENCE
Signification: **Header Reference** - Référence à l'en-tête de la section.
Attributs courants: `r:id="rIdXX"`, `w:type="default"`, `"first"`, `"even"`
Exemple: 
```xml
<w:headerReference r:id="rId8" w:type="default"/>
<w:headerReference r:id="rId9" w:type="first"/>
```

### `<w:pgNumType>` — PAGE NUMBER TYPE
Signification: **Page Number Type** - Format et départ de la numérotation.
Attributs courants: `w:start="nombre"`
Exemple: 
```xml
<w:pgNumType w:start="1"/>
```

### `<w:pageBreakBefore>` — PAGE BREAK BEFORE
Signification: **Page Break Before** - Force un saut de page avant ce paragraphe.
Attributs courants: `w:val="0"` ou `w:val="1"`
Exemple: 
```xml
<w:pageBreakBefore w:val="1"/>
```

---

## 📌 Balises de Signets et Marques

### `<w:bookmarkStart>` — BOOKMARK START
Signification: **Bookmark Start** - Début d'un signet/marque dans le document.
Attributs courants: `w:id="nombre"`, `w:name="NomDuSignet"`
Exemple: 
```xml
<w:bookmarkStart w:id="0" w:name="TableOfContents"/>
```

### `<w:bookmarkEnd>` — BOOKMARK END
Signification: **Bookmark End** - Fin d'un signet/marque.
Attributs courants: `w:id="nombre"` (doit correspondre au bookmarkStart)
Exemple: 
```xml
<w:bookmarkEnd w:id="0"/>
```

---

## 🖼️ Balises de Contenu Spécial

### `<w:drawing>` — DRAWING
Signification: **Drawing** - Conteneur pour des éléments dessinés (images, formes, etc.).
Attributs courants: aucun
Contient: `<wp:inline>` ou `<wp:anchor>`

### `<w:tabs>` — TABS
Signification: **Tabs** - Définition des tabulations du paragraphe.
Attributs courants: aucun
Contient: `<w:tab>`

### `<w:tab>` — TAB
Signification: **Tab** - Définition d'une tabulation.
Attributs courants: `w:val="left"`, `"center"`, `"right"`, `"decimal"`, `w:pos="position"`, `w:leader="..."` (style du leader)
Exemple: 
```xml
<w:tab w:val="left" w:leader="none" w:pos="1354"/>
<w:tab w:val="right" w:leader="dot" w:pos="9180"/>  <!-- Leader avec points -->
```

### `<w:gridSpan>` — GRID SPAN
Signification: **Grid Span** - Continue une cellule sur plusieurs colonnes.
Attributs courants: `w:val="nombre"`
Exemple: 
```xml
<w:gridSpan w:val="2"/>  <!-- Fusion de 2 colonnes -->
```

---

## 🔗 Balises de Contrôle

### `<w:keepNext>` — KEEP NEXT
Signification: **Keep Next** - Garde ce paragraphe avec le suivant (ne pas les séparer par un saut de page).
Attributs courants: `w:val="0"` ou `w:val="1"`
Exemple: 
```xml
<w:keepNext w:val="1"/>  <!-- Garder ensemble -->
```

### `<w:keepLines>` — KEEP LINES
Signification: **Keep Lines** - Garde toutes les lignes ensemble sur une page.
Attributs courants: `w:val="0"` ou `w:val="1"`
Exemple: 
```xml
<w:keepLines w:val="1"/>
```

### `<w:widowControl>` — WIDOW CONTROL
Signification: **Widow Control** - Évite les "veuves" = dernière ligne isolée en haut de page.
Attributs courants: `w:val="0"` ou `w:val="1"`
Exemple: 
```xml
<w:widowControl w:val="1"/>
```

### `<w:cantSplit>` — CAN'T SPLIT
Signification: **Can't Split** - Empêche la division d'un tableau ou paragraphe.
Attributs courants: `w:val="0"` ou `w:val="1"`
Exemple: 
```xml
<w:cantSplit w:val="1"/>
```

##############################################################################################

##############################################################################################

## 🏷️ Dictionnaire Complet des Attributs

### Attributs Structurels et ID

#### `w:id` — IDENTIFIANT
Signification: Identifiant unique d'un élément (signet, bookmark, etc.).
Valeurs typiques: Nombre entier (0, 1, 2, etc.)
Exemple: 
```xml
<w:bookmarkStart w:id="0" w:name="TOC"/>
<w:bookmarkEnd w:id="0"/>
<w:numId w:val="3"/>
```

#### `w:name` — NAME
Signification: Nom donné à un élément (signet, style, etc.).
Valeurs typiques: Chaîne de caractères
Exemple: 
```xml
<w:pStyle w:val="Heading1"/>
<w:bookmarkStart w:name="TableOfContents"/>
```

#### `w:val` — VALUE (Valeur)
Signification: Valeur générique pour un attribut. Sens dépend du contexte.
Contextes:
- **Pour `<w:jc>`**: "left", "center", "right", "justify"
- **Pour `<w:b>`, `<w:i>`, `<w:u>`**: "0", "1" ou absent=vrai
- **Pour `<w:sz>`**: nombre (demi-points)
- **Pour `<w:pStyle>`**: Nom du style (Heading1, Normal, etc.)
- **Pour `<w:u>`**: "single", "double", "none", "wave", "dotted", etc.
- **Pour `<w:type>`**: "page", "column", "continuous", etc.
Exemple: 
```xml
<w:jc w:val="center"/>
<w:sz w:val="40"/>
<w:pStyle w:val="Heading1"/>
<w:b w:val="1"/>
```

#### `r:id` — RELATION ID
Signification: ID de relation externe (références fichiers, images, liens).
Format: "rIdXX" où XX est un nombre
Exemple: 
```xml
<w:headerReference r:id="rId8" w:type="default"/>
<w:drawing><a:blip r:embed="rId4"/></w:drawing>
```

---

### Attributs de Dimensions (Spacing, Positionnement)

#### `w:before` — SPACING BEFORE (Espace Avant)
Signification: Espace/espacement avant le paragraphe/élément.
Unité: twips (1 twip = 1/20 point)
Valeurs typiques: 0, 120, 240, 288, 341 (multiples de espacements courants)
Exemple: 
```xml
<w:spacing w:before="240"/>    <!-- 12 points de marge avant -->
<w:spacing w:before="1"/>      <!-- Minimal -->
```

#### `w:after` — SPACING AFTER (Espace Après)
Signification: Espace/espacement après le paragraphe/élément.
Unité: twips
Exemple: 
```xml
<w:spacing w:after="0"/>       <!-- Pas d'espace après -->
<w:spacing w:after="120"/>     <!-- 6 points après -->
```

#### `w:line` — LINE SPACING (Interligne)
Signification: Espacement entre les lignes du paragraphe.
Unité: twips (généralement 240 = simple, 480 = double)
Fonctionne avec `w:lineRule`
Exemple: 
```xml
<w:spacing w:line="240" w:lineRule="auto"/>   <!-- Simple -->
<w:spacing w:line="480" w:lineRule="auto"/>   <!-- Double -->
```

#### `w:lineRule` — LINE RULE (Règle Interligne)
Signification: Mode de calcul de l'interligne.
Valeurs possibles:
- `"auto"` — Automatique (parole standard)
- `"exact"` — Exactement la valeur donnée
- `"atLeast"` — Minimum (au moins)
Exemple: 
```xml
<w:spacing w:line="240" w:lineRule="auto"/>
<w:spacing w:line="300" w:lineRule="exact"/>
```

#### `w:pos` — POSITION (Position de Tabulation)
Signification: Position d'une tabulation dans le paragraphe.
Unité: twips depuis la marge gauche
Exemple: 
```xml
<w:tab w:val="left" w:pos="1354"/>     <!-- Tabulation à ~0.94 inches -->
<w:tab w:val="right" w:pos="9180"/>    <!-- Tab droite à ~6.38 inches -->
```

#### `w:left` — LEFT INDENTATION (Indentation Gauche)
Signification: Indentation/marge gauche du paragraphe.
Unité: twips
Exemple: 
```xml
<w:ind w:left="1080"/>         <!-- ~0.75 inches -->
<w:ind w:left="1602"/>         <!-- ~1.11 inches -->
```

#### `w:right` — RIGHT INDENTATION (Indentation Droite)
Signification: Indentation/marge droite du paragraphe.
Unité: twips
Exemple: 
```xml
<w:ind w:right="0"/>
<w:ind w:right="1476"/>
```

#### `w:firstLine` — FIRST LINE INDENTATION (Indentation Première Ligne)
Signification: Indentation supplémentaire de la première ligne.
Unité: twips (peut être négatif)
Exemple: 
```xml
<w:ind w:firstLine="424"/>     <!-- Première ligne indentée en plus -->
<w:ind w:firstLine="0"/>       <!-- Pas d'indentation de 1ère ligne -->
```

#### `w:hanging` — HANGING INDENTATION (Indentation Négative)
Signification: Indentation négative pour les listes/bullets (première ligne moins indentée).
Unité: twips
Exemple: 
```xml
<w:ind w:left="1080" w:hanging="360"/>
<!-- Première ligne à 1080-360=720, autres lignes à 1080 -->
```

#### `w:top`, `w:bottom`, `w:left`, `w:right` — BORDER SPACING
Signification: Épaisseur/taille de bordure.
Unité: huitièmes de point (donc 24 = 3 points)
Exemple: 
```xml
<w:top w:space="0" w:sz="12" w:val="single"/>
```

#### `w:space` — SPACING (Espace)
Signification: Espace/distance de la bordure au contenu.
Unité: points (pt)
Exemple: 
```xml
<w:top w:space="1" w:sz="12"/>
<w:between w:space="0"/>
```

---

### Attributs de Formatage et Apparence

#### `w:sz` — SIZE (Taille)
Signification: Taille de police.
Unité: **demi-points** (donc 24 = 12pt, 40 = 20pt)
Exemple: 
```xml
<w:sz w:val="20"/>     <!-- 10 points -->
<w:sz w:val="24"/>     <!-- 12 points -->
<w:sz w:val="40"/>     <!-- 20 points (Heading) -->
```

#### `w:color` — COLOR (Couleur)
Signification: Couleur du texte.
Format: Hexadécimal RGB (RRGGBB)
Valeurs courantes:
- `"000000"` — Noir
- `"FFFFFF"` — Blanc
- `"FF0000"` — Rouge
- `"538CD3"` ou `"538cd3"` — Bleu
- `"EC7C2F"` ou `"ec7c2f"` — Orange
Exemple: 
```xml
<w:color w:val="538cd3"/>
<w:color w:val="ec7c2f"/>
<w:color w:val="000000"/>
```

#### `w:fill` — FILL (Remplissage/Fond)
Signification: Couleur de fond/surlignage.
Format: Hexadécimal RGB
Exemple: 
```xml
<w:shd w:fill="FFFF00"/>   <!-- Fond jaune -->
<w:shd w:val="clear" w:fill="auto"/>
```

#### `w:leader` — LEADER (Caractère Répétitif)
Signification: Caractère à répéter pour remplir une tabulation.
Valeurs possibles:
- `"none"` — Pas de caractère
- `"dot"` — Tirets/points: . . . . .
- `"heavy"` — Tirets lourds
- `"hyphen"` — Tirets: - - - - -
Exemple: 
```xml
<w:tab w:val="right" w:leader="dot" w:pos="9180"/>
<!-- Créé des points jusqu'à la tabulation -->
```

---

### Attributs de Fonts/Polices

#### `w:ascii` — ASCII FONT
Signification: Police pour caractères ASCII (Latin).
Exemple: 
```xml
<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
```

#### `w:hAnsi` — HIGH ANSI FONT
Signification: Police pour caractères "High ANSI" (Europe occidentale).
Exemple: 
```xml
<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
```

#### `w:cs` — COMPLEX SCRIPT FONT
Signification: Police pour écritures complexes (arabe, hébreu, asiatiques).
Exemple: 
```xml
<w:rFonts w:ascii="Arial" w:cs="Arial" w:eastAsia="SimSun"/>
```

#### `w:eastAsia` — EAST ASIA FONT
Signification: Police pour caractères d'Asie de l'Est (chinois, japonais, coréen).
Exemple: 
```xml
<w:rFonts w:eastAsia="SimSun"/>
<w:rFonts w:eastAsia="Calibri"/>
```

---

### Attributs de Pages et Sections

#### `w:w` — WIDTH (Largeur)
Signification: Largeur de page, tableau, ou cellule.
Unité: Dépend de contexte (twips, DXAs, ou auto)
Exemple: 
```xml
<w:pgSz w:w="11920"/>               <!-- Largeur page A4 -->
<w:tblW w:w="5000" w:type="dxa"/>   <!-- Largeur tableau -->
```

#### `w:h` — HEIGHT (Hauteur)
Signification: Hauteur de page ou cellule.
Unité: twips ou DXAs
Exemple: 
```xml
<w:pgSz w:h="16840"/>    <!-- Hauteur page A4 -->
<w:trHeight w:val="400"/>
```

#### `w:orient` — ORIENTATION (Orientation Page)
Signification: Orientation du papier.
Valeurs possibles:
- `"portrait"` — Vertical (standard)
- `"landscape"` — Horizontal
Exemple: 
```xml
<w:pgSz w:orient="portrait"/>
<w:pgSz w:orient="landscape"/>
```

#### `w:type` (dans pgSz/pgBorders/etc.) — TYPE
Signification: Dépend du contexte.
- **Pour breaks**: "page", "column", "continuous"
- **Pour bordures**: "nil", "single", "double", "wave", "dotted"
- **Pour headers**: "default", "first", "even"
Exemple: 
```xml
<w:pgBorders w:type="triple"/>
<w:headerReference w:type="default"/>
<w:headerReference w:type="first"/>
```

#### `w:start` — START (Démarrage Numérotation)
Signification: Numéro de démarrage pour la numérotation des pages.
Exemple: 
```xml
<w:pgNumType w:start="1"/>
<w:pgNumType w:start="5"/>    <!-- Commencer à 5 -->
```

#### `w:header`, `w:footer` — HEADER/FOOTER MARGIN
Signification: Distance entre le bord de la page et l'en-tête/pied de page.
Unité: twips
Exemple: 
```xml
<w:pgMar w:header="720" w:footer="720"/>  <!-- 0.5 inches -->
```

---

### Attributs de Contrôle Conditionnels

#### `w:val` (dans contrôles) — VALUE (0 ou 1)
Signification: Active/désactive une propriété de contrôle.
Valeurs:
- `"0"` ou absent → Désactivé
- `"1"` ou `"true"` → Activé
Éléments courants:
- `<w:keepNext w:val="1"/>` — Garder ensemble
- `<w:widowControl w:val="1"/>` — Éviter les veuves
- `<w:cantSplit w:val="1"/>` — Ne pas diviser
Exemple: 
```xml
<w:keepNext w:val="1"/>
<w:widowControl w:val="1"/>
<w:pageBreakBefore w:val="1"/>
```

---

### Attributs de Révision et Métadonnées

#### `w:rsidR` — REVISION ID RECEIVED
Signification: ID de révision du document (ajout). Utilisé pour le tracking des changements.
Format: Hexadécimal (8 caractères)
Exemple: 
```xml
<w:p w:rsidR="00000000" w:rsidDel="00000000">
```

#### `w:rsidDel` — REVISION ID DELETE
Signification: ID de révision pour suppression.
Format: Hexadécimal
Exemple: 
```xml
w:rsidDel="00000000"
```

#### `w:rsidP` — REVISION ID PARAGRAPH
Signification: ID de révision du paragraphe lui-même.
Format: Hexadécimal
Exemple: 
```xml
w:rsidP="00000000"
```

#### `w:rsidRDefault` — REVISION ID RUN DEFAULT
Signification: ID de révision par défaut pour les runs.
Format: Hexadécimal
Exemple: 
```xml
w:rsidRDefault="00000000"
```

#### `w:rsidRPr` — REVISION ID RUN PROPERTIES
Signification: ID de révision pour propriétés de run.
Format: Hexadécimal
Exemple: 
```xml
w:rsidRPr="00000000"
```

---

### Attributs de Bookmarks et Références

#### `w:colFirst` — COLUMN FIRST
Signification: Index de la première colonne du bookmark.
Exemple: 
```xml
<w:bookmarkStart w:colFirst="0" w:colLast="0"/>
```

#### `w:colLast` — COLUMN LAST
Signification: Index de la dernière colonne du bookmark.
Exemple: 
```xml
<w:bookmarkStart w:colFirst="0" w:colLast="0"/>
```

---

### Attributs de Tableaux

#### `w:gridSpan` — GRID SPAN (Fusion Colonnes)
Signification: Nombre de colonnes fusionnées pour cette cellule.
Exemple: 
```xml
<w:gridSpan w:val="2"/>    <!-- Cette cellule s'étend sur 2 colonnes -->
<w:gridSpan w:val="3"/>
```

#### `w:tblHeader` (attribut) — TABLE HEADER
Signification: Marque une ligne comme en-tête de tableau.
Exemple: 
```xml
<w:tr w:tblHeader="1">
```

#### `w:vAlign` — VERTICAL ALIGN (Alignement Vertical)
Signification: Alignement vertical du contenu (notamment dans cellules).
Valeurs possibles:
- `"top"` — En haut
- `"center"` — Centré
- `"bottom"` — En bas
Exemple: 
```xml
<w:vAlign w:val="top"/>
<w:vAlign w:val="center"/>
```

#### `w:tblLayout` — TABLE LAYOUT
Signification: Type de layout du tableau.
Valeurs possibles:
- `"autofit"` — Ajustement automatique
- `"fixed"` — Largeur fixe
Exemple: 
```xml
<w:tblLayout w:type="autofit"/>
```

---

### Attributs Divers

#### `w:outline` — OUTLINE LEVEL
Signification: Niveau de plan/outline du paragraphe (pour table des matières).
Exemple: 
```xml
<w:pPr><w:outlineLevel w:val="0"/></w:pPr>
```

#### `w:hRule` — HORIZONTAL RULE
Signification: Type de ligne horizontale/règle.
Exemple: 
```xml
w:hRule="clearLeft"
```

#### `w14:paraId` — PARAGRAPH ID (Office 2010+)
Signification: Identifiant unique du paragraphe (Office 2010 et versions ultérieures).
Format: Hexadécimal 8 caractères
Exemple: 
```xml
<w:p w14:paraId="00000001">
```

#### `xml:space` — XML SPACE (Gestion des Espaces XML)
Signification: Indique si les espaces blancs doivent être préservés.
Valeurs:
- `"default"` — Traitement standard
- `"preserve"` — Préserver exactement les espaces
Exemple: 
```xml
<w:t xml:space="preserve">TEXTE  AVEC   ESPACES</w:t>
```

---

| Attribut | Catégorie | Type | Exemple |
|----------|-----------|------|---------|
| `w:val` | Valeur générique | string/number | `w:val="center"` |
| `w:id` | ID/Identité | number | `w:id="3"` |
| `w:before` | Spacing | twips | `w:before="240"` |
| `w:after` | Spacing | twips | `w:after="0"` |
| `w:line` | Spacing | twips | `w:line="240"` |
| `w:sz` | Formatage | demi-points | `w:sz="40"` |
| `w:color` | Formatage | hex RGB | `w:color="538cd3"` |
| `w:left` | Indentation | twips | `w:left="1080"` |
| `w:hanging` | Indentation | twips | `w:hanging="360"` |
| `w:ascii`, `w:hAnsi` | Police | string | `w:ascii="Arial"` |
| `w:orient` | Page | portrait/landscape | `w:orient="portrait"` |
| `w:rsidR`, etc. | Révision | hex | `w:rsidR="00000000"` |
| `w:gridSpan` | Tableau | number | `w:gridSpan="2"` |
| `xml:space` | XML | preserve | `xml:space="preserve"` |

---

## 📏 Unités Communes

- **twips** (twentieth of a point) = 1/20 point
  - 1 inch = 1440 twips
  - 1 cm ≈ 567 twips
  
- **demi-points** (pour sizes): 1 pt = 2 demi-points
  - 12pt = 24 (demi-points)
  - 20pt = 40 (demi-points)

- **hex RGB** (pour couleurs): format RRGGBB
  - `000000` = noir
  - `FFFFFF` = blanc
  - `FF0000` = rouge
  - `538CD3` = bleu
  - `EC7C2F` = orange

---

## 🎯 Exemple Complet: Heading1 en XML

```xml
<w:p w:rsidR="00000000" w14:paraId="00000006">
  <!-- Propriétés du paragraphe -->
  <w:pPr>
    <w:spacing w:before="1"/>                    <!--  Espace avant minimal -->
    <w:ind w:left="1602" w:right="1476"/>       <!--  Indentation -->
    <w:jc w:val="center"/>                      <!--  ALIGNEMENT: CENTER -->
    <w:rPr>
      <w:rFonts w:ascii="Arial"/>               <!--  Police: Arial -->
      <w:b w:val="1"/>                          <!--  GRAS: OUI -->
      <w:bCs w:val="1"/>
      <w:sz w:val="40"/>                        <!--  TAILLE: 40 demi-pts = 20pt -->
      <w:szCs w:val="40"/>
    </w:rPr>
  </w:pPr>
  
  <!-- Le texte avec son formatage -->
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Arial"/>
      <w:b w:val="1"/>
      <w:bCs w:val="1"/>
      <w:color w:val="538cd3"/>                 <!--  COULEUR: BLEU -->
      <w:sz w:val="40"/>
      <w:szCs w:val="40"/>
      <w:rtl w:val="0"/>                        <!--  Direction: gauche-à-droite -->
    </w:rPr>
    <w:t xml:space="preserve">DOSSIER DE COMPETENCES</w:t>
  </w:r>
</w:p>
```

**Traduction**: Paragraphe centré, gras, 20pt, bleu, contenant "DOSSIER DE COMPETENCES"

---

**Note**: Ce dictionnaire couvre la plupart des balises trouvées dans DC_JNZ_2026_RAW.xml. Certaines balises avancées (comme celles concernant les formes, dessins complexes, ou namespaces étendus) ne sont pas couvertes ici.
