# 📊 Patterns XML Découverts dans DC_JNZ_2026

## 🎯 Résumé Exécutif

Le fichier `DC_JNZ_2026_RAW.xml` (12,622 lignes) révèle trois patterns distincts correspondant aux trois niveaux de hiérarchie:

---

## 1️⃣ **HEADING1** (Titres Principaux)

### Exemple: "DOSSIER DE COMPETENCES"
```xml
<w:p>
  <w:pPr>
    <w:ind w:left="1602" w:right="1476" w:firstLine="0"/>
    <w:jc w:val="center"/>                    ✅ CENTER
    <w:rPr>
      <w:b w:val="1"/>                       ✅ BOLD
      <w:sz w:val="40"/>                     ✅ SIZE 40 (20pt)
      <w:color w:val="538cd3"/>              ✅ BLUE (538cd3)
    </w:rPr>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:color w:val="538cd3"/>
      <w:sz w:val="40"/>
      <w:rtl w:val="0"/>
    </w:rPr>
    <w:t>DOSSIER DE COMPETENCES</w:t>
  </w:r>
</w:p>
```

### Signature XML de Heading1:
```
✓ w:pPr/w:jc w:val="center"      (CENTER ALIGNMENT)
✓ w:pPr/w:rPr/w:b w:val="1"      (BOLD)
✓ w:pPr/w:rPr/w:sz w:val="40"    (SIZE 40 = 20pt)
✓ w:r/w:rPr/w:color = "538cd3" ou "f69545" (BLUE ou ORANGE)
✗ PAS de <w:pStyle w:val="Heading1"/> appliqué
```

### Autres Exemples Observés:
- **"JNZ"** - Center, Bold, size=40, color=f69545 (ORANGE)
- **"Expérience (I2)"** - Center, Bold, size=40, color=f69545 (ORANGE)

---

## 2️⃣ **HEADING2** (Sous-titres / Catégories)

### Exemple: "Domaines de compétences"
```xml
<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading1"/>            ⚠️ MARQUÉ COMME HEADING1 DANS SOURCE
    <w:spacing w:before="77"/>
    <w:ind w:firstLine="424"/>
    <w:rPr>
      <w:u w:val="none"/>
    </w:rPr>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:color w:val="538cd4"/>            ✅ BLUE (538cd4)
      <w:u w:val="none"/>
    </w:rPr>
    <w:t>Domaines de compétences</w:t>
  </w:r>
</w:p>
```

### Sous-Heading2 Numéroté: "Gestion de projets"
```xml
<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading2"/>            ⚠️ MARQUÉ COMME HEADING2 DANS SOURCE
    <w:numPr>
      <w:ilvl w:val="0"/>                  ✅ ILVL=0
      <w:numId w:val="3"/>
    </w:numPr>
    <w:spacing w:before="288"/>
    <w:ind w:left="1080" w:hanging="360"/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:color w:val="ec7c2f"/>            ✅ ORANGE (ec7c2f)
    </w:rPr>
    <w:t>Gestion de projets</w:t>
  </w:r>
</w:p>
```

### Signature XML des Heading2:
```
Option A (Simples):
✓ w:pStyle w:val="Heading1"     (SOURCE a Heading1 appliqué)
✓ w:r/w:rPr/w:color = "538cd4"  (BLUE)
✗ Pas de bold/size/center

Option B (Numérotés):
✓ w:pStyle w:val="Heading2"     (SOURCE a Heading2 appliqué)
✓ w:numPr/w:ilvl w:val="0"      (ILVL=0)
✓ w:r/w:rPr/w:color = "ec7c2f"  (ORANGE)
```

---

## 3️⃣ **NORMAL** (Contenu / Bullets)

### Exemple: "Recueil des besoins"
```xml
<w:p>
  <w:pPr>
    <w:numPr>
      <w:ilvl w:val="1"/>                  ✅ ILVL=1 (NESTED)
      <w:numId w:val="4"/>
    </w:numPr>
    <w:ind w:left="2210" w:right="0" w:hanging="360"/>
    <w:jc w:val="left"/>
    <w:rPr>
      <w:b w:val="0"/>                     ✅ NOT BOLD
      <w:sz w:val="20"/>                   ✅ SIZE 20 (10pt)
      <w:color w:val="000000"/>            ✅ BLACK
    </w:rPr>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:b w:val="0"/>
      <w:sz w:val="20"/>
      <w:color w:val="000000"/>
    </w:rPr>
    <w:t>Recueil des besoins</w:t>
  </w:r>
</w:p>
```

### Signature XML du Normal:
```
✓ w:numPr/w:ilvl w:val="1" ou plus  (NESTED LIST)
✓ w:jc w:val="left"                 (LEFT ALIGN)
✓ w:rPr/w:b w:val="0"               (NOT BOLD)
✓ w:rPr/w:sz w:val="20"             (SIZE 20 = 10pt)
✓ w:rPr/w:color w:val="000000"      (BLACK)
```

---

## 🔍 Analyse Critique - DÉCOUVERTE CHOC

### Le mystère de la perte de styles:

Le XML **SOURCE** (DC_JNZ_2026) contient:
- ✅ `w:pStyle w:val="Heading1"` pour "Domaines de compétences"
- ✅ `w:pStyle w:val="Heading2"` pour "Gestion de projets"

Mais après reformatage:
- ❌ Tous ces styles disparaissent!
- ❌ Convertitur en Normal
- ❌ Les paragraphes CENTER+BOLD+SIZE40 deviennent orphelins

### Conclusion:

**`parse_reformat.py` supprime TOUS les styles et les remplace par Normal**, même quand le source a Heading1/Heading2 explicites!

### Solution d'Heuristique (CORRIGÉE):

Les Headings perdus peuvent être détectés par:

1. **Heading1** (VISUAL):
   ```
   center=true AND bold=true AND size≥40 AND color∈{538cd3,f69545,538cd4}
   ```
   → Exemples: "DOSSIER DE COMPETENCES", "JNZ", "Expérience (I2)"

2. **Heading2** (HYBRID):
   ```
   (color∈{538cd4,ec7c2f} AND short_text) OR (ilvl=0 AND numId>0)
   ```
   → Exemples: "Domaines de compétences", "Gestion de projets"

3. **Normal** (DEFAULT):
   ```
   Tout le reste: ilvl≥1, left, small font, black
   ```

---

## 📈 Distribution Observée

Sur les 35 premiers paragraphes analysés:
- **4 Heading1** (DOSSIER, JNZ, Expérience, Job Title)
- **2-3 Heading2** (Domaines, Skills, Traitement...)
- **25+ Normal** (Content bullets)

---

## ✅ Fichier Source XML Available

📂 **`DC_JNZ_2026_RAW.xml`** (12,622 lignes)
- Indentation complète
- Structure préservée
- Prêt pour exploration manuelle
- Tous les namespaces en place

Vous pouvez l'ouvrir dans VS Code et utiliser Ctrl+F pour chercher:
- `<w:t>` pour trouver les textes
- `<w:jc w:val="center"` pour les centrés
- `w:color w:val="` pour les couleurs
- `w:ilvl` pour les niveaux d'indentation
