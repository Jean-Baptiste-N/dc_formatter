# Guide : Template & Styles pour le Reformattage DC

## 🎯 Vue d'ensemble

Au lieu de laisser `parse_reformat.py` supprimer TOUS les headings (comme actuellement), on peut créer un **template DOCX avec des styles prédéfinis** qui serviront de base de reformattage cohérente.

### Files créés :
1. **`TEMPLATE_DC.docx`** (généré) - Template avec 9 styles prédéfinis
2. **`tools/create_dc_template.py`** - Script de création du template
3. **`tools/apply_dc_styles.py`** - Utilitaire d'application des styles

---

## 📋 Styles Disponibles

| Style | Usage | Formatage |
|-------|-------|-----------|
| **DC_Header** | Titre principal (DOSSIER DE COMPÉTENCES) | Arial 20pt, Gras, Bleu, Centré |
| **DC_Name** | Nom de la personne | Arial 20pt, Gras, Orange, Centré |
| **DC_Title** | Titre de poste (Data Engineer, etc.) | Arial 16pt, Gras, Centré |
| **DC_Section** | Grandes sections (Domaines, Formations) | Arial 18pt, Gras, Bleu foncé |
| **DC_Subsection** | Sous-sections (Gestion, Traitement données) | Arial 14pt, Gras, Bleu, Indenté |
| **DC_Bullet** | Points principaux | Arial 11pt, Indentés |
| **DC_SubBullet** | Sous-points (niveau 2+) | Arial 10pt, Doublement indentés |
| **DC_TableHeader** | En-têtes de table | Arial 11pt, Gras, Blanc sur bleu |
| **Normal (modifié)** | Texte standard | Arial 11pt |

---

## 🚀 Utilisation Rapide

### 1️⃣ Générer le template (une fois)
```bash
python tools/create_dc_template.py
# Crée: TEMPLATE_DC.docx
```

### 2️⃣ Appliquer les styles à un document
```bash
python tools/apply_dc_styles.py test/original.docx output_reformatted.docx
```

### 3️⃣ Ouvrir dans Word
Le document aura automatiquement les styles DC appliqués ! ✨

---

## 🔧 Intégration dans `parse_reformat.py`

### Approche 1 : Utiliser le template comme base

Modifier la fonction `write_reformat_dc()` pour :

```python
from tools.apply_dc_styles import DCStyleApplier

def write_reformat_dc(
    reformat_data: dict,
    output_path: str,
    template_path: str = './TEMPLATE_DC.docx'
) -> None:
    """Écrit le DC reformatté avec styles prédéfinis."""
    
    # Créer le document à partir du template
    doc = Document(template_path)
    
    # Appliquer chaque bloc avec le style approprié
    for block_name, block_data in reformat_data.items():
        
        if block_name == 'h_dc':
            para = doc.add_paragraph(block_data['texte'], style='DC_Header')
        
        elif block_name == 'trigram':
            para = doc.add_paragraph(block_data['texte'], style='DC_Name')
        
        elif block_name in ['h_main_skills', 'h_experiences', 'h_formations', 'h_langues']:
            # Grandes sections
            para = doc.add_paragraph(block_data['texte'], style='DC_Section')
        
        elif block_name.startswith('lst_'):
            # Points de liste
            for item in block_data:
                para = doc.add_paragraph(item, style='DC_Bullet')
        
        # ... et ainsi de suite pour tous les blocs
    
    doc.save(output_path)
    print(f"✓ DC reformatté avec styles: {output_path}")
```

### Approche 2 : Appliquer automatiquement après reformattage

```python
def write_reformat_dc(...) -> None:
    # ... code de reformattage existant ...
    doc.save(output_path)
    
    # Appliquer les styles DC
    applier = DCStyleApplier('./TEMPLATE_DC.docx')
    doc = Document(output_path)
    doc = applier.detect_and_apply_styles(doc)
    doc.save(output_path)
```

---

## 💡 Avantages

✅ **Structure préservée** : Plus de Heading1/Heading2 perdus  
✅ **Cohérence visuelle** : Tous les DC reformattés ont la même apparence  
✅ **Facilité d'évolution** : Changer un style = changer tous les DC qui l'utilisent  
✅ **Hiérarchie maintemue** : Les sections sont logiquement structurées  
✅ **Accessible** : Lecteurs d'écran peuvent naviguer correctement  
✅ **Table of Contents fonctionnelle** : Word peut générer une TOC automatique  

---

## 🔍 Comparaison : Avant vs Après

### ❌ Avant (actuel)
```
- 3 sections → 1 section (perte)
- 25 Headings → 0 (100% perte)
- Structure perdue
- Tous les paragraphes en "Normal"
```

### ✅ Après (avec template)
```
- Structure préservée avec styles DC
- Headings remplacés par DC_Section et DC_Subsection
- Hiérarchie claire et maintenable
- Styles spécifiques pour chaque bloc
```

---

## 📝 Prochaines Étapes

1. **Tester la détection automatique** : Vérifier si `detect_and_apply_styles()` fonctionne bien
2. **Intégrer dans `parse_reformat.py`** : Ajouter l'appel à `DCStyleApplier`
3. **Tester sur plusieurs DC** : S'assurer que la détection fonctionne sur différents formats
4. **Affiner les styles** : Ajuster les couleurs, tailles selon les préférences
5. **Permettre la customization** : Passer des paramètres pour modifier les styles à la volée

---

## 🛠️ Possibilités Avancées

### Créer des styles personnalisés pour votre entreprise

```python
def create_custom_template(company_name: str, colors: dict) -> None:
    """Crée un template avec les couleurs de l'entreprise."""
    doc = Document()
    
    # Créer DC_Section avec couleur personnalisée
    style = doc.styles.add_style('DC_Section', WD_STYLE_TYPE.PARAGRAPH)
    style.font.color.rgb = RGBColor(*colors['primary'])  # Couleur primaire
    
    doc.save(f'TEMPLATE_{company_name}.docx')
```

### Appliquer les styles via mapping personnalisé

```python
applier = DCStyleApplier('TEMPLATE_DC.docx')
applier.STYLE_MAPPING['competence'] = 'DC_Subsection'  # Custom mapping
doc = applier.detect_and_apply_styles(doc)
```

---

## 📚 Ressources

- [python-docx Documentation](https://python-docx.readthedocs.io/)
- [Office Open XML Styles](https://docs.microsoft.com/en-us/office/open-xml/working-with-paragraphs)
- Style Types : `Normal`, `Heading 1-9`, `List Bullet`, etc.

