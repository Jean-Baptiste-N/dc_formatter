# Documentation du Reformatage de Documents (DC)

## Vue d'ensemble

Le script `parse_reformat.py` reformate un document Word (.docx) en appliquant une structure standardisée basée sur des sections prédéfinies avec styles spécifiques. Le processus analysé statiquement le contenu pour identifier les sections (titre, compétences, formations, expériences) et réapplique des styles cohérents.

## Flux d'exécution global

```
Utilisateur exécute:
  python -m tools.parse_reformat input.docx
        ↓
  main(src_path, dst_path)
        ↓
  parse_and_reformat_dc(src_dc, dst_path)
        ↓
  [Boucle sur BLOCK_DEFINITIONS]
        ↓
  output_reformated.docx
```

## Détail du processus

### 1. **Chargement des documents**
```python
def main(src_path, dst_path):
    src_dc = docx.Document(src_path)  # Charge le DC source
    parse_and_reformat_dc(src_dc, dst_path)  # Lance le reformatage
```

### 2. **Initialisation**
```python
def parse_and_reformat_dc(src_dc, dst_path):
    dst_dc = docx.Document(TEMPLATE_DC_PATH)  # Charge le template (assets/template.docx)
    it_src_dc = tee(src_dc.iter_inner_content())  # Crée un itérateur sur les éléments
```

Le template définit :
- Les styles de page et numérotation
- Le logo Epsyl
- Les listes de numérotation

### 3. **Boucle de traitement par blocs**

Le script itère sobre les **bloc_definitions** (18 sections prédéfinies) dans cet ordre :

#### **Page 1 - Entête du CV**
1. `h_dc` → Titre principal (centré, bleu, 20pt)
2. `trigram` → Trigramme (centré, orange, 20pt)
3. `role` → Poste/fonction (centré, 20pt)
4. `years_of_experience` → Années d'expérience (centré, 20pt)

#### **Page 2 - Compétences**
5. `h_main_skills` → Titre "Compétences" ou "Compétence"
6. `lst_main_skills` → Liste/table de compétences (en gras orange pour titre, justifié)
7. `tbl_main_skills` → Tableau optionnel de compétences

#### **Page 3+ - Formations**
8. `h_education_1` → Titre "Formation/Diplôme/Certification/Langue"
9. `tbl_education_1` → Tableau de formation
10. `h_education_2` → 2e titre formation (optionnel)
11. `tbl_education_2` → 2e tableau (optionnel)
12. `h_education_3` → 3e titre formation (optionnel)
13. `tbl_education_3` → 3e tableau (optionnel)

#### **Pages suivantes - Expériences**
14. `h_experiences` → Titre "Expérience"
15. `tbl_company_header` → Tableau avec nom/dates (nom en Géorgia 20pt, dates en gris, bordure bas)
16. `project_summary` → Résumé du projet en paragraphes justifiés
17. `lst_project_details` → Liste des détails (en gras bleu)
18. `technical_environment` → Environnement technique (en gras orange, indentation 0.5")

**Après le bloc 18**, la boucle recommence au bloc 14 (`tbl_company_header`) pour les expériences suivantes.

### 4. **Pour chaque bloc : Critère de validation**

À chaque itération, le script :

```python
criterion_validator, kwargs = block_definition["criterion"]
expected, unexpected = criterion_validator(it_src_dc, **kwargs)
```

Les **validateurs de critères** (`cv_*`) examinent les éléments du document source et séparent le contenu **attendu** de l'**inattendu** :

#### **Validateurs disponibles** :

| Validateur | Action |
|---|---|
| `cv_paragraphs_until_empty()` | Retourne tous les paragraphes jusqu'au premier vide |
| `cv_text_match()` | Retourne le prochain paragraphe s'il contient l'un des `candidates` (case-insensitive) |
| `cv_table()` | Retourne le prochain élément s'il s'agit d'un tableau |
| `cv_successive_list_elements()` | Retourne les paragraphes consécutifs avec indentation de liste (`ilvl`) |
| `cv_paragraphs_before_list()` | Retourne les paragraphes avant le début d'une liste |

Chaque validateur peut avoir un paramètre `optional=True` : si le contenu n'est pas trouvé, il ne lève pas d'erreur.

### 5. **Écriture des blocs attendus**

```python
write_block(expected_block, dst_dc, styles)
```

Applique les styles définis pour ce bloc :

- **Paragraphes** : Police, couleur, gras, taille, indentation, espacement
- **Listes** : Numérotation avec niveau d'indentation (`ilvl`)
- **Tableaux** : Largeur colonnes, marges, bordures

#### Exemple - `lst_main_skills` :
```python
"styles": [
    {  # Style pour le titre de la liste
        "paragraph_format": {"alignment": JUSTIFY, "space_before": 0.1"},
        "font": {"bold": True, "color": ORANGE},
        "is_list": True
    },
    {  # Style pour les éléments de la liste
        "paragraph_format": {"alignment": JUSTIFY},
        "preserve_src_bold": True  # Garde le gras du source
    }
]
```

### 6. **Écriture des blocs inattendus**

```python
write_unexpected_block(unexpected_block, dst_dc)
```

Écrit en **JAUNE** (surligné) pour signaler un contenu non identifié. Cela aide à déboguer les structures non conformes.

### 7. **Sauvegarde**

```python
dst_dc.save(dst_path)
```

Sauvegarde le document reformaté (ex: `input.docx` → `input_reformated.docx`)

---

## Fichiers utilitaires

### `utils.py` - Fonctions utilitaires

| Fonction | Action |
|---|---|
| `get_ilvl(p)` | Retourne le niveau d'indentation de liste d'un paragraphe |
| `get_font_props(run, parent)` | Extrait les propriétés de police (taille, couleur, gras, italique, souligné) en respectant la hiérarchie des styles |
| `get_format_props(p)` | Extrait les propriétés de format de paragraphe |
| `rec_add_xml_children(elm, children)` | Construit récursivement une arborescence XML (utilisée pour les bordures, marges de tableau) |

### `print_dc.py` - Débogage

Script permettant d'afficher la structure complète d'un DC avec tous les attributs des éléments (sections, styles, paragraphes, runs, polices, etc.). Utilisé pour comprendre la structure source avant reformatage.

### `write.py` - Écriture simplifiée

Script alternatif de réécriture plus simple : fusionne les runs consécutifs avec les mêmes propriétés et nettoie les espacements. Utilisé à titre d'expérimentation.

---

## Comment le DOCX est ingéré

1. **Ouverture** : `docx.Document(chemin)` charge le fichier (classe Document de python-docx)
2. **Itération** : Iteratre sur `document.iter_inner_content()` → fournit paragraphes et tableaux dans l'ordre
3. **Inspection** :
   - Paragraphes : `.text`, `.runs`, `.paragraph_format`, `.alignment`
   - Tableaux : `.rows`, `.columns`, `.cells`
   - Runs : `.font`, `.text`
   - Propriétés de style : hiérarchie run → parent.style → defaults
4. **Transformation** : Éléments lus transformés et réécrits dans le template
5. **Export** : `document.save()` génère un nouveau `.docx`

---

## Points clés

- ⚠️ **Ordre des blocs** : défini et séquentiel (boucle cyclique pour expériences)
- 🔍 **Validation** : chaque bloc cherche un critère spécifique (texte, tableau, liste)
- 🎨 **Styles** : prédéfinis et appliqués systématiquement
- 🟨 **Débogage** : contenu inattendu surligné en jaune
- 📄 **Template** : base neutre chargée et enrichie progressivement

