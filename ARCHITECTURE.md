# 🔄 Architecture et Flux de Données

## Architecture Générale

```
┌─────────────────────────────────────────────────────────────────────┐
│                           PIPELINE                                   │
├─────────────────────────────────────────────────────────────────────┤
│                                                                       │
│  ┌─────────────────────────────────────────────────────────────┐    │
│  │ PHASE 1: EXTRACTION                                         │    │
│  ├─────────────────────────────────────────────────────────────┤    │
│  │                                                               │    │
│  │  Step 1: Extract Template Dimensions (ONCE)                 │    │
│  │  ─────────────────────────────────────                      │    │
│  │  template.docx ──> extract_page_dimensions ──> page_dims   │    │
│  │                                                               │    │
│  │  Step 2: Extract XML                                        │    │
│  │  ───────────────────                                        │    │
│  │  document.docx ──> export_all_xml ──> GLOBAL.xml          │    │
│  │                                                               │    │
│  │  Step 3: Convert XML to JSON RAW                            │    │
│  │  ──────────────────────────────────                         │    │
│  │  GLOBAL.xml ──> xml_to_json ──> GLOBAL_raw.json           │    │
│  │                                                               │    │
│  └─────────────────────────────────────────────────────────────┘    │
│                                                                       │
│  ┌─────────────────────────────────────────────────────────────┐    │
│  │ PHASE 2: TRANSFORMATION & RENDERING                         │    │
│  ├─────────────────────────────────────────────────────────────┤    │
│  │                                                               │    │
│  │  Step 4: Transform JSON RAW (with page_dims)                │    │
│  │  ──────────────────────────────────────────                 │    │
│  │  GLOBAL_raw.json + page_dims ──> apply_tags_and_styles     │    │
│  │                          ──> GLOBAL_transformed.json        │    │
│  │                                                               │    │
│  │  Step 5: Render to DOCX                                     │    │
│  │  ──────────────────────                                     │    │
│  │  GLOBAL_transformed.json ──> json_to_docx ──> output.docx  │    │
│  │                                                               │    │
│  └─────────────────────────────────────────────────────────────┘    │
│                                                                       │
└─────────────────────────────────────────────────────────────────────┘
```

## 📊 Flux de Données

### Phase 1: Extraction (Dims + XML + JSON RAW)

```
Template DOCX
    │
    ├──> extract_page_dimensions_from_template()
    │        │
    │        └──> page_dimensions = {
    │               'page_width': 11906,
    │               'page_height': 16838,
    │               'usable_width': 9638,
    │               'col_fixed_width_3': 1701,
    │               'col_fixed_width_5': 2835,
    │               ...
    │            }

Document DOCX
    │
    ├──> export_all_xml()
    │        │
    │        └──> structures/DOCX_GLOBAL.xml
    │
    ├──> xml_to_json()
    │        │
    │        └──> structures/DOCX_GLOBAL_raw.json
    │            {
    │              "document": {
    │                "content": [
    │                  { "type": "Paragraph", "text": "..." },
    │                  { "type": "Table", "rows": [...] },
    │                  ...
    │                ]
    │              }
    │            }
```

### Phase 2: Transformation & Rendering

```
GLOBAL_raw.json + page_dimensions
    │
    ├──> apply_tags_and_styles()
    │        │
    │        ├─ Add 'page_dimensions' to data
    │        │   data['page_dimensions'] = page_dimensions
    │        │
    │        ├─ Detect sections
    │        │   (header, main_skills, education, professional_experience)
    │        │
    │        ├─ Apply tags to elements
    │        │   para['section'] = 'education'
    │        │
    │        ├─ Create/transform tables
    │        │   Uses page_dimensions from data
    │        │   Gets page_dims = data['page_dimensions']
    │        │
    │        ├─ Apply styles
    │        │   paragraph styles, run properties, etc.
    │        │
    │        └──> structures/DOCX_GLOBAL_transformed.json
    │            {
    │              "document": {
    │                "content": [
    │                  { 
    │                    "type": "Paragraph",
    │                    "section": "education",
    │                    "properties": { "style": "..." },
    │                    "runs": [...]
    │                  },
    │                  { 
    │                    "type": "Table",
    │                    "section": "education",
    │                    "rows": [...]
    │                  },
    │                  ...
    │                ],
    │                "page_dimensions": { ... }
    │              }
    │            }
    │
    ├──> json_to_docx()
    │        │
    │        ├─ Load template DOCX
    │        │
    │        ├─ For each element in JSON:
    │        │   ├─ If Paragraph: add_paragraph_from_json()
    │        │   └─ If Table: add_table_from_json()
    │        │
    │        └──> renders/DOCX_GLOBAL_generated.docx
```

## 🔗 Connexions Entre Modules

### extract_xml_raw.py
```
Public:
  ✓ export_all_xml(docx_file, output_folder) -> str
  ✓ extract_document_xml(docx_file, output_folder) -> str

Private:
  - indent_xml_string()
  - extract_xml_raw()
  - create_global_xml()
```

### parse_template.py
```
Public:
  ✓ extract_page_dimensions_from_template(template_path) -> dict

Private:
  - (none, very focused)
```

### parse_xml_raw_to_json_raw.py
```
Public:
  ✓ xml_to_json(xml_file, output_file) -> str

Private:
  - parse_global_xml()
  - parse_paragraph()
  - parse_table()
  - extract_run_properties()
  - extract_paragraph_properties()
  - extract_runs_from_paragraph()
  - normalize_paragraph_runs()
  - extract_table_properties()
  - extract_cell_width()
  - extract_row_height()
```

### process_json_raw_to_json_transformed.py
```
Public:
  ✓ apply_tags_and_styles(raw_json_file, output_dir, page_dimensions) -> str

Imports:
  - extract_page_dimensions_from_template (REMOVED: no longer called internally)
  - xml_to_json (from parse_xml_raw_to_json_raw)

Private:
  - get_text_from_element()
  - detect_section_by_keyword()
  - apply_section_tags()
  - create_empty_table_2x2()
  - create_edu_table()
  - create_xp_tables()
  - apply_styles_in_json()
  - ... (many more helper functions)
```

### render_json_transformed_to_docx.py
```
Public:
  ✓ json_to_docx(json_file, template_file, output_dir) -> str

Private:
  - parse_alignment()
  - add_paragraph_from_json()
  - add_table_from_json()
  - set_table_borders()
```

### pipeline.py
```
Public:
  ✓ main() - CLI entry point

Imports ALL modules:
  - extract_xml_raw.export_all_xml
  - parse_template.extract_page_dimensions_from_template
  - parse_xml_raw_to_json_raw.xml_to_json
  - process_json_raw_to_json_transformed.apply_tags_and_styles
  - render_json_transformed_to_docx.json_to_docx

Commands:
  ✓ extract-dims
  ✓ extract-xml
  ✓ xml-to-json
  ✓ transform
  ✓ render
  ✓ extract (combined)
  ✓ transform-render (combined)
  ✓ full (complete pipeline)
```

## 📋 Data Structures

### page_dimensions dict
```python
{
    'page_width': 11906,           # twips
    'page_height': 16838,          # twips
    'usable_width': 9638,          # page_width - left - right margins
    'top_margin': 1134,            # twips
    'bottom_margin': 1134,         # twips
    'left_margin': 1134,           # twips
    'right_margin': 1134,          # twips
    'col_fixed_width_3': 1701,     # 3 cm in twips
    'col_fixed_width_5': 2835,     # 5 cm in twips
}
```

### JSON RAW Structure
```python
{
    "document": {
        "type": "Document",
        "source": "DC_JNZ_2026.docx",
        "content": [
            {
                "type": "Paragraph",
                "index": 0,
                "properties": {
                    "style": "Normal",
                    "alignment": "left"
                },
                "runs": [
                    {
                        "text": "Dossier de compétences",
                        "properties": {
                            "bold": true,
                            "size": 28
                        }
                    }
                ]
            },
            {
                "type": "Table",
                "index": 1,
                "row_count": 3,
                "col_count": 2,
                "rows": [...]
            }
        ],
        "stats": {
            "total_elements": 150,
            "paragraphs": 120,
            "tables": 5
        }
    }
}
```

### JSON TRANSFORMED Structure (Addition)
```python
{
    "document": {
        "type": "Document",
        "source": "DC_JNZ_2026.docx",
        "content": [
            {
                "type": "Paragraph",
                "index": 0,
                "section": "header",  # NEW: section tag
                "properties": { ... },
                "runs": [ ... ]
            },
            # ... rest similar to RAW
        ],
        "page_dimensions": { ... }  # NEW: stored for reference
    }
}
```

## ⚙️ Key Design Decisions

1. **One-Time Extraction of Dimensions**
   - Extracted at the very beginning
   - Passed as `page_dimensions` parameter to `apply_tags_and_styles`
   - Stored in `data['page_dimensions']` for internal access
   - Eliminates repeated template file reading

2. **Mandatory Parameters**
   - All parameters are required (no defaults)
   - Makes code paths explicit and testable
   - Prevents silent failures or unexpected behavior

3. **Two-Phase Architecture**
   - Phase 1: Extraction (can be done once for many documents)
   - Phase 2: Transformation (can reuse Phase 1 outputs)
   - Supports batch processing efficiently

4. **Clear CLI Structure**
   - Individual commands for each step
   - Combined commands for common workflows
   - Consistent error handling and reporting

## 📈 Performance Implications

- **Phase 1 (Extraction)**: ~1-2 seconds per document
  - XML extraction is I/O bound
  - JSON RAW creation requires parsing

- **Phase 2 (Transformation)**: ~2-5 seconds per document
  - Tag detection requires full content scan
  - Table creation/modification is CPU bound
  - Style application is regex heavy

- **Batch Processing**: Reusing Phase 1 outputs saves 50% time for 3+ documents
  - Extract once: 2 seconds
  - Transform 3 documents: 3×4 seconds = 12 seconds
  - Total: 14 seconds vs 3×(2+4) = 18 seconds
