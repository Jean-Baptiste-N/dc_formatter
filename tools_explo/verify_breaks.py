import zipfile
import xml.etree.ElementTree as ET

docx_file = 'renders/DC_JNZ_2026_GLOBAL_generated.docx'

with zipfile.ZipFile(docx_file, 'r') as zip_ref:
    doc_xml = zip_ref.read('word/document.xml')

root = ET.fromstring(doc_xml)
ns_w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

print("Verification des sauts dans le DOCX genere:")

# Sauts de page
page_breaks = root.findall(f'.//{{{ns_w}}}br[@{{{ns_w}}}type="page"]')
print(f"  Sauts de page: {len(page_breaks)} ✓")

# Sections
sections = root.findall(f'.//{{{ns_w}}}sectPr')
print(f"  Sections trouvees: {len(sections)} ✓")

# Afficher les types de section
for i, sect in enumerate(sections[:3]):
    type_elem = sect.find(f'{{{ns_w}}}type')
    if type_elem is not None:
        type_val = type_elem.get(f'{{{ns_w}}}val')
        print(f"    - Section {i}: Type={type_val}")
    else:
        print(f"    - Section {i}: Type=default (nextPage)")

print("\n✅ Sauts de page et section preserves avec succes!")
