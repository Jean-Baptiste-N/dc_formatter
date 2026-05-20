#!/usr/bin/env python3
"""Analyser le XML brut du .docx pour identifier les heuristiques"""

import zipfile
from xml.etree import ElementTree as ET

# Extraire et parser le XML
with zipfile.ZipFile("test/DC_JNZ_2026.docx", "r") as zip_ref:
    xml_content = zip_ref.read("word/document.xml")

root = ET.fromstring(xml_content)
ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

paragraphs = root.findall(".//w:p", ns)

print("=" * 100)
print("ANALYSE DES PARAGRAPHES XML BRUTS - DC_JNZ_2026")
print("=" * 100)

count = 0
for i, para in enumerate(paragraphs):
    text_nodes = para.findall(".//w:t", ns)
    text = "".join([t.text for t in text_nodes if t.text])
    
    if text and text.strip() and count < 35:
        count += 1
        
        print(f"\n╔════ PARA #{count} (XML index {i}) ════")
        print(f"║ Texte: {text[:70]}")
        
        pPr = para.find("w:pPr", ns)
        if pPr is not None:
            # Alignment
            jc = pPr.find("w:jc", ns)
            if jc is not None:
                align = jc.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
                print(f"║ 📐 Alignment: {align}")
            
            # List indent level
            numPr = pPr.find("w:numPr", ns)
            if numPr is not None:
                ilvl = numPr.find("w:ilvl", ns)
                if ilvl is not None:
                    ilvl_val = ilvl.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
                    print(f"║ 📊 ilvl: {ilvl_val}")
            
            # Paragraph-level rPr
            pRPr = pPr.find("w:rPr", ns)
            if pRPr is not None:
                bold = pRPr.find("w:b", ns)
                sz = pRPr.find("w:sz", ns)
                color = pRPr.find("w:color", ns)
                if bold is not None:
                    print(f"║ 🔤 Bold (pPr): ✓")
                if sz is not None:
                    sz_val = int(sz.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"))
                    print(f"║ 📏 Size (pPr): {sz_val} ({sz_val/2}pt)")
                if color is not None:
                    print(f"║ 🎨 Color (pPr): {color.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')}")
        
        # Run-level properties
        runs = para.findall("w:r", ns)
        for r_idx, run in enumerate(runs):
            t = run.find("w:t", ns)
            run_text = t.text[:40] if t is not None and t.text else ""
            
            rPr = run.find("w:rPr", ns)
            if rPr is not None or run_text:
                props = []
                if rPr is not None:
                    bold = rPr.find("w:b", ns)
                    sz = rPr.find("w:sz", ns)
                    color = rPr.find("w:color", ns)
                    
                    if bold is not None:
                        bold_val = bold.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "1")
                        if bold_val != "0":
                            props.append("BOLD")
                    if sz is not None:
                        sz_val = int(sz.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"))
                        props.append(f"sz={sz_val/2}pt")
                    if color is not None:
                        props.append(f"col={color.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')[:3]}")
                
                if props or run_text:
                    prop_str = f" | {', '.join(props)}" if props else ""
                    print(f"║   └─ Run[{r_idx}]: {run_text}{prop_str}")
        
        print(f"╚" + "=" * 96)

print(f"\n✅ Total paragraphes analysés: {count}")
