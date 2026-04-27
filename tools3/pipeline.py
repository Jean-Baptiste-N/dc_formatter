"""
Pipeline complet pour traiter un document DOCX:
1. Extraction du template, XML brut et JSON brut
2. Transformation en JSON avec tags/styles et rendu DOCX

Structure CLI:
  pipeline.py extract-dims TEMPLATE
  pipeline.py extract-xml XML [-o OUTPUT_DIR]
  pipeline.py xml-to-json JSON_RAW [-o OUTPUT_DIR]
  pipeline.py transform JSON [-o OUTPUT_DIR]
  pipeline.py render DOCX [-o OUTPUT_DIR]
  pipeline.py extract DOCX [-o OUTPUT_DIR]              # dims + xml + json raw
  pipeline.py transform-render DOCX [-o OUTPUT_DIR]     # json raw to docx
  pipeline.py full DOCX [-o OUTPUT_DIR]                 # pipeline complète

Dossiers de résultats (par défaut):
  - OUTPUT1_XML-RAW/        # XML global
  - OUTPUT2_JSON-RAW/       # JSON RAW
  - OUTPUT3_JSON-TRANSFORMED/ # JSON transformé
  - OUTPUT4_DOCX-RESULT/    # DOCX final
"""

import argparse
import sys
from pathlib import Path

from .extract_xml_raw import export_all_xml
from .parse_template import extract_page_dimensions_from_template
from .parse_xml_raw_to_json_raw import xml_to_json
from .process_json_raw_to_json_transformed import apply_tags_and_styles
from .render_json_transformed_to_docx import json_to_docx

# ===== CONSTANTES =====
TEMPLATE_PATH = 'assets/TEMPLATE.docx'
OUTPUT_XML_RAW = 'OUTPUT1_XML-RAW'
OUTPUT_JSON_RAW = 'OUTPUT2_JSON-RAW'
OUTPUT_JSON_TRANSFORMED = 'OUTPUT3_JSON-TRANSFORMED'
OUTPUT_DOCX_RESULT = 'OUTPUT4_DOCX-RESULT'


def cmd_extract_dims(args):
    """Extrait les dimensions du template DOCX (par défaut: assets/TEMPLATE.docx)"""
    print(f"\n🔧 ÉTAPE 1: EXTRACTION DES DIMENSIONS DU TEMPLATE")
    print(f"{'='*70}\n")
    
    dims = extract_page_dimensions_from_template(TEMPLATE_PATH)
    
    print(f"\n✅ Dimensions extraites:")
    for key, value in dims.items():
        print(f"  - {key}: {value}")
    
    return dims


def cmd_extract_xml(args):
    """Extrait le XML brut d'un DOCX"""
    print(f"\n📦 ÉTAPE 2: EXTRACTION DU XML")
    print(f"{'='*70}\n")
    
    docx_file = args.docx
    output_dir = args.output or OUTPUT_XML_RAW
    
    xml_file = export_all_xml(docx_file, output_dir)
    
    if xml_file:
        print(f"✅ XML extrait: {xml_file}")
    else:
        print("❌ Erreur lors de l'extraction du XML")
        sys.exit(1)
    
    return xml_file


def cmd_xml_to_json(args):
    """Convertit le XML en JSON RAW"""
    print(f"\n🔄 ÉTAPE 3: CONVERSION XML → JSON RAW")
    print(f"{'='*70}\n")
    
    xml_file = args.xml
    json_output = args.json_output
    
    json_file = xml_to_json(xml_file, json_output)
    
    if json_file:
        print(f"✅ JSON RAW créé: {json_file}")
    else:
        print("❌ Erreur lors de la conversion XML → JSON")
        sys.exit(1)
    
    return json_file


def cmd_transform(args):
    """Transforme le JSON RAW en JSON avec tags et styles"""
    print(f"\n✨ ÉTAPE 4: TRANSFORMATION JSON (tags + styles)")
    print(f"{'='*70}\n")
    
    json_raw = args.json_raw
    output_dir = args.output or OUTPUT_JSON_TRANSFORMED
    
    # Extraire les dimensions une seule fois (template par défaut)
    page_dims = extract_page_dimensions_from_template(TEMPLATE_PATH)
    
    json_transformed = apply_tags_and_styles(json_raw, output_dir, page_dims)
    
    if json_transformed:
        print(f"✅ JSON transformé créé: {json_transformed}")
    else:
        print("❌ Erreur lors de la transformation")
        sys.exit(1)
    
    return json_transformed


def cmd_render(args):
    """Rend le JSON transformé en DOCX"""
    print(f"\n🎨 ÉTAPE 5: RENDU JSON → DOCX")
    print(f"{'='*70}\n")
    
    json_file = args.json
    output_dir = args.output or OUTPUT_DOCX_RESULT
    
    docx_file = json_to_docx(json_file, TEMPLATE_PATH, output_dir)
    
    if docx_file:
        print(f"✅ DOCX généré: {docx_file}")
    else:
        print("❌ Erreur lors du rendu")
        sys.exit(1)
    
    return docx_file


def cmd_extract_all(args):
    """Combine: extract_dims + extract_xml + xml_to_json"""
    print(f"\n📁 PHASE 1: EXTRACTION (DIMS + XML + JSON RAW)")
    print(f"{'='*70}\n")
    
    docx_file = args.docx
    output_base = args.output or '.'
    output_base = Path(output_base)
    
    # Dossiers de sortie pour cette phase
    output_xml = output_base / OUTPUT_XML_RAW
    output_json = output_base / OUTPUT_JSON_RAW
    
    # 1. Extraire les dimensions du template
    print("▶ Extraction des dimensions du template...")
    page_dims = extract_page_dimensions_from_template(TEMPLATE_PATH)
    print(f"✓ Dimensions extraites\n")
    
    # 2. Extraire le XML
    print("▶ Extraction du XML brut...")
    xml_file = export_all_xml(docx_file, str(output_xml))
    if not xml_file:
        print("❌ Erreur lors de l'extraction du XML")
        sys.exit(1)
    print(f"✓ XML extrait: {Path(xml_file).name}\n")
    
    # 3. Convertir XML → JSON RAW
    print("▶ Conversion XML → JSON RAW...")
    docx_stem = Path(docx_file).stem
    json_output = output_json / f"{docx_stem}_GLOBAL_raw.json"
    json_file = xml_to_json(xml_file, str(json_output))
    if not json_file:
        print("❌ Erreur lors de la conversion")
        sys.exit(1)
    print(f"✓ JSON RAW créé: {Path(json_file).name}\n")
    
    print(f"\n{'='*70}")
    print(f"✅ PHASE 1 COMPLÉTÉE")
    print(f"{'='*70}\n")
    
    return {
        'page_dims': page_dims,
        'xml_file': xml_file,
        'json_raw': json_file
    }


def cmd_transform_and_render(args):
    """Combine: transform + render"""
    print(f"\n🎯 PHASE 2: TRANSFORMATION + RENDU")
    print(f"{'='*70}\n")
    
    docx_file = args.docx
    output_base = args.output or '.'
    output_base = Path(output_base)
    
    # Dossiers de sortie pour cette phase
    output_json_raw = output_base / OUTPUT_JSON_RAW
    output_json_transformed = output_base / OUTPUT_JSON_TRANSFORMED
    output_docx = output_base / OUTPUT_DOCX_RESULT
    
    # Inférer le JSON RAW depuis le DOCX
    docx_stem = Path(docx_file).stem
    json_raw = output_json_raw / f"{docx_stem}_GLOBAL_raw.json"
    
    if not json_raw.exists():
        print(f"❌ Erreur: {json_raw} non trouvé")
        print(f"   Exécutez d'abord: pipeline.py extract {docx_file}")
        sys.exit(1)
    
    # 1. Extraire les dimensions
    print("▶ Extraction des dimensions du template...")
    page_dims = extract_page_dimensions_from_template(TEMPLATE_PATH)
    print(f"✓ Dimensions extraites\n")
    
    # 2. Transformer le JSON RAW
    print("▶ Transformation du JSON RAW...")
    json_transformed = apply_tags_and_styles(str(json_raw), str(output_json_transformed), page_dims)
    if not json_transformed:
        print("❌ Erreur lors de la transformation")
        sys.exit(1)
    print(f"✓ JSON transformé: {Path(json_transformed).name}\n")
    
    # 3. Rendre en DOCX
    print("▶ Rendu JSON → DOCX...")
    docx_output = json_to_docx(json_transformed, TEMPLATE_PATH, str(output_docx))
    if not docx_output:
        print("❌ Erreur lors du rendu")
        sys.exit(1)
    print(f"✓ DOCX généré: {Path(docx_output).name}\n")
    
    print(f"\n{'='*70}")
    print(f"✅ PHASE 2 COMPLÉTÉE")
    print(f"{'='*70}\n")
    
    return {
        'page_dims': page_dims,
        'json_transformed': json_transformed,
        'docx_output': docx_output
    }


def cmd_pipeline_full(args):
    """Pipeline complète: extraction + transformation + rendu"""
    print(f"\n🚀 PIPELINE COMPLÈTE")
    print(f"{'='*70}\n")
    
    docx_file = args.docx
    output_base = args.output or '.'
    output_base = Path(output_base)
    
    # Dossiers de sortie pour toutes les phases
    output_xml = output_base / OUTPUT_XML_RAW
    output_json = output_base / OUTPUT_JSON_RAW
    output_json_transformed = output_base / OUTPUT_JSON_TRANSFORMED
    output_docx = output_base / OUTPUT_DOCX_RESULT
    
    # PHASE 1: Extraction
    print("📁 PHASE 1: EXTRACTION (DIMS + XML + JSON RAW)")
    print(f"{'-'*70}\n")
    
    # 1.1 Extraire les dimensions
    print("▶ Extraction des dimensions du template...")
    page_dims = extract_page_dimensions_from_template(TEMPLATE_PATH)
    print(f"✓ Dimensions extraites\n")
    
    # 1.2 Extraire le XML
    print("▶ Extraction du XML brut...")
    xml_file = export_all_xml(docx_file, str(output_xml))
    if not xml_file:
        print("❌ Erreur lors de l'extraction du XML")
        sys.exit(1)
    print(f"✓ XML extrait\n")
    
    # 1.3 Convertir XML → JSON RAW
    print("▶ Conversion XML → JSON RAW...")
    docx_stem = Path(docx_file).stem
    json_output = output_json / f"{docx_stem}_GLOBAL_raw.json"
    json_raw = xml_to_json(xml_file, str(json_output))
    if not json_raw:
        print("❌ Erreur lors de la conversion")
        sys.exit(1)
    print(f"✓ JSON RAW créé\n")
    
    print(f"{'-'*70}\n")
    
    # PHASE 2: Transformation + Rendu
    print("🎯 PHASE 2: TRANSFORMATION + RENDU")
    print(f"{'-'*70}\n")
    
    # 2.1 Transformer le JSON RAW
    print("▶ Transformation du JSON RAW...")
    json_transformed = apply_tags_and_styles(json_raw, str(output_json_transformed), page_dims)
    if not json_transformed:
        print("❌ Erreur lors de la transformation")
        sys.exit(1)
    print(f"✓ JSON transformé créé\n")
    
    # 2.2 Rendre en DOCX
    print("▶ Rendu JSON → DOCX...")
    docx_output = json_to_docx(json_transformed, TEMPLATE_PATH, str(output_docx))
    if not docx_output:
        print("❌ Erreur lors du rendu")
        sys.exit(1)
    print(f"✓ DOCX généré\n")
    
    print(f"\n{'='*70}")
    print(f"✅ PIPELINE COMPLÈTE RÉUSSIE")
    print(f"{'='*70}\n")
    print(f"Fichiers générés:")
    print(f"  - XML: {xml_file}")
    print(f"  - JSON RAW: {json_raw}")
    print(f"  - JSON transformé: {json_transformed}")
    print(f"  - DOCX: {docx_output}\n")


def main():
    """CLI principale avec subcommandes"""
    parser = argparse.ArgumentParser(
        description="Pipeline de traitement de documents DOCX",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
RÉSULTATS:
  Les fichiers sont organisés dans 4 dossiers distincts:
  - OUTPUT1_XML-RAW/        → XML global
  - OUTPUT2_JSON-RAW/       → JSON RAW (structure brute)
  - OUTPUT3_JSON-TRANSFORMED/ → JSON avec tags et styles
  - OUTPUT4_DOCX-RESULT/    → DOCX final généré

TEMPLATE:
  Utilise par défaut: assets/TEMPLATE.docx (pas besoin de le spécifier)

EXEMPLES:
  # Pipeline complète (d'un coup)
  python -m tools3.pipeline full document.docx
  
  # Pipeline complète avec résultats dans un dossier custom
  python -m tools3.pipeline full document.docx -o results/
  
  # Phase 1: Extraction (dims + xml + json raw)
  python -m tools3.pipeline extract document.docx
  
  # Phase 2: Transformation + Rendu (après phase 1)
  python -m tools3.pipeline transform-render document.docx
  
  # Commandes individuelles (rarement utilisées)
  python -m tools3.pipeline extract-dims
  python -m tools3.pipeline extract-xml document.docx
  python -m tools3.pipeline xml-to-json input.xml output.json
  python -m tools3.pipeline transform raw.json
  python -m tools3.pipeline render transformed.json
        """
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Commande à exécuter')
    
    # ===== COMMANDES INDIVIDUELLES =====
    
    # extract-dims
    extract_dims_parser = subparsers.add_parser(
        '--extract-dims',
        help='Extrait les dimensions du template DOCX (défaut: assets/TEMPLATE.docx)'
    )
    extract_dims_parser.set_defaults(func=cmd_extract_dims)
    
    # extract-xml
    extract_xml_parser = subparsers.add_parser(
        '--extract-xml',
        help='Extrait le XML brut d\'un DOCX'
    )
    extract_xml_parser.add_argument('-s', '--source_docx', help='Chemin du fichier source DOCX')
    extract_xml_parser.add_argument('-o', '--output_dir', help='Dossier de sortie (défaut: OUTPUT1_XML-RAW)')
    extract_xml_parser.set_defaults(func=cmd_extract_xml)
    
    # xml-to-json
    xml_to_json_parser = subparsers.add_parser(
        '--xml-to-json',
        help='Convertit XML en JSON RAW'
    )
    xml_to_json_parser.add_argument('-s', '--source_xml_raw', help='Chemin du fichier source XML RAW')
    xml_to_json_parser.add_argument('-o', '--output_dir', help='Dossier de sortie (défaut: OUTPUT2_JSON-RAW)')
    xml_to_json_parser.set_defaults(func=cmd_xml_to_json)
    
    # transform
    transform_parser = subparsers.add_parser(
        '--transform',
        help='Transforme le JSON RAW (tags + styles) avec template par défaut'
    )
    transform_parser.add_argument('-s', '--source_json_raw', help='Chemin du fichier source JSON RAW')
    transform_parser.add_argument('-o', '--output_dir', help='Dossier de sortie (défaut: OUTPUT3_JSON-TRANSFORMED)')
    transform_parser.set_defaults(func=cmd_transform)
    
    # render
    render_parser = subparsers.add_parser(
        '--render',
        help='Rend le JSON transformé en DOCX avec template par défaut'
    )
    render_parser.add_argument('-s', '--source_json_transformed', help='Chemin du fichier source JSON TRANSFORMED')
    render_parser.add_argument('-o', '--output_dir', help='Dossier de sortie (défaut: OUTPUT4_DOCX-RESULT)')
    render_parser.set_defaults(func=cmd_render)
    
    # ===== COMMANDES COMPOSÉES =====
    
    # extract
    extract_parser = subparsers.add_parser(
        'extract',
        help='Phase 1: Extraction (dims + xml + json raw) avec template par défaut'
    )
    extract_parser.add_argument('-s', '--source_docx', help='Chemin du fichier source DOCX')
    extract_parser.add_argument('-o', '--output_dir', help='Dossier parent des résultats (défaut: OUTPUT1_XML-RAW, OUTPUT2_JSON-RAW)')
    extract_parser.set_defaults(func=cmd_extract_all)
    
    # transform-render
    tr_parser = subparsers.add_parser(
        'transform-render',
        help='Phase 2: Transformation + Rendu (après extraction) avec template par défaut'
    )
    tr_parser.add_argument('-s', '--source_docx', help='Chemin du fichier source DOCX (pour inférer le JSON RAW)')
    tr_parser.add_argument('-o', '--output_dir', help='Dossier parent des résultats (défaut: OUTPUT3_JSON-TRANSFORMED, OUTPUT4_DOCX-RESULT)')
    tr_parser.set_defaults(func=cmd_transform_and_render)
    
    # full
    full_parser = subparsers.add_parser(
        'full',
        help='Pipeline complète: extraction + transformation + rendu avec template par défaut'
    )
    full_parser.add_argument('-s', '--source_docx', help='Chemin du fichier source DOCX')
    full_parser.add_argument('-o', '--output_dir', help='Dossier parent des résultats (défaut: OUTPUT1_XML-RAW, OUTPUT2_JSON-RAW, OUTPUT3_JSON-TRANSFORMED, OUTPUT4_DOCX-RESULT)')
    full_parser.set_defaults(func=cmd_pipeline_full)
    
    # ===== EXÉCUTION =====
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        sys.exit(0)
    
    # Exécuter la commande
    if hasattr(args, 'func'):
        args.func(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()