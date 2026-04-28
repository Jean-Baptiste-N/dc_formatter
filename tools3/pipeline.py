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

from argparse import ArgumentParser, RawDescriptionHelpFormatter
import sys
from pathlib import Path

from .extract_xml_raw import export_all_xml
from .parse_template import extract_page_dimensions_from_template
from .parse_xml_raw_to_json_raw import xml_to_json
from .process_json_raw_to_json_transformed import apply_tags_and_styles
from .render_json_transformed_to_docx import json_to_docx

# ===== CONSTANTES =====
TEMPLATE_PATH = 'assets/TEMPLATE.docx'
SOURCE_DOCX_DIR = 'DC_SOURCES'
OUTPUT_XML_RAW = 'OUTPUT1_XML-RAW'
OUTPUT_JSON_RAW = 'OUTPUT2_JSON-RAW'
OUTPUT_JSON_TRANSFORMED = 'OUTPUT3_JSON-TRANSFORMED'
OUTPUT_DOCX_RESULT = 'OUTPUT4_DOCX-RESULT'


def _resolve_source_path(filename):
    """Résout le chemin du fichier source
    Si le fichier n'existe pas en chemin absolu/relatif, cherche dans DC_SOURCES/
    """
    filepath = Path(filename)
    if filepath.exists():
        return str(filepath)

    # Chercher dans le dossier DC_SOURCES
    source_path = Path(SOURCE_DOCX_DIR) / filename
    if source_path.exists():
        return str(source_path)

    # Si rien ne correspond, retourner le chemin original
    return str(filepath)


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

    docx_file = _resolve_source_path(args.source)
    output_dir = args.output_dir or OUTPUT_XML_RAW

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

    xml_file = args.source
    json_output = None

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

    json_raw = args.source
    output_dir = args.output_dir or OUTPUT_JSON_TRANSFORMED

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

    json_file = args.source
    output_dir = args.output_dir or OUTPUT_DOCX_RESULT

    docx_file = json_to_docx(json_file, TEMPLATE_PATH, output_dir)

    if docx_file:
        print(f"✅ DOCX généré: {docx_file}")
    else:
        print("❌ Erreur lors du rendu")
        sys.exit(1)

    return docx_file


def cmd_extract_all(args):
    """Combine: extract_dims + extract_xml + xml_to_json

    -s: Fichier source dans DC_SOURCES/
    -o: Dossier de sortie DIRECT (pas de sous-dossiers OUTPUTX_*)
        Si non fourni, utilise les dossiers OUTPUTX_* par défaut
    """
    print("\n📁 Phase 1: Extraction...")

    docx_file = _resolve_source_path(args.source)

    # Dossiers de sortie pour cette phase
    if args.output_dir:
        # Si -o est fourni: utiliser directement ce dossier (SANS sous-dossiers)
        output_xml = Path(args.output_dir)
        output_json = Path(args.output_dir)
    else:
        # Sinon: utiliser les dossiers OUTPUTX_* par défaut
        output_xml = Path(OUTPUT_XML_RAW)
        output_json = Path(OUTPUT_JSON_RAW)

    # Créer les répertoires de sortie
    output_xml.mkdir(parents=True, exist_ok=True)
    output_json.mkdir(parents=True, exist_ok=True)

    # 1. Extraire les dimensions du template
    page_dims = extract_page_dimensions_from_template(TEMPLATE_PATH)

    # 2. Extraire le XML
    xml_file = export_all_xml(docx_file, str(output_xml))
    if not xml_file:
        print("❌ Erreur lors de l'extraction du XML")
        sys.exit(1)

    # 3. Convertir XML → JSON RAW
    json_file = xml_to_json(xml_file, str(output_json))
    if not json_file:
        print("❌ Erreur lors de la conversion")
        sys.exit(1)

    print(f"✅ Phase 1 complétée\n")

    return {
        'page_dims': page_dims,
        'xml_file': xml_file,
        'json_raw': json_file
    }


def cmd_transform_and_render(args):
    """Combine: transform + render

    Deux cas d'usage:
    1. -s = nom DOCX: Inférer le JSON RAW associé (comportement par défaut)
       Cherche JSON RAW toujours dans OUTPUT2_JSON-RAW/ (indépendant de -o)
    2. -s = chemin JSON RAW: Utiliser directement sans inférence

    -o: Dossier de sortie DIRECT pour résultats (SANS sous-dossiers OUTPUTX_*)
        Si non fourni, utilise les dossiers OUTPUTX_* par défaut
    """
    print("\n🎯 Phase 2: Transformation + Rendu...")

    source_path = Path(args.source)

    # Cas 1: -s est un JSON RAW (pas d'inférence)
    if args.source.endswith('.json'):
        json_raw = source_path

        if not json_raw.exists():
            print(f"❌ Erreur: {json_raw} non trouvé")
            sys.exit(1)

    # Cas 2: -s est un DOCX (inférence du JSON RAW)
    else:
        docx_file = _resolve_source_path(args.source)
        docx_stem = Path(docx_file).stem

        # JSON RAW input: chercher TOUJOURS dans OUTPUT2_JSON-RAW/ (par défaut)
        # -o n'affecte PAS la recherche du JSON RAW source, seulement la sortie finale
        json_raw = Path(OUTPUT_JSON_RAW) / f"{docx_stem}_GLOBAL_raw.json"

        if not json_raw.exists():
            print(f"❌ Erreur: {json_raw} non trouvé")
            print(f"   Exécutez d'abord: python3 -m tools3.pipeline extract -s {args.source}")
            sys.exit(1)

    # Dossiers de sortie pour cette phase
    if args.output_dir:
        # Si -o est fourni: utiliser directement ce dossier (SANS sous-dossiers)
        output_json_transformed = Path(args.output_dir)
        output_docx = Path(args.output_dir)
    else:
        # Sinon: utiliser les dossiers OUTPUTX_* par défaut
        output_json_transformed = Path(OUTPUT_JSON_TRANSFORMED)
        output_docx = Path(OUTPUT_DOCX_RESULT)

    # Créer les répertoires de sortie
    output_json_transformed.mkdir(parents=True, exist_ok=True)
    output_docx.mkdir(parents=True, exist_ok=True)

    # 1. Extraire les dimensions
    page_dims = extract_page_dimensions_from_template(TEMPLATE_PATH)

    # 2. Transformer le JSON RAW
    json_transformed = apply_tags_and_styles(str(json_raw), str(output_json_transformed), page_dims)
    if not json_transformed:
        print("❌ Erreur lors de la transformation")
        sys.exit(1)

    # 3. Rendre en DOCX
    docx_output = json_to_docx(json_transformed, TEMPLATE_PATH, str(output_docx))
    if not docx_output:
        print("❌ Erreur lors du rendu")
        sys.exit(1)

    print(f"✅ Phase 2 complétée\n")

    return {
        'page_dims': page_dims,
        'json_transformed': json_transformed,
        'docx_output': docx_output
    }


def cmd_pipeline_full(args):
    """Pipeline complète: extraction + transformation + rendu

    -s: Fichier source dans DC_SOURCES/
    -o: Dossier de sortie DIRECT pour tous les résultats (SANS sous-dossiers OUTPUTX_*)
        Si non fourni, utilise les dossiers OUTPUTX_* par défaut
    """
    print("\n🚀 Pipeline complète...")

    docx_file = _resolve_source_path(args.source)

    # Dossiers de sortie pour toutes les phases
    if args.output_dir:
        # Si -o est fourni: utiliser directement ce dossier (SANS sous-dossiers)
        output_xml = Path(args.output_dir)
        output_json = Path(args.output_dir)
        output_json_transformed = Path(args.output_dir)
        output_docx = Path(args.output_dir)
    else:
        # Sinon: utiliser les dossiers OUTPUTX_* par défaut
        output_xml = Path(OUTPUT_XML_RAW)
        output_json = Path(OUTPUT_JSON_RAW)
        output_json_transformed = Path(OUTPUT_JSON_TRANSFORMED)
        output_docx = Path(OUTPUT_DOCX_RESULT)

    # Créer les répertoires de sortie
    output_xml.mkdir(parents=True, exist_ok=True)
    output_json.mkdir(parents=True, exist_ok=True)
    output_json_transformed.mkdir(parents=True, exist_ok=True)
    output_docx.mkdir(parents=True, exist_ok=True)

    # 1. Extraire les dimensions
    page_dims = extract_page_dimensions_from_template(TEMPLATE_PATH)

    # 2. Extraire le XML
    xml_file = export_all_xml(docx_file, str(output_xml))
    if not xml_file:
        print("❌ Erreur lors de l'extraction du XML")
        sys.exit(1)

    # 3. Convertir XML → JSON RAW
    json_raw = xml_to_json(xml_file, str(output_json))
    if not json_raw:
        print("❌ Erreur lors de la conversion XML")
        sys.exit(1)

    # 4. Transformer le JSON RAW
    json_transformed = apply_tags_and_styles(json_raw, str(output_json_transformed), page_dims)
    if not json_transformed:
        print("❌ Erreur lors de la transformation")
        sys.exit(1)

    # 5. Rendre en DOCX
    docx_output = json_to_docx(json_transformed, TEMPLATE_PATH, str(output_docx))
    if not docx_output:
        print("❌ Erreur lors du rendu")
        sys.exit(1)

    print(f"✅ Pipeline réussi")
    print(f"  - XML: {xml_file}")
    print(f"  - JSON RAW: {json_raw}")
    print(f"  - JSON TRANSFORMED: {json_transformed}")
    print(f"  - DOCX: {docx_output}\n")


def main():
    """CLI principale avec subcommandes"""
    parser = ArgumentParser(
        description="Pipeline de traitement de documents DOCX",
        formatter_class=RawDescriptionHelpFormatter,
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
        'extract-dims',
        help='Extrait les dimensions du template DOCX (défaut: assets/TEMPLATE.docx)'
    )
    extract_dims_parser.set_defaults(func=cmd_extract_dims)

    # extract-xml
    extract_xml_parser = subparsers.add_parser(
        'extract-xml',
        help='Extrait le XML brut d\'un DOCX'
    )
    extract_xml_parser.add_argument('-s', '--source', help='Nom du fichier DOCX dans DC_SOURCES/ ou chemin complet')
    extract_xml_parser.add_argument('-o', '--output_dir', help='Dossier de sortie (défaut: OUTPUT1_XML-RAW)')
    extract_xml_parser.set_defaults(func=cmd_extract_xml)

    # xml-to-json
    xml_to_json_parser = subparsers.add_parser(
        'xml-to-json',
        help='Convertit XML en JSON RAW'
    )
    xml_to_json_parser.add_argument('-s', '--source', help='Chemin du fichier source XML RAW')
    xml_to_json_parser.add_argument('-o', '--output_dir', help='Dossier de sortie (défaut: OUTPUT2_JSON-RAW)')
    xml_to_json_parser.set_defaults(func=cmd_xml_to_json)

    # transform
    transform_parser = subparsers.add_parser(
        'transform',
        help='Transforme le JSON RAW (tags + styles) avec template par défaut'
    )
    transform_parser.add_argument('-s', '--source', help='Chemin du fichier source JSON RAW')
    transform_parser.add_argument('-o', '--output_dir', help='Dossier de sortie (défaut: OUTPUT3_JSON-TRANSFORMED)')
    transform_parser.set_defaults(func=cmd_transform)

    # render
    render_parser = subparsers.add_parser(
        'render',
        help='Rend le JSON transformé en DOCX avec template par défaut'
    )
    render_parser.add_argument('-s', '--source', help='Chemin du fichier source JSON TRANSFORMED')
    render_parser.add_argument('-o', '--output_dir', help='Dossier de sortie (défaut: OUTPUT4_DOCX-RESULT)')
    render_parser.set_defaults(func=cmd_render)

    # ===== COMMANDES COMPOSÉES =====

    # extract
    extract_parser = subparsers.add_parser(
        'extract',
        help='Phase 1: Extraction (dims + xml + json raw) avec template par défaut'
    )
    extract_parser.add_argument('-s', '--source', required=True, help='Nom du fichier DOCX dans DC_SOURCES/ ou chemin complet')
    extract_parser.add_argument('-o', '--output_dir', help='Dossier parent des résultats (défaut: OUTPUT1_XML-RAW, OUTPUT2_JSON-RAW)')
    extract_parser.set_defaults(func=cmd_extract_all)

    # transform-render
    tr_parser = subparsers.add_parser(
        'transform-render',
        help='Phase 2: Transformation + Rendu (après extraction) avec template par défaut'
    )
    tr_parser.add_argument('-s', '--source', required=True, help='Nom du fichier DOCX dans DC_SOURCES/ ou chemin complet')
    tr_parser.add_argument('-o', '--output_dir', help='Dossier parent des résultats (défaut: OUTPUT3_JSON-TRANSFORMED, OUTPUT4_DOCX-RESULT)')
    tr_parser.set_defaults(func=cmd_transform_and_render)

    # full
    full_parser = subparsers.add_parser(
        'full',
        help='Pipeline complète: extraction + transformation + rendu avec template par défaut'
    )
    full_parser.add_argument('-s', '--source', required=True, help='Nom du fichier DOCX dans DC_SOURCES/ ou chemin complet')
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