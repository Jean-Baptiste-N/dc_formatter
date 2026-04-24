import argparse

from .extract_xml_raw import extract_xml_raw
from .parse_template import extract_page_dimensions_from_template
from .parse_xml_raw_to_json_raw import parse_xml_raw_to_json_raw
from .process_json_raw_to_json_transformed import transform_json_raw_to_transformed
from .render_json_transformed_to_docx import render_transformed_to_docx


def main():
    parser = argparse.ArgumentParser(description="Pipeline d'extraction, transformation, rendu, et de parsing du template")
    parser.add_argument('--docx', type=str, required=True, help='Chemin du fichier .docx à traiter')
    args = parser.parse_args()

    print(f"📁 Fichier DOCX à traiter: {args.docx}")

    # Étape 1: Extraire le XML brut
    xml_path = extract_xml_raw(args.docx)
    # Étape 2: Extraire les dimensions du template
    dimensions = extract_page_dimensions_from_template('assets/TEMPLATE.docx')
    # Étape 3: Parser le XML brut en JSON brut
    json_raw = parse_xml_raw_to_json_raw(xml_path)
    # Étape 4: Transformer le JSON brut en JSON transformé
    json_transformed = transform_json_raw_to_transformed(json_raw, dimensions)
    # Étape 5: Rendre le JSON transformé en DOCX
    render_transformed_to_docx(json_transformed, dimensions)

if __name__ == "__main__":
    main()