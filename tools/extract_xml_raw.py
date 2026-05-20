from argparse import ArgumentParser
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from datetime import datetime
from xml.dom import minidom


def indent_xml_string(xml_string):
    """
    Indente une chaîne XML avec minidom pour meilleure lisibilité.

    Args:
        xml_string (str): Chaîne XML brute

    Returns:
        str: Chaîne XML indentée et formatée
    """
    try:
        dom = minidom.parseString(xml_string)
        pretty_xml = dom.toprettyxml(indent="  ")
        # Supprimer la première ligne (déclaration XML auto-ajoutée par minidom)
        lines = pretty_xml.split('\n')[1:]
        return '\n'.join(lines).strip()
    except Exception as e:
        print(f"⚠ Erreur d'indentation: {e}")
        return xml_string


def extract_xml_raw(docx_file):
    """
    Extrait le contenu XML brut complet d'un document Word (.docx).

    Un fichier .docx est un archive ZIP contenant plusieurs fichiers XML.
    Cette fonction extrait tous les fichiers XML pertinents du document.

    Args:
        docx_file (str): Chemin vers le fichier .docx

    Returns:
        dict: Dictionnaire avec les fichiers XML trouvés
              Clés: chemin du fichier (ex: 'word/document.xml')
              Valeurs: contenu XML brut
    """
    xml_contents = {}

    try:
        # Ouvrir le .docx comme archive ZIP
        with zipfile.ZipFile(docx_file, 'r') as zip_ref:
            # Lister tous les fichiers de l'archive
            file_list = zip_ref.namelist()

            # Extraire tous les fichiers XML
            for file_name in file_list:
                if file_name.endswith('.xml'):
                    try:
                        xml_content = zip_ref.read(file_name).decode('utf-8')
                        xml_contents[file_name] = xml_content
                    except Exception as e:
                        print(f"Erreur lors de la lecture de {file_name}: {e}")

        if not xml_contents:
            print(f"Aucun fichier XML trouvé dans {docx_file}")
        else:
            print(f"{len(xml_contents)} fichier(s) XML extrait(s) avec succès")

    except zipfile.BadZipFile:
        print(f"Erreur: {docx_file} n'est pas un fichier ZIP valide (document Word valide?)")
    except FileNotFoundError:
        print(f"Erreur: Le fichier {docx_file} n'existe pas")
    except Exception as e:
        print(f"Erreur inattendue: {e}")

    return xml_contents


def create_global_xml(xml_contents, docx_name):
    """
    Crée un XML global combinant tous les fichiers XML d'un DOCX.
    Structure: <docx> contient chaque fichier XML comme section séparée

    Args:
        xml_contents (dict): Dictionnaire des fichiers XML extraits
        docx_name (str): Nom du fichier DOCX source

    Returns:
        str: XML global formaté
    """

    # Créer l'élément racine
    root = ET.Element('docx')
    root.set('source', docx_name)
    root.set('export-date', datetime.now().isoformat())
    root.set('xml-files', str(len(xml_contents)))

    # Ajouter chaque fichier XML comme section
    for xml_path in sorted(xml_contents.keys()):
        xml_content = xml_contents[xml_path]

        try:
            # Parser le contenu XML
            xml_elem = ET.fromstring(xml_content)

            # Créer un élément pour ce fichier
            file_elem = ET.SubElement(root, 'xml-file')
            file_elem.set('path', xml_path)
            file_elem.set('tag', xml_elem.tag)

            # Ajouter le contenu XML complet du fichier
            file_elem.append(xml_elem)

        except ET.ParseError as e:
            print(f"  ⚠ Erreur parsing {xml_path}: {e}")

    # Convertir en string
    global_xml_str = ET.tostring(root, encoding='unicode')

    # Indenter avec minidom
    indented = indent_xml_string(global_xml_str)

    return indented


def extract_document_xml(docx_file, output_dir):
    """
    Extrait uniquement le word/document.xml d'un DOCX et le sauvegarde formaté.
    Crée un fichier comme DC_JNZ_2026_RAW.xml

    Args:
        docx_file (str): Chemin du fichier DOCX
        output_dir (str): Répertoire de destination

    Returns:
        str: Chemin du fichier créé
    """
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)

    try:
        # Extraire le document.xml du DOCX
        with zipfile.ZipFile(docx_file, 'r') as zip_ref:
            doc_xml_content = zip_ref.read('word/document.xml').decode('utf-8')

        # Indenter
        indented_xml = indent_xml_string(doc_xml_content)

        # Sauvegarder
        docx_stem = Path(docx_file).stem
        output_file = output_path / f"{docx_stem}_RAW.xml"

        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(indented_xml)

        return str(output_file)

    except KeyError:
        return None
    except Exception as e:
        return None


def export_all_xml(docx_file, output_dir):
    """
    Fonction complète qui:
    Extrait tous les XML et crée un global

    Args:
        docx_file (str): Chemin du fichier DOCX
        output_dir (str): Répertoire de destination
    """
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)

    xml_contents = extract_xml_raw(docx_file)

    if not xml_contents:
        return None
    global_xml = create_global_xml(xml_contents, Path(docx_file).name)

    # Sauvegarder
    docx_stem = Path(docx_file).stem
    output_file = output_path / f"{docx_stem}_GLOBAL.xml"

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write(global_xml)

    file_size = output_file.stat().st_size / 1024
    print(f"\n✓ XML global créé: {output_file}")
    print(f"  Fichiers combined: {len(xml_contents)}")
    print(f"  Taille: {file_size:.1f} KB")

    print(f"\n{'='*70}\n")

    return str(output_file)


def main(docx_path, output_dir, export_docxml=False, export_allxml=False):
    """Fonction principale pour extraire tous les XML et exporter"""

    if export_docxml:
        # 1. Extraire le document.xml brut
        extract_document_xml(docx_path, output_dir)

    if export_allxml:
        # 2. Extraire tous les XML et créer un global
        export_all_xml(docx_path, output_dir)


if __name__ == "__main__":

    parser = ArgumentParser(description="Extrait et exporte le contenu XML d'un fichier .docx")

    # Argument positional: chemin du fichier DOCX
    parser.add_argument(
        "-s", "--source_docx_file",
        help="Chemin du fichier .docx à traiter"
    )

    # Arguments optionnels
    parser.add_argument(
        "-xi", "--export_docxml",
        action="store_true",
        help="Exporter uniquement le document.xml formaté (RAW)"
    )

    parser.add_argument(
        "-xn", "--export_allxml",
        action="store_true",
        help="Exporter tous les XML dans un fichier global"
    )

    parser.add_argument(
        "-o", "--output_dir",
        default="OUTPUT1_XML-RAW",
        help="Dossier de sortie (défaut: OUTPUT1_XML-RAW)"
    )

    args = parser.parse_args()

    if not args.export_docxml and not args.export_allxml:
        args.export_docxml = False
        args.export_allxml = True

    main(args.source_docx_file, args.output_dir, args.export_docxml, args.export_allxml)
