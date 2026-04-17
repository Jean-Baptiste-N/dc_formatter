"""
Script pour archiver les fichiers DOCX et indenter tous les XML découverts.
Utilise lxml pour l'indentation et ET en fallback.
"""

from argparse import ArgumentParser
import shutil
from datetime import datetime
from pathlib import Path
import zipfile
from xml.etree import ElementTree as ET


def archive_docx(docx_file, archive_folder='archive'):
    """
    Compresse un fichier DOCX et l'enregistre dans un dossier archive.

    Args:
        docx_file (str): Chemin du fichier DOCX à archiver
        archive_folder (str): Dossier de destination (par défaut: 'archive')

    Returns:
        str: Chemin du fichier archivé ou None en cas d'erreur
    """
    try:
        # Créer le dossier archive s'il n'existe pas
        archive_path = Path(archive_folder)
        archive_path.mkdir(exist_ok=True)

        # Récupérer le nom du fichier
        original_path = Path(docx_file)
        if not original_path.exists():
            print(f"✗ Erreur: Le fichier {docx_file} n'existe pas")
            return None

        # Créer le nom du fichier archivé avec timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_stem = original_path.stem
        archived_file = archive_path / f"{file_stem}_{timestamp}.zip"

        # Copier et renommer le fichier
        shutil.copy2(docx_file, archived_file)

        file_size = archived_file.stat().st_size / (1024 * 1024)  # En MB
        print(f"✓ Archivé: {archived_file}")
        print(f"  Taille: {file_size:.2f} MB")

        return str(archived_file)

    except Exception as e:
        print(f"✗ Erreur lors de l'archivage: {e}")
        return None


def indent_xml(elem, level=0):
    """
    Indente un élément XML pour une meilleure lisibilité (fallback ElementTree).

    Args:
        elem (ET.Element): Élément XML à indenter
        level (int): Niveau d'indentation (par défaut: 0)
    """
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        for child in elem:
            indent_xml(child, level + 1)
        if not child.tail or not child.tail.strip():
            child.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i


def indent_xml_files_in_archive(archive_folder='archive'):
    """
    Découvre tous les fichiers ZIP dans le dossier archive,
    extrait les XML, les indente avec lxml/ET et les réenregistre.

    Args:
        archive_folder (str): Dossier contenant les archives
    """
    archive_path = Path(archive_folder)

    if not archive_path.exists():
        print(f"✗ Dossier {archive_folder} n'existe pas")
        return

    # Trouver tous les fichiers ZIP (.docx et .zip)
    zip_files = list(archive_path.glob('*.zip')) + list(archive_path.glob('*.docx'))

    if not zip_files:
        print(f"✗ Aucun fichier ZIP/DOCX trouvé dans {archive_folder}")
        return

    print(f"📁 Traitement de {len(zip_files)} fichier(s) archive\n")

    for zip_file in zip_files:
        print(f"📦 Traitement: {zip_file.name}")
        xml_count = 0

        try:
            # Créer un nouveau ZIP avec XML indentés
            temp_zip = zip_file.with_suffix('.temp.zip')

            with zipfile.ZipFile(zip_file, 'r') as zip_in:
                with zipfile.ZipFile(temp_zip, 'w', zipfile.ZIP_DEFLATED) as zip_out:

                    for item_info in zip_in.infolist():
                        content = zip_in.read(item_info.filename)

                        # Traiter les fichiers XML
                        if item_info.filename.endswith('.xml'):
                            try:
                                # Parser avec lxml (plus robuste)
                                try:
                                    from lxml import etree as lxml_et
                                    root = lxml_et.fromstring(content)

                                    # Indenter avec lxml
                                    lxml_et.indent(root, space="  ")

                                    # Convertir en bytes avec déclaration XML
                                    indented_content = lxml_et.tostring(
                                        root,
                                        pretty_print=True,
                                        xml_declaration=True,
                                        encoding='UTF-8'
                                    )

                                    zip_out.writestr(item_info, indented_content)
                                    print(f"  ✓ {item_info.filename} (lxml)")
                                    xml_count += 1

                                except ImportError:
                                    # Fallback avec ElementTree si lxml n'est pas disponible
                                    root = ET.fromstring(content)
                                    indent_xml(root)
                                    indented_content = ET.tostring(root, encoding='utf-8')
                                    zip_out.writestr(item_info, indented_content)
                                    print(f"  ✓ {item_info.filename} (ET)")
                                    xml_count += 1

                            except Exception as e:
                                print(f"  ⚠ Erreur pour {item_info.filename}: {e}")
                                zip_out.writestr(item_info, content)
                        else:
                            # Copier les autres fichiers
                            zip_out.writestr(item_info, content)

            # Remplacer l'original par le fichier traité
            shutil.move(str(temp_zip), str(zip_file))
            file_size = zip_file.stat().st_size / (1024 * 1024)
            print(f"  ✓ {xml_count} fichier(s) XML indentés - Taille: {file_size:.2f} MB\n")

        except Exception as e:
            print(f"  ✗ Erreur: {e}\n")
            if temp_zip.exists():
                temp_zip.unlink()


def main(docx_path, archive_folder="archive"):
    """Fonction principale pour archiver et indenter les XML."""
    print(f"\n{'='*70}")
    print(f"🗂️  ARCHIVAGE ET INDENTATION XML")
    print(f"{'='*70}\n")

    # Archiver le fichier DOCX
    archived = archive_docx(docx_path, archive_folder)

    if archived:
        print(f"\n✓ Fichier archivé avec succès!")
        print(f"{'='*70}\n")

        # Indenter tous les XML des archives
        indent_xml_files_in_archive(archive_folder=archive_folder)

        print(f"\n{'='*70}")
        print(f"✓ Tous les XML ont été indentés!")
        print(f"{'='*70}\n")
    else:
        print("✗ Archivage échoué")


if __name__ == "__main__":

    parser = ArgumentParser(description="Zippe le contenu d'un fichier .docx dans une archive et indente tous les fichiers XML découverts.")

    # Argument positional: chemin du fichier DOCX
    parser.add_argument(
        "docx_file",
        help="Chemin du fichier .docx à traiter"
    )

    parser.add_argument(
        "-o", "--output",
        default="archive",
        help="Dossier de sortie (défaut: archive)"
    )

    args = parser.parse_args()

    main(args.docx_file, args.output)