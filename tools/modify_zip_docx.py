"""
Script pour extraire/modifier et récompresser un DOCX archivé en ZIP.
Permet de modifier les fichiers XML directement puis de recréer le DOCX.
"""

from argparse import ArgumentParser
import shutil
from pathlib import Path
import zipfile
import subprocess

OUTPUT_DOCX_DIR = 'OUTPUT4_DOCX-RESULT'


def extract_zip(zip_file, extract_folder=None):
    """
    Extrait une archive ZIP.

    Args:
        zip_file (str): Chemin du fichier ZIP
        extract_folder (str): Dossier d'extraction (défaut: même nom que le ZIP sans extension)

    Returns:
        str: Chemin du dossier extrait ou None en cas d'erreur
    """
    try:
        zip_path = Path(zip_file)
        if not zip_path.exists():
            print(f"✗ Erreur: Le fichier {zip_file} n'existe pas")
            return None

        # Déterminer le dossier d'extraction
        if extract_folder is None:
            extract_folder = zip_path.stem + "_extracted"

        extract_path = Path(extract_folder)
        extract_path.mkdir(parents=True, exist_ok=True)

        # Extraire le ZIP
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_path)

        print(f"✓ Extrait: {extract_path}")
        print(f"  Dossier: {extract_path.absolute()}\n")

        # Afficher la structure
        print("📁 Structure du contenu:")
        for item in extract_path.rglob('*'):
            if item.is_file():
                rel_path = item.relative_to(extract_path)
                print(f"  {rel_path}")

        return str(extract_path)

    except Exception as e:
        print(f"✗ Erreur lors de l'extraction: {e}")
        return None


def repackage_docx(extracted_folder, output_docx=None):
    """
    Récompresse les fichiers modifiés en DOCX.

    Args:
        extracted_folder (str): Dossier contenant les fichiers extraits
        output_docx (str): Chemin du DOCX de sortie (défaut: basé sur le dossier)

    Returns:
        str: Chemin du DOCX recréé ou None en cas d'erreur
    """
    try:
        extracted_path = Path(extracted_folder)
        if not extracted_path.exists():
            print(f"✗ Erreur: Le dossier {extracted_folder} n'existe pas")
            return None

        # Déterminer le nom du DOCX de sortie
        if output_docx is None:
            # Extraire le nom de base du dossier extrait (enlever "_extracted")
            base_name = extracted_path.name.replace("_extracted", "")
            output_docx = OUTPUT_DOCX_DIR + "/" + base_name + ".docx"

        output_path = Path(output_docx)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # Créer le DOCX en réarchivant
        print(f"📦 Création du DOCX: {output_path}")
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx_out:
            for item in extracted_path.rglob('*'):
                if item.is_file():
                    arcname = item.relative_to(extracted_path)
                    docx_out.write(item, arcname)
                    print(f"  ✓ {arcname}")

        file_size = output_path.stat().st_size / (1024 * 1024)
        print(f"\n✓ DOCX créé: {output_path}")
        print(f"  Taille: {file_size:.2f} MB")

        return str(output_path)

    except Exception as e:
        print(f"✗ Erreur lors de la récompression: {e}")
        return None


def cleanup(folder):
    """Supprime le dossier extrait."""
    try:
        shutil.rmtree(folder)
        print(f"✓ Nettoyage: {folder} supprimé")
    except Exception as e:
        print(f"⚠ Erreur lors du nettoyage: {e}")


def main():
    parser = ArgumentParser(
        description='Extrait/modifie/récompresse un DOCX archivé en ZIP'
    )
    subparsers = parser.add_subparsers(dest='action', help='Action à effectuer')

    # Commande: extract
    extract_parser = subparsers.add_parser('extract', help='Extrait l\'archive ZIP')
    extract_parser.add_argument('-s', '--source', required=True, help='Chemin du ZIP')
    extract_parser.add_argument('-o', '--output', help='Dossier de destination')

    # Commande: repackage
    repackage_parser = subparsers.add_parser('repackage', help='Récompresse en DOCX')
    repackage_parser.add_argument('-s', '--source', required=True, help='Dossier extrait')
    repackage_parser.add_argument('-o', '--output', help='Chemin du DOCX de sortie')
    repackage_parser.add_argument('--cleanup', action='store_true', help='Supprimer le dossier après')

    # Commande: convert (extract + repackage en une seule)
    convert_parser = subparsers.add_parser(
        'convert',
        help='Extrait ZIP, modifie et récompresse en DOCX'
    )
    convert_parser.add_argument('-s', '--source', required=True, help='Chemin du ZIP source')
    convert_parser.add_argument('-o', '--output', help='Chemin du DOCX de sortie')
    convert_parser.add_argument('--edit', action='store_true', help='Ouvrir l\'éditeur après extraction')

    args = parser.parse_args()

    print(f"\n{'='*70}")
    print(f"🔧 MODIFICATION DE DOCX ARCHIVÉ EN ZIP")
    print(f"{'='*70}\n")

    if args.action == 'extract':
        extract_zip(args.source, args.output)

    elif args.action == 'repackage':
        repackage_docx(args.source, args.output)
        if args.cleanup:
            cleanup(args.source)

    elif args.action == 'convert':
        extracted = extract_zip(args.source)
        if extracted:
            if args.edit:
                print(f"\n💡 Tu peux maintenant modifier les fichiers XML dans: {extracted}")
                print("   En particulier: document.xml")
                input("\n➜ Appuie sur Entrée quand tu as fini les modifications...")

            repackage_docx(extracted, args.output)
            print(f"\n❓ Veux-tu supprimer le dossier extrait? [y/N] ", end='')
            if input().lower() == 'y':
                cleanup(extracted)

    else:
        parser.print_help()

    print(f"\n{'='*70}\n")


if __name__ == '__main__':
    main()
