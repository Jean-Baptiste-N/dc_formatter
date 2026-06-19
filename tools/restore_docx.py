"""
Script pour régénérer un fichier DOCX à partir d'une archive ZIP.
Inverse du script zip_docx.py
"""

from argparse import ArgumentParser
import shutil
from pathlib import Path
import zipfile

OUTPUT_DOCX_DIR = 'OUTPUT4_DOCX-RESULT'


def restore_docx(zip_file, output_folder=OUTPUT_DOCX_DIR):
    """
    Restaure un fichier DOCX à partir d'une archive ZIP.

    Args:
        zip_file (str): Chemin du fichier ZIP à restaurer
        output_folder (str): Dossier de destination (par défaut: OUTPUT4_DOCX-RESULT)

    Returns:
        str: Chemin du fichier DOCX restauré ou None en cas d'erreur
    """
    try:
        # Vérifier que le fichier ZIP existe
        zip_path = Path(zip_file)
        if not zip_path.exists():
            print(f"✗ Erreur: Le fichier {zip_file} n'existe pas")
            return None

        # Créer le dossier de sortie s'il n'existe pas
        output_path = Path(output_folder)
        output_path.mkdir(parents=True, exist_ok=True)

        # Créer le nom du fichier DOCX restauré
        # Enlever le timestamp du nom (exemple: DC_CBD_2025_VF_20260619_113107.zip -> DC_CBD_2025_VF_restored.docx)
        file_stem = zip_path.stem
        # Si le stem contient un timestamp (pattern: NAME_YYYYMMDD_HHMMSS), l'enlever
        parts = file_stem.rsplit('_', 2)
        if len(parts) == 3 and len(parts[1]) == 8 and len(parts[2]) == 6:
            # C'est un timestamp, on le retire
            base_name = parts[0]
        else:
            base_name = file_stem

        restored_file = output_path / f"{base_name}_restored.docx"

        # Extraire le contenu du ZIP en tant que DOCX
        with zipfile.ZipFile(zip_path, 'r') as zip_in:
            zip_in.extractall(path=restored_file.parent / f".{restored_file.stem}_temp")

        # Créer le DOCX en réarchivant le contenu
        temp_dir = restored_file.parent / f".{restored_file.stem}_temp"
        with zipfile.ZipFile(restored_file, 'w', zipfile.ZIP_DEFLATED) as docx_out:
            for item in temp_dir.rglob('*'):
                if item.is_file():
                    arcname = item.relative_to(temp_dir)
                    docx_out.write(item, arcname)

        # Nettoyer le dossier temporaire
        shutil.rmtree(temp_dir)

        file_size = restored_file.stat().st_size / (1024 * 1024)  # En MB
        print(f"✓ Restauré: {restored_file}")
        print(f"  Taille: {file_size:.2f} MB")

        return str(restored_file)

    except Exception as e:
        print(f"✗ Erreur lors de la restauration: {e}")
        return None


def main():
    """Fonction principale avec arguments CLI."""
    parser = ArgumentParser(
        description='Régénère un fichier DOCX à partir d\'une archive ZIP'
    )
    parser.add_argument(
        '-s', '--source',
        required=True,
        help='Chemin du fichier ZIP à restaurer'
    )
    parser.add_argument(
        '-o', '--output',
        default=OUTPUT_DOCX_DIR,
        help=f'Dossier de destination (défaut: {OUTPUT_DOCX_DIR})'
    )

    args = parser.parse_args()

    print(f"\n{'='*70}")
    print(f"🔄 RESTAURATION DOCX DEPUIS ARCHIVE ZIP")
    print(f"{'='*70}\n")

    restored = restore_docx(args.source, args.output)

    if restored:
        print(f"\n✓ Fichier restauré avec succès!")
        print(f"{'='*70}\n")
    else:
        print(f"\n✗ Erreur lors de la restauration")
        print(f"{'='*70}\n")


if __name__ == '__main__':
    main()
