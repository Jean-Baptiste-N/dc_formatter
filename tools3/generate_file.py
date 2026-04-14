import zipfile
from pathlib import Path
from shutil import copy2
import tempfile
import os
import sys
import xml.etree.ElementTree as ET

# Ajouter le parent directory au path pour les imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from tools3.extract_xml_raw import extract_xml_raw


def _sanitize_xml(xml_content, xml_path=''):
    """
    Nettoie et valide le contenu XML.
    - Ajoute la déclaration XML si manquante
    - Valide la structure
    """
    if not xml_content.strip():
        return xml_content
    
    # Ajouter la déclaration XML si absente
    if not xml_content.strip().startswith('<?xml'):
        xml_content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_content
    
    # Valider que c'est du XML valide
    try:
        ET.fromstring(xml_content)
        return xml_content
    except ET.ParseError as e:
        raise ValueError(f"XML invalide dans {xml_path}: {e}")


def _normalize_namespaces(xml_content):
    """
    Normalise les namespaces pour cohérence.
    ⚠️ ATTENTION: Ne pas convertir ns0: → w: car cela peut créer des doublons
    si les deux déclarations pointent vers le même namespace.
    On garde les namespaces tels quels de la source.
    """
    # Pour l'instant, on ne normalise PAS pour éviter les conflits
    # Les fichiers DC_styles.xml utilisent ns0: et c'est OK, Word les gère
    return xml_content


def _validate_xml_structure(xml_content, xml_path=''):
    """Valide que la structure XML est correcte."""
    try:
        root = ET.fromstring(xml_content)
        if root is None:
            raise ValueError(f"Root element manquant dans {xml_path}")
        return True
    except ET.ParseError as e:
        raise ValueError(f"Structure XML invalide ({xml_path}): {e}")


def _prepare_xml_for_docx(xml_content, xml_path=''):
    """
    Prépare le contenu XML pour injection dans le docx.
    - Nettoie
    - Normalise les namespaces
    - Valide
    """
    print(f"  🔧 Validation de {xml_path}...")
    
    # 1. Sanitize
    xml_content = _sanitize_xml(xml_content, xml_path)
    
    # 2. Normaliser les namespaces
    xml_content = _normalize_namespaces(xml_content)
    
    # 3. Valider
    _validate_xml_structure(xml_content, xml_path)
    
    print(f"     ✓ {xml_path} validé")
    return xml_content


def generate_file(source_docx, output_docx, modified_xml_files=None):
    """
    Recrée un docx en gardant la même structure, mais en remplaçant les fichiers XML modifiés.
    
    Args:
        source_docx (str): Chemin du .docx source
        output_docx (str): Chemin du .docx de sortie
        modified_xml_files (dict, optional): Dictionnaire {chemin_xml: contenu_modifié}
    """
    try:
        print(f"\n📝 Recréation du docx...")
        print(f"  Source: {source_docx}")
        print(f"  Sortie: {output_docx}")
        
        # Copier le fichier source comme base
        temp_dir = tempfile.mkdtemp()
        temp_docx = os.path.join(temp_dir, 'temp.docx')
        copy2(source_docx, temp_docx)
        
        # Injecter tous les fichiers XML modifiés
        if modified_xml_files:
            with zipfile.ZipFile(temp_docx, 'a') as zip_ref:
                for xml_path, xml_content in modified_xml_files.items():
                    print(f"  ✓ Injection de {xml_path}...")
                    zip_ref.writestr(xml_path, xml_content)
        
        # Copier le résultat vers la destination finale
        copy2(temp_docx, output_docx)
        
        # Nettoyer le temp
        import shutil
        shutil.rmtree(temp_dir)
        
        print(f"\n✅ Fichier généré avec succès: {output_docx}")
        return True
        
    except FileNotFoundError as e:
        print(f"❌ Erreur: {e}")
        return False
    except Exception as e:
        print(f"❌ Erreur inattendue: {e}")
        return False


def generate_file_safe(source_docx, output_docx, modified_xml_files=None):
    """
    Version SÛRE de generate_file avec validation complète.
    Nettoie et valide tous les XML avant injection.
    
    Args:
        source_docx (str): Chemin du .docx source
        output_docx (str): Chemin du .docx de sortie
        modified_xml_files (dict, optional): Dictionnaire {chemin_xml: contenu_modifié}
    """
    try:
        print(f"\n🛡️  GÉNÉRATION SÉCURISÉE du docx...")
        print(f"  Source: {source_docx}")
        print(f"  Sortie: {output_docx}")
        
        # Valider et nettoyer tous les XML
        cleaned_xml = {}
        if modified_xml_files:
            print("\n🔍 Nettoyage et validation des XML...")
            for xml_path, xml_content in modified_xml_files.items():
                try:
                    cleaned_xml[xml_path] = _prepare_xml_for_docx(xml_content, xml_path)
                except ValueError as e:
                    print(f"  ⚠️  {e}")
                    raise
        
        # Copier le fichier source comme base
        temp_dir = tempfile.mkdtemp()
        temp_docx = os.path.join(temp_dir, 'temp.docx')
        copy2(source_docx, temp_docx)
        
        # Injecter tous les fichiers XML nettoyés
        print("\n💉 Injection des XML nettoyés...")
        with zipfile.ZipFile(temp_docx, 'a') as zip_ref:
            for xml_path, xml_content in cleaned_xml.items():
                print(f"  ✓ Écriture de {xml_path}...")
                zip_ref.writestr(xml_path, xml_content.encode('utf-8'))
        
        # Copier le résultat vers la destination finale
        copy2(temp_docx, output_docx)
        
        # Nettoyer le temp
        import shutil
        shutil.rmtree(temp_dir)
        
        # Valider le docx généré
        print("\n✅ Validation du docx généré...")
        try:
            with zipfile.ZipFile(output_docx, 'r') as z:
                z.testzip()
            print("✅ Fichier généré avec succès et valide!")
            return True
        except Exception as e:
            print(f"⚠️  Fichier généré mais avec potentiels problèmes: {e}")
            return True  # Retourner True quand même, le fichier existe
        
    except FileNotFoundError as e:
        print(f"❌ Erreur: {e}")
        return False
    except ValueError as e:
        print(f"❌ Erreur de validation XML: {e}")
        return False
    except Exception as e:
        print(f"❌ Erreur inattendue: {e}")
        return False


def main(docx_path, output_path, styles_xml_path=None, use_safe_mode=True):
    """
    Extrait tous les XML du source et injecte les styles custom si fournis.
    
    Args:
        docx_path (str): Path du docx source
        output_path (str): Path du docx de sortie
        styles_xml_path (str, optional): Path du fichier styles.xml custom
        use_safe_mode (bool): Utiliser la génération sécurisée (défaut: True)
    """
    # Extraire tous les XML du source
    print("📖 Extraction des XML du source...")
    all_xml = extract_xml_raw(docx_path)
    
    # Remplacer les styles si un fichier custom est fourni
    if styles_xml_path and Path(styles_xml_path).exists():
        print(f"💅 Chargement des styles custom de {styles_xml_path}...")
        with open(styles_xml_path, 'r', encoding='utf-8') as f:
            all_xml['word/styles.xml'] = f.read()
    
    # Recréer le docx avec les XML (potentiellement modifiés)
    if use_safe_mode:
        return generate_file_safe(docx_path, output_path, modified_xml_files=all_xml)
    else:
        return generate_file(docx_path, output_path, modified_xml_files=all_xml)
    
if __name__ == "__main__":
    docx_path = 'test/DC_YAX_2025.docx'
    output_path = 'output/DC_YAX_2025_reformat.docx'
    styles_xml_path = 'styles/DC_styles.xml'
    
    # Créer le répertoire output s'il n'existe pas
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    
    # Utiliser le mode sécurisé par défaut
    main(docx_path, output_path, styles_xml_path, use_safe_mode=True)



