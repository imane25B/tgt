import os
import extract_msg
import io
import sys
from pathlib import Path

# Fonction supposée définie ailleurs pour extraire des PDF
def extract_pdf(pdf_data):
    print("Traitement d'un PDF en mémoire")
    # Votre code d'extraction pour les PDF (avec les données binaires directement)
    # ...


def process_msg_data(msg_data, filename="fichier imbriqué"):
    """
    Traite les données binaires d'un fichier .msg
    """
    try:
        # Utiliser BytesIO pour travailler en mémoire sans écrire sur le disque
        msg_io = io.BytesIO(msg_data)
        msg = extract_msg.Message(msg_io)
        print(f"Traitement des données MSG: {filename}")
        print(f"Sujet: {msg.subject}")
        
        # Parcourir toutes les pièces jointes
        for attachment in msg.attachments:
            # Obtenir le nom du fichier joint
            attachment_filename = attachment.longFilename or attachment.shortFilename
            if not attachment_filename:
                continue
            
            # Obtenir les données binaires de la pièce jointe
            attachment_data = attachment.data
            
            # Vérifier l'extension du fichier
            _, file_extension = os.path.splitext(attachment_filename.lower())
            
            # Traiter selon le type de fichier
            if file_extension == '.pdf':
                extract_pdf(attachment_data)
            elif file_extension == '.msg':
                # Appel récursif pour traiter les fichiers MSG imbriqués
                process_msg_data(attachment_data, attachment_filename)
            else:
                print(f"Pièce jointe ignorée: {attachment_filename} (extension non prise en charge)")
        
        # Fermer le fichier MSG
        msg.close()
    
    except Exception as e:
        print(f"Erreur lors du traitement des données MSG {filename}: {e}")


def process_msg_file(msg_path):
    """
    Traite un fichier .msg et ses pièces jointes directement depuis le fichier
    """
    try:
        msg = extract_msg.Message(msg_path)
        print(f"Traitement du fichier MSG: {msg_path}")
        print(f"Sujet: {msg.subject}")
        
        # Parcourir toutes les pièces jointes
        for attachment in msg.attachments:
            # Obtenir le nom du fichier joint
            attachment_filename = attachment.longFilename or attachment.shortFilename
            if not attachment_filename:
                continue
            
            # Obtenir les données binaires de la pièce jointe
            attachment_data = attachment.data
            
            # Vérifier l'extension du fichier
            _, file_extension = os.path.splitext(attachment_filename.lower())
            
            # Traiter selon le type de fichier
            if file_extension == '.pdf':
                extract_pdf(attachment_data)
            elif file_extension == '.msg':
                # Appel récursif pour traiter les fichiers MSG imbriqués
                process_msg_data(attachment_data, attachment_filename)
            else:
                print(f"Pièce jointe ignorée: {attachment_filename} (extension non prise en charge)")
        
        # Fermer le fichier MSG
        msg.close()
    
    except Exception as e:
        print(f"Erreur lors du traitement du fichier {msg_path}: {e}")


def scan_directory_for_msg(directory):
    """
    Parcourt récursivement un dossier et traite tous les fichiers .msg
    """
    try:
        # Convertir en chemin absolu
        directory = os.path.abspath(directory)
        print(f"Scanning du dossier: {directory}")
        
        # Parcourir le dossier et ses sous-dossiers
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.lower().endswith('.msg'):
                    full_path = os.path.join(root, file)
                    process_msg_file(full_path)
    
    except Exception as e:
        print(f"Erreur lors du scan du dossier {directory}: {e}")


if __name__ == "__main__":
    # Vérification des arguments
    if len(sys.argv) != 2:
        print("Usage: python script.py <dossier_à_scanner>")
        sys.exit(1)
    
    directory_to_scan = sys.argv[1]
    
    if not os.path.isdir(directory_to_scan):
        print(f"Erreur: '{directory_to_scan}' n'est pas un dossier valide.")
        sys.exit(1)
    
    # Lancer le scan
    scan_directory_for_msg(directory_to_scan)
    print("Traitement terminé.")
