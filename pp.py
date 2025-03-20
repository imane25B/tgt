import os
import extract_msg

def extract_pdf(pdf_path):
    """Fonction de traitement des fichiers PDF"""
    print(f"Traitement du PDF : {pdf_path}")

def process_msg_file(msg_path):
    """Traite un fichier .msg et vérifie les pièces jointes"""
    print(f"Traitement du fichier MSG : {msg_path}")
    msg = extract_msg.Message(msg_path)
    msg.attachments  # Liste des pièces jointes
    
    for attachment in msg.attachments:
        attachment_name = attachment.longFilename or attachment.shortFilename
        if not attachment_name:
            continue

        attachment_path = os.path.join(os.path.dirname(msg_path), attachment_name)

        # Sauvegarde temporaire de la pièce jointe
        with open(attachment_path, "wb") as f:
            f.write(attachment.data)

        # Vérification du type de fichier
        if attachment_name.lower().endswith(".pdf"):
            extract_pdf(attachment_path)
        elif attachment_name.lower().endswith(".msg"):
            process_msg_file(attachment_path)  # Appel récursif

        # Supprime le fichier temporaire après traitement
        os.remove(attachment_path)

def traverse_directory(root_dir):
    """Parcourt un dossier et ses sous-dossiers pour traiter les fichiers .msg"""
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.lower().endswith(".msg"):
                msg_path = os.path.join(dirpath, filename)
                process_msg_file(msg_path)

# Exemple d'utilisation
dossier_cible = "chemin/vers/le/dossier"
traverse_directory(dossier_cible)
