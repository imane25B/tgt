import io
import fitz  # PyMuPDF
import extract_msg
import os

processed_msg_files = set()  # Ensemble pour suivre les fichiers .msg déjà traités

def extract_pdfs_from_msg(msg_path, output_dir):
    """
    Extrait les fichiers PDF d'un fichier .msg et gère les fichiers imbriqués.
    """
    if msg_path in processed_msg_files:
        print(f"⚠️ Fichier déjà traité : {msg_path}. Ignoré pour éviter les boucles.")
        return

    processed_msg_files.add(msg_path)  # Marquer le fichier comme traité

    try:
        msg = extract_msg.Message(msg_path)
    except Exception as e:
        print(f"❌ Erreur lors de l'ouverture de {msg_path} : {e}")
        return

    if not hasattr(msg, 'attachments') or not msg.attachments:
        print(f"Aucune pièce jointe trouvée dans {msg_path}.")
        return

    for attachment in msg.attachments:
        if not attachment.longFilename:
            print("⚠️ Pièce jointe sans nom détectée. Ignorée.")
            continue

        filename = attachment.longFilename.rstrip('\x00').lower()

        if filename.endswith('.pdf'):
            print(f"📄 PDF trouvé : {filename}")
            pdf_data = io.BytesIO(attachment.data)
            pdf_text = extract_text_from_pdf(pdf_data)

            if pdf_text.strip():
                output_file_path = os.path.join(output_dir, f"{filename}_extracted_text.txt")
                with open(output_file_path, "w", encoding="utf-8") as file:
                    file.write(pdf_text)
                print(f"✅ Texte extrait sauvegardé dans : {output_file_path}")
            else:
                print(f"⚠️ Aucun texte extrait de {filename}. Le fichier peut être scanné ou vide.")

        elif filename.endswith('.msg'):
            print(f"📧 Fichier .msg imbriqué trouvé : {filename}")
            nested_msg_path = os.path.join(output_dir, filename)
            with open(nested_msg_path, "wb") as nested_msg_file:
                nested_msg_file.write(attachment.data)
            
            # Appel récursif pour analyser le fichier .msg imbriqué
            extract_pdfs_from_msg(nested_msg_path, output_dir)

def extract_text_from_pdf(pdf_data):
    """
    Extrait le texte d'un fichier PDF à partir de ses données binaires.
    """
    text = ""
    try:
        with fitz.open(stream=pdf_data, filetype="pdf") as pdf_document:
            for page in pdf_document:
                text += page.get_text()
    except Exception as e:
        print(f"❌ Erreur lors de la lecture du PDF : {e}")
    
    return text
