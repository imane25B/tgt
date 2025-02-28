import os
import extract_msg
import fitz 
import io

def extract_pdfs_from_msg(msg_path, output_dir):
    """
    Extrait les fichiers PDF d'un fichier .msg et les stocke dans output_dir.
    """
    msg = extract_msg.Message(msg_path)

    if not hasattr(msg, 'attachments') or not msg.attachments:
        print("Aucune pièce jointe trouvée dans le fichier .msg.")
        return {}

    pdf_texts = {}

    for attachment in msg.attachments:
        filename = attachment.longFilename.rstrip('\x00').lower()
        
        if filename.endswith('.pdf'):
            print(f"📄 PDF trouvé : {filename}")
            
            # Lecture du contenu PDF en mémoire
            pdf_data = io.BytesIO(attachment.data)
            pdf_text = extract_text_from_pdf(pdf_data)

            if pdf_text.strip():
                pdf_texts[filename] = pdf_text
            else:
                print(f"⚠️ Aucun texte extrait de {filename}. Le fichier peut être scanné ou vide.")

    return pdf_texts

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


if __name__ == "__main__":
    msg_file = "tt/9.msg"  # Remplace par le chemin de ton fichier .msg
    output_folder = "tt"
    
    os.makedirs(output_folder, exist_ok=True)
    
    extracted_texts = extract_pdfs_from_msg(msg_file, output_folder)
    
    if extracted_texts:
        for pdf, text in extracted_texts.items():
            textfinal = f"{text}\n{'-'*50}"
            with open("extracted_text.txt", "w", encoding="utf-8") as file:
                file.write(textfinal)
            print(f"Texte extrait du PDF {pdf} sauvegardé dans 'extracted_text.txt'.")
    else:
        print("\n🚫 Aucun texte extrait. Vérifie si les fichiers sont bien extraits et lisibles.")
