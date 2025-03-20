import os
from PDF_extraction import extract_pdfs_from_msg

def process_msg_files_recursively(root_folder, output_folder):
    """
    Parcourt récursivement un dossier racine pour traiter tous les fichiers .msg et extrait les PDF.
    """
    # Créer le dossier de sortie s'il n'existe pas
    os.makedirs(output_folder, exist_ok=True)

    for dirpath, _, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.endswith(".msg"):
                msg_file_path = os.path.join(dirpath, filename)
                print(f"📂 Traitement de {msg_file_path}...")
                
                # Extraire les PDF du fichier .msg
                extracted_texts = extract_pdfs_from_msg(msg_file_path, output_folder)
                
                if extracted_texts:
                    # Sauvegarder les textes extraits dans un fichier texte séparé pour chaque .msg
                    relative_path = os.path.relpath(dirpath, root_folder)  # Chemin relatif pour organiser la sortie
                    output_subfolder = os.path.join(output_folder, relative_path)
                    os.makedirs(output_subfolder, exist_ok=True)

                    output_text_file = os.path.join(output_subfolder, f"{filename}_extracted_texts.txt")
                    with open(output_text_file, "w", encoding="utf-8") as file:
                        for pdf, text in extracted_texts.items():
                            textfinal = f"--- Texte extrait du PDF : {pdf} ---\n{text}\n{'-'*50}\n"
                            file.write(textfinal)
                    print(f"✅ Texte extrait sauvegardé dans {output_text_file}.")
                else:
                    print(f"🚫 Aucun texte extrait pour {msg_file_path}. Vérifiez le fichier.")

if __name__ == "__main__":
    root_folder = "Virements vers 23 mails_2 ans/"  # Remplacez par le chemin de votre dossier racine
    output_folder = "tt"  # Dossier où les résultats seront stockés
    
    process_msg_files_recursively(root_folder, output_folder)
