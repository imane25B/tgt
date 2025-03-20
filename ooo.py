import os
import io
import fitz  # PyMuPDF
import extract_msg
import re

# Set pour suivre les fichiers .msg déjà traités (éviter les boucles infinies)
processed_msg_files = set()

def extract_information(text):
    """
    Applique les regex pour extraire les informations importantes du texte PDF.
    """
    extracted_data = {}

    # Dictionnaire de regex : chaque clé correspond à une liste de regex
    patterns = {
        "Entité": [r"4\s*44\s*=\s*([A-Z\s]+)\n[A-Z\s]+\n",r"AXA\s+[A-Za-zÉÈÊÛÔÎÏ ]+"],
        "Direction": [r"DIRECTION FINANCIERE\s+SERVICE TRESORERIE",r"Direction\s+[^\n]*"],
        "contact1AXA": [r"Direction Financière Service Trésorerie\s*([\w\s]+Denis\s+\d{2}\s\d{2}\s\d{2}\s\d{2}\s\d{2})",r"Direction Financière – Service Trésorerie\s*\n\s*([^\n]+\d{2}\s\d{2}\s\d{2}\s\d{2}\s\d{2})"],
        "contact2AXA":[r"Direction Financière Service Trésorerie\s*.*?\s*(HAMON Pascal\s+\d{2}\s\d{2}\s\d{2}\s\d{2}\s\d{2})",r"Direction Financière – Service Trésorerie\s*\n\s*[^\n]+\d{2}\s\d{2}\s\d{2}\s\d{2}\s\d{2}\s*\n\s*([^\n]+\d{2}\s\d{2}\s\d{2}\s\d{2}\s\d{2})"],
        "contact3AXA":[r"Direction Financière Service Trésorerie[\s\S]*?(VUONG THI Thien\s+\d{2}\s\d{2}\s\d{2}\s\d{2}\s\d{2})",r"Direction Financière – Service Trésorerie\s*(?:\n\s*[^\n]+){2}\n\s*([^\n]+\d{2}\s\d{2}\s\d{2}\s\d{2}\s\d{2})"],
        "Destinataire": [r"Mail\s*:\s*[^\n]*\n\s*([^\n]+)"],
        "Tel Destinataire": [r"Tel\s*:\s*(\d{2}\s\d{2}\s\d{2}\s\d{2}\s\d{2})"],
        "Fax Destinataire": [r"Fax\s*:\s*(\d{2}\s\d{2}\s\d{2}\s\d{2}\s\d{2})"],
        "Date Document": [r"(?:\ble\b|\bon\b)\s*(\d{2}/\d{2}/\d{4})"],
        "Référence": [r"Notre référence\s*/\s*Our reference:\s*(\d+)"],
        "Compte à débiter": [r"Par le débit de notre\s+compte n°\s*/\s*From\s+our bank\s+account number\s+([A-Z]{2}\d{2}(?:\s?[A-Z0-9]{4}){5}\s?[A-Z0-9]{3})\s+Swift:",r"Par le débit de notre compte n°\s*/\s*From\s+our bank account number\s+([A-Z]{2}\d{2}(?:\s?[A-Z0-9]{4}){5}\s?[A-Z0-9]{3})\s+Swift:",r"Par le débit de notre compte n°\s*/\s*From\s+our bank account number\s*([A-Z]{2}\d{2}(?:\s\d{4}){5}\s\d{3})",r"Par le débit de notre\s*compte n°\s*/\s*From our\s*bank account number\s*([A-Z]{2}\d{2}(?:\s\d{4}){4})",r"Par le débit de notre\s*compte n°\s*/\s*From our bank\s*account number\s*([A-Z]{2}\d{2}(?:\s\d{4}){4})",r"Par le débit de notre compte\s*n°\s*/\s*From our bank account\s*number\s*([A-Z]{2}\d{2}(?:\s\d{4}){4})",r"Par le débit de notre compte n°\s*/\s*From\s*our bank account number\s*([A-Z]{2}\d{2}(?:\s\d{4}){4})",r"Par le débit de notre compte n°\s*/\s*From our bank account\s*number\s*([A-Z]{2}\d{2}(?:\s\d{4}){4})",r"Par le débit de notre\s*compte n°\s*/\s*From our bank\s*account number\s*([A-Z]{2}\d{2}(?:\s\d{4}){5})",r"bank account number\s*(\w{2}\d{2}\s\d{4}\s\d{4}\s\d{4}\s\d{4}\s\d{4}\s\d{3})",r"Par le débit de notre compte n°\s*/\s*From our bank account\s*number\s*([A-Z]{2}\d{2}(?:\s\d{4}){5})",r"Par le débit de notre compte\s*n°\s*/\s*From our bank account\s*number\s*([A-Z]{2}\d{2}(?:\s\d{4}){5})"],
        "SWIFT": [r"Swift:\s*([A-Z0-9]+)"],
        "Titulaire de compte": [r"Swift:\s*[A-Z0-9]+\s*(.*)"],
        "Montant décaissement": [r"Veuillez virer la somme\s*de\s*/\s*Please transfer the\s*amount of\s*([\d,]+\.\d{2})",r"Veuillez virer la somme de\s*/\s*Please\s*transfer the amount of\s*\n\s*([\d,]+\.\d{2})",r"Veuillez virer la somme de\s*/\s*Please transfer the\s*amount of\s*([\d.,]+)",r"transfer the amount of\s*\n\s*(\d{1,3}(?:\s\d{3})*,\d{2})",r"Veuillez virer la somme de\s*/\s*Please transfer the amount\s*of\s*\n\s*([\d,]+\.\d{2})"],
        "Devise": [r"Veuillez virer la somme de\s*/\s*Please\s*transfer the amount of\s*[\d\s,.]+\s([A-Z]{3})",r"Veuillez virer la somme\s*de\s*/\s*Please transfer the\s*amount of\s*[\d,]+\.\d{2}\s([A-Z]{3})",r"Veuillez virer la somme de\s*/\s*Please\s*transfer the amount of\s*\n\s*[\d,]+\.\d{2}\s([A-Z]+)",r"Veuillez virer la somme de\s*/\s*Please transfer the amount\s*of\s*\n\s*[^\d\n]*[\d,]+\.[\d]{2}\s([^\s]+)"], 
        "Date valeur compensée": [r"Date de valeur\s*compensée\s*/\s*Compensated value\s*date\s*(\d{2}/\d{2}/\d{4})",r"Date de valeur\s*compensée\s*/\s*Compensated value date\s*([\d/]+)"],
        "Bénéficiaire": [r"Nom bénéficiaire\s*/\s*Beneficiary name\s*IBAN\s*/\s*IBAN\s*(.*)",r"IBAN / IBAN\s*\n\s*([^\n]+)"],
        "IBAN Bénéficiaire": [r"IBAN\s*/\s*IBAN[\s\S]*?AXA FRANCE VIE[\s\S]*?HO[\s\S]*?(FR\d{2}(?:\s?\d{4}){5}\s?[A-Z0-9]{3})[\s\S]*?Banque bénéficiaire\s*/\s*Beneficiary[\s\S]*?bank[\s\S]*?Code Swift\s*/\s*Swift code",r"IBAN\s*/\s*IBAN\s+AXA FRANCE VIE\s+HO\s+([A-Z]{2}\d{2}(?:\s\d{4}){5}\s[A-Z0-9]{3})\s+Banque bénéficiaire\s*/\s*Beneficiary bank\s+Code Swift\s*/\s*Swift code",r"IBAN\s*/\s*IBAN[\s\S]*?([A-Z]{2}\d{2}[A-Z0-9]+)",r"Nom bénéficiaire\s*/\s*Beneficiary name\s*IBAN\s*/\s*IBAN[\s\S]*?(FR\d{2}\s\d{4}\s\d{4}\s\d{4}\s\d{4}\s\d{3})",r"IBAN / IBAN\s+[^\n]*\n[^\n]*\n([A-Z]{2}\d{2}(?:\s\d{4}){5})"],
        "Banque Bénéficiaire": [r"Banque bénéficiaire\s*/\s*Beneficiary bank\s*Code Swift\s*/\s*Swift code\s*(\w+)",r"Banque bénéficiaire\s*/\s*Beneficiary\s*bank\s*Code Swift\s*/\s*Swift code\s*([A-Z]{4})"],
        "Swift Bénéficiaire": [r"Banque bénéficiaire\s*/\s*Beneficiary bank\s*Code Swift\s*/\s*Swift code[\s\S]*?([A-Z]{4}[A-Z0-9]{3,})",r"Code Swift\s*/\s*Swift code\s*(?:\n\s*[^\n]*){1,2}\s*([A-Z]{8}[A-Z0-9]{3})"],
        "Motif du paiement": [r"Référence à indiquer sur le\s+virement -Détail Réf de\s+l'opération\s*/\s*Transfer\s+reference\s*([A-Z0-9]+)",r"Motif du paiement\s*/\s*Payment purpose\s*/\s*Transfer reference\s*([^\s]+)",r"Détail Réf de l'opération\s*/\s*Transfer reference\s*(.*)",r"Détail Réf de l'opération\s*/\s*Transfer\s*reference\s*([^\s]+)"],
        "Référence de l'opération": [r"Détail Réf de l'opération\s*/\s*Transfer reference[\s\S]*?(Transfer id\s*\d+\s.*)",r"Transfer id[^\n]*"],
        "Signataire1": [r"Signatures autorisées\s*/\s*Authorized signatures[\s\S]*?(\b[A-Z]+\s[A-Z]+\s[A-Za-z]+)",r"Signatures autorisées\s*/\s*Authorized signatures\s*\n\s*([^\n]+)"],
        "Signataire2": [r"Signatures autorisées\s*/\s*Authorized signatures[\s\S]*\s([A-Z]+\s[A-Za-z]+(?:\s[A-Za-z]+)*)\s*$",r"Signatures autorisées\s*/\s*Authorized signatures[\s\S]*?\n\s*([A-Z]+\s[A-Z]+\s[A-Za-z]+)\s*\n\s*([A-Z]+\s[A-Z]+\s[A-Za-z]+)",r"Signatures autorisées / Authorized signatures\s*(?:.*\n){1}\s*(.*)"]
    }

    # Appliquer les regex pour chaque clé
    for key, regex_list in patterns.items():
        extracted_data[key] = "Non trouvé"  # Valeur par défaut si aucune regex ne fonctionne
        for pattern in regex_list:
            match = re.search(pattern, text, re.MULTILINE)
            if match:
                try:
                    extracted_data[key] = match.group(1).strip()
                except IndexError:
                    extracted_data[key] = match.group(0).strip()
                break  # Arrête de tester une fois qu'une regex a fonctionné

    return extracted_data

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

def extract_and_process_pdfs_from_msg(msg_path, output_dir, results_dir):
    """
    Extrait les fichiers PDF d'un fichier .msg, applique les regex et gère les fichiers imbriqués.
    Retourne un dictionnaire contenant les informations extraites de chaque PDF.
    """
    if msg_path in processed_msg_files:
        print(f"⚠️ Fichier déjà traité : {msg_path}. Ignoré pour éviter les boucles.")
        return {}
    
    processed_msg_files.add(msg_path)  # Marquer le fichier comme traité
    
    # Dictionnaire pour stocker les informations extraites par PDF
    all_extracted_info = {}
    
    try:
        msg = extract_msg.Message(msg_path)
    except Exception as e:
        print(f"❌ Erreur lors de l'ouverture de {msg_path} : {e}")
        return {}
    
    # Vérifier si des pièces jointes existent
    if not hasattr(msg, 'attachments') or not msg.attachments:
        print(f"Aucune pièce jointe trouvée dans {msg_path}.")
        return {}
    
    # Créer le sous-dossier pour les résultats basé sur le chemin du .msg
    base_msg_name = os.path.basename(msg_path)
    msg_results_dir = os.path.join(results_dir, base_msg_name + "_results")
    os.makedirs(msg_results_dir, exist_ok=True)
    
    for attachment in msg.attachments:
        if not attachment.longFilename:
            print("⚠️ Pièce jointe sans nom détectée. Ignorée.")
            continue
            
        filename = attachment.longFilename.rstrip('\x00').lower()
        
        if filename.endswith('.pdf'):
            print(f"📄 PDF trouvé : {filename}")
            
            # Extraction du texte du PDF
            pdf_data = io.BytesIO(attachment.data)
            pdf_text = extract_text_from_pdf(pdf_data)
            
            if pdf_text.strip():
                # Sauvegarder le texte extrait
                txt_output_path = os.path.join(output_dir, f"{filename}_extracted_text.txt")
                with open(txt_output_path, "w", encoding="utf-8") as text_file:
                    text_file.write(pdf_text)
                
                # Appliquer les regex pour extraire des informations
                extracted_info = extract_information(pdf_text)
                all_extracted_info[filename] = extracted_info
                
                # Sauvegarder les informations extraites
                info_output_path = os.path.join(msg_results_dir, f"{filename}_extracted_info.txt")
                with open(info_output_path, "w", encoding="utf-8") as info_file:
                    for key, value in extracted_info.items():
                        info_file.write(f"{key}: {value}\n")
                
                print(f"✅ Traitement terminé pour {filename}. Informations extraites sauvegardées dans {info_output_path}")
            else:
                print(f"⚠️ Aucun texte extrait de {filename}. Le fichier peut être scanné ou vide.")
        
        elif filename.endswith('.msg'):
            print(f"📧 Fichier .msg imbriqué trouvé : {filename}")
            
            # Sauvegarder le fichier .msg imbriqué
            nested_msg_path = os.path.join(output_dir, filename)
            with open(nested_msg_path, "wb") as nested_msg_file:
                nested_msg_file.write(attachment.data)
            
            # Traiter récursivement le fichier .msg imbriqué
            nested_results = extract_and_process_pdfs_from_msg(nested_msg_path, output_dir, results_dir)
            
            # Ajouter les résultats du .msg imbriqué aux résultats globaux
            for pdf_name, info in nested_results.items():
                all_extracted_info[f"{filename}>{pdf_name}"] = info  # Utiliser une notation pour indiquer l'imbrication
    
    return all_extracted_info

def process_msg_files_recursively(root_folder, output_folder, results_folder):
    """
    Parcourt récursivement un dossier racine pour traiter tous les fichiers .msg,
    extraire les PDF et appliquer les regex.
    """
    # Créer les dossiers de sortie s'ils n'existent pas
    os.makedirs(output_folder, exist_ok=True)
    os.makedirs(results_folder, exist_ok=True)
    
    # Statistiques pour le résumé final
    total_msg_files = 0
    total_pdf_files = 0
    total_nested_msg = 0
    
    for dirpath, _, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.endswith(".msg"):
                total_msg_files += 1
                msg_file_path = os.path.join(dirpath, filename)
                print(f"\n📂 Traitement de {msg_file_path}...")
                
                # Créer un sous-dossier dans output_folder basé sur le chemin relatif
                relative_path = os.path.relpath(dirpath, root_folder)
                output_subfolder = os.path.join(output_folder, relative_path)
                os.makedirs(output_subfolder, exist_ok=True)
                
                # Extraire et traiter les PDF du fichier .msg
                extracted_info = extract_and_process_pdfs_from_msg(msg_file_path, output_subfolder, results_folder)
                
                # Mettre à jour les statistiques
                pdf_count = sum(1 for key in extracted_info.keys() if not '>' in key and '.pdf' in key)
                nested_msg_count = sum(1 for key in extracted_info.keys() if '>' in key)
                
                total_pdf_files += pdf_count
                total_nested_msg += nested_msg_count
                
                # Créer un fichier de synthèse pour ce .msg
                summary_file = os.path.join(results_folder, f"{filename}_summary.txt")
                with open(summary_file, "w", encoding="utf-8") as summary:
                    summary.write(f"Résumé du traitement pour {msg_file_path}\n")
                    summary.write(f"Nombre de PDF extraits: {pdf_count}\n")
                    summary.write(f"Nombre de .msg imbriqués: {nested_msg_count}\n\n")
                    
                    if extracted_info:
                        summary.write("Liste des fichiers traités avec informations clés:\n")
                        for pdf_name, info in extracted_info.items():
                            summary.write(f"\n--- {pdf_name} ---\n")
                            
                            # Extraire et afficher quelques informations importantes
                            key_info = {
                                "Montant": info.get("Montant décaissement", "Non trouvé"),
                                "Devise": info.get("Devise", "Non trouvé"),
                                "Bénéficiaire": info.get("Bénéficiaire", "Non trouvé"),
                                "IBAN": info.get("IBAN Bénéficiaire", "Non trouvé"),
                                "Date": info.get("Date Document", "Non trouvé"),
                                "Référence": info.get("Référence", "Non trouvé")
                            }
                            
                            for key, value in key_info.items():
                                summary.write(f"{key}: {value}\n")
                    else:
                        summary.write("Aucune information extraite.\n")
    
    # Créer un rapport global
    global_report_path = os.path.join(results_folder, "rapport_global.txt")
    with open(global_report_path, "w", encoding="utf-8") as report:
        report.write(f"Rapport global d'extraction\n")
        report.write(f"==========================\n\n")
        report.write(f"Total de fichiers .msg traités: {total_msg_files}\n")
        report.write(f"Total de fichiers PDF extraits: {total_pdf_files}\n")
        report.write(f"Total de fichiers .msg imbriqués: {total_nested_msg}\n")
    
    print(f"\n✅ Traitement terminé!")
    print(f"Rapport global disponible à: {global_report_path}")

if __name__ == "__main__":
    root_folder = "Virements vers 23 mails_2 ans/"  # Dossier racine contenant les fichiers .msg
    output_folder = "extracted_files"  # Dossier pour stocker les fichiers extraits et le texte brut
    results_folder = "resultats_extraction"  # Dossier pour stocker les informations extraites
    
    process_msg_files_recursively(root_folder, output_folder, results_folder)
