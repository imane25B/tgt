import os
import io
import fitz  # PyMuPDF
import extract_msg
import re
import hashlib
from pathlib import Path
import shutil
import os
import unicodedata

def remove_accents(input_str):
    """
    Supprime les accents d'une chaîne de caractères en utilisant unicodedata.
    Args:
        input_str (str): Chaîne d'entrée.
    Returns:
        str: Chaîne sans accents.
    """
    return ''.join(
        c for c in unicodedata.normalize('NFD', str(input_str)) if unicodedata.category(c) != 'Mn'
    )

def save_extracted_data_to_txt(all_extracted_info, output_filename="extracted_data.txt"):
    """
    Sauvegarde les données extraites par les regex dans un fichier texte avec | comme délimiteur.
    Ajoute également le chemin du fichier PDF source.
    
    Args:
        all_extracted_info (dict): Dictionnaire où les clés sont les noms des fichiers PDF 
                                  et les valeurs sont les dictionnaires de données extraites.
        output_filename (str): Chemin du fichier de sortie.
    """
    # Définir l'ordre des colonnes à partir des clés possibles dans extract_information
    column_order = [
        "OBJET", "Mail_Expediteur", "Expediteur", "DATE HEURE ENVOI", "N PAGE", "Mail destinataire",
        "Entité", "Direction", "contact1AXA", "contact2AXA", "contact3AXA",
        "Destinataire", "Tel Destinataire", "Fax Destinataire", "Date Document",
        "Référence", "Compte à débiter", "SWIFT", "Titulaire de compte",
        "Montant décaissement", "Devise", "Date valeur compensée", "Bénéficiaire",
        "IBAN Bénéficiaire", "Banque Bénéficiaire", "Swift Bénéficiaire",
        "Motif du paiement", "Référence de l'opération", "Signataire1", "Signataire2", "PATH"
    ]
    
    # Vérifier si le fichier existe pour déterminer si l'en-tête doit être écrit
    file_exists = os.path.exists(output_filename)
    is_empty = not file_exists or os.stat(output_filename).st_size == 0
    
    with open(output_filename, 'a', encoding='utf-8') as f:
        # Écrire l'en-tête si le fichier est vide
        if is_empty:
            f.write("|".join(remove_accents(col).upper() for col in column_order) + "\n")
        
        # Parcourir les données extraites
        for pdf_path, data in all_extracted_info.items():
            # Ajouter le chemin du PDF aux données
            data["PATH"] = pdf_path
            
            # Générer une ligne de données en respectant l'ordre des colonnes
            line = []
            for col in column_order:
                value = data.get(col, "Non trouvé")
                # Nettoyer la valeur (supprimer accents et mettre en majuscules)
                cleaned_value = remove_accents(value).upper()
                line.append(cleaned_value)
            
            # Écrire la ligne dans le fichier
            f.write("|".join(line) + "\n")
    
    print(f"✅ Données sauvegardées avec succès dans {output_filename}")
# Set pour suivre les fichiers .msg déjà traités (éviter les boucles infinies)
processed_msg_files = set()

def sanitize_filename(filename):
    """Remove invalid characters from filename and ensure it's not too long."""
    # Remove invalid characters
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    
    # Limit length
    if len(filename) > 180:  # Leave some room for directory path
        base, ext = os.path.splitext(filename)
        filename = base[:176] + ext  # Truncate to fit
    
    return filename

def create_safe_path(base_dir, relative_path, filename):
    """Create a safe file path that won't exceed Windows path limitations."""
    # Create a hash for long paths to ensure uniqueness
    if len(relative_path) > 100:
        hash_obj = hashlib.md5(relative_path.encode())
        short_path = hash_obj.hexdigest()[:8]
        path = Path(base_dir) / short_path
    else:
        path = Path(base_dir) / relative_path
    
    # Create directory
    os.makedirs(path, exist_ok=True)
    
    # Return full path with safe filename
    return path / sanitize_filename(filename)

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
        
        # Extraction des informations du message
        objet = msg.subject if hasattr(msg, 'subject') else "Non trouvé"
        date = msg.date if hasattr(msg, 'date') else "Non trouvé"
        expediteur = msg.sender if hasattr(msg, 'sender') else "Non trouvé"
        mail_destinataire = msg.to if hasattr(msg, 'to') else "Non trouvé"
        
        # Extraction de l'adresse email de l'expéditeur
        mail_expediteur = "Non trouvé"
        if hasattr(msg, 'sender') and msg.sender:
            # Recherche d'un email dans le format "Nom <email@domaine.com>"
            email_match = re.search(r'<([^>]+)>', msg.sender)
            if email_match:
                mail_expediteur = email_match.group(1)
            else:
                # Si pas de format avec <>, recherche simple d'email
                email_match = re.search(r'[\w\.-]+@[\w\.-]+', msg.sender)
                if email_match:
                    mail_expediteur = email_match.group(0)
        
        print(f"📧 Informations du message :")
        print(f"  📌 Objet: {objet}")
        print(f"  📅 Date: {date}")
        print(f"  👤 Expéditeur: {expediteur}")
        print(f"  📩 Mail expéditeur: {mail_expediteur}")
        print(f"  👥 Destinataire: {mail_destinataire}")
        
    except Exception as e:
        print(f"❌ Erreur lors de l'ouverture de {msg_path} : {e}")
        return {}
    
    # Vérifier si des pièces jointes existent
    if not hasattr(msg, 'attachments') or not msg.attachments:
        print(f"Aucune pièce jointe trouvée dans {msg_path}.")
        return {}
    
    # Créer le sous-dossier pour les résultats basé sur le chemin du .msg
    base_msg_name = os.path.basename(msg_path)
    # Utiliser un hash pour les noms longs
    if len(base_msg_name) > 50:
        hash_obj = hashlib.md5(base_msg_name.encode())
        base_msg_name = hash_obj.hexdigest()[:10] + "_msg"
    
    msg_results_dir = os.path.join(results_dir, sanitize_filename(base_msg_name + "_results"))
    os.makedirs(msg_results_dir, exist_ok=True)
    
    for attachment in msg.attachments:
        if not attachment.longFilename:
            print("⚠️ Pièce jointe sans nom détectée. Ignorée.")
            continue
            
        filename = attachment.longFilename.rstrip('\x00').lower()
        safe_filename = sanitize_filename(filename)
        
        if filename.endswith('.pdf'):
            print(f"📄 PDF trouvé : {filename}")
            
            # Extraction du texte du PDF
            pdf_data = io.BytesIO(attachment.data)
            
            # Comptage du nombre de pages du PDF
            numero_pages = 0
            try:
                with fitz.open(stream=pdf_data, filetype="pdf") as pdf_document:
                    numero_pages = len(pdf_document)
                    pdf_data.seek(0)  # Réinitialiser le curseur pour la lecture suivante
            except Exception as e:
                print(f"⚠️ Erreur lors du comptage des pages PDF : {e}")
            
            pdf_text = extract_text_from_pdf(pdf_data)
            
            if pdf_text.strip():
                # Sauvegarder le texte extrait
                txt_output_path = os.path.join(output_dir, f"{safe_filename}_extracted_text.txt")
                try:
                    with open(txt_output_path, "w", encoding="utf-8") as text_file:
                        text_file.write(pdf_text)
                except (OSError, IOError) as e:
                    print(f"⚠️ Erreur lors de l'écriture du fichier texte: {e}")
                    # Sauvegarder dans un chemin plus court en cas d'erreur
                    alt_output_path = os.path.join(output_dir, f"{hashlib.md5(filename.encode()).hexdigest()[:10]}_text.txt")
                    with open(alt_output_path, "w", encoding="utf-8") as text_file:
                        text_file.write(pdf_text)
                    txt_output_path = alt_output_path
                
                # Appliquer les regex pour extraire des informations
                extracted_info = extract_information(pdf_text)
                
                # Ajouter les informations du message au dictionnaire des informations extraites
                extracted_info["OBJET"] = objet
                extracted_info["Mail_Expediteur"] = mail_expediteur
                extracted_info["Expediteur"] = expediteur
                extracted_info["DATE HEURE ENVOI"] = date
                extracted_info["N PAGE"] = str(numero_pages)
                extracted_info["Mail destinataire"] = mail_destinataire
                
                all_extracted_info[filename] = extracted_info
                
                # Sauvegarder les informations extraites
                info_output_path = os.path.join(msg_results_dir, f"{safe_filename}_extracted_info.txt")
                try:
                    with open(info_output_path, "w", encoding="utf-8") as info_file:
                        for key, value in extracted_info.items():
                            info_file.write(f"{key}: {value}\n")
                except (OSError, IOError) as e:
                    print(f"⚠️ Erreur lors de l'écriture du fichier d'informations: {e}")
                    # Sauvegarder dans un chemin plus court en cas d'erreur
                    alt_info_path = os.path.join(msg_results_dir, f"{hashlib.md5(filename.encode()).hexdigest()[:10]}_info.txt")
                    with open(alt_info_path, "w", encoding="utf-8") as info_file:
                        for key, value in extracted_info.items():
                            info_file.write(f"{key}: {value}\n")
                    info_output_path = alt_info_path
                
                print(f"✅ Traitement terminé pour {filename}. Informations extraites sauvegardées dans {info_output_path}")
            else:
                print(f"⚠️ Aucun texte extrait de {filename}. Le fichier peut être scanné ou vide.")
        
        elif filename.endswith('.msg'):
            print(f"📧 Fichier .msg imbriqué trouvé : {filename}")
            
            # Sauvegarder le fichier .msg imbriqué
            nested_msg_path = os.path.join(output_dir, safe_filename)
            try:
                with open(nested_msg_path, "wb") as nested_msg_file:
                    nested_msg_file.write(attachment.data)
            except (OSError, IOError) as e:
                print(f"⚠️ Erreur lors de l'écriture du fichier .msg imbriqué: {e}")
                # Sauvegarder dans un chemin plus court en cas d'erreur
                alt_msg_path = os.path.join(output_dir, f"{hashlib.md5(filename.encode()).hexdigest()[:10]}.msg")
                with open(alt_msg_path, "wb") as nested_msg_file:
                    nested_msg_file.write(attachment.data)
                nested_msg_path = alt_msg_path
            
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
    # Convertir en objets Path pour une meilleure gestion des chemins
    root_folder_path = Path(root_folder)
    output_folder_path = Path(output_folder)
    results_folder_path = Path(results_folder)
    
    # Créer les dossiers de sortie s'ils n'existent pas
    output_folder_path.mkdir(exist_ok=True, parents=True)
    results_folder_path.mkdir(exist_ok=True, parents=True)
    
    # Statistiques pour le résumé final
    total_msg_files = 0
    total_pdf_files = 0
    total_nested_msg = 0
    
    # Dictionnaire pour stocker toutes les informations extraites
    all_pdf_data = {}
    
    for dirpath, _, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.lower().endswith(".msg"):
                total_msg_files += 1
                msg_file_path = os.path.join(dirpath, filename)
                print(f"\n📂 Traitement de {msg_file_path}...")
                
                # Créer un sous-dossier dans output_folder basé sur le chemin relatif
                try:
                    relative_path = os.path.relpath(dirpath, root_folder)
                    # Limiter la profondeur du chemin relatif pour éviter des chemins trop longs
                    path_parts = Path(relative_path).parts
                    if len(path_parts) > 3:  # Limiter à 3 niveaux de dossiers
                        short_path = os.path.join(*path_parts[-3:])
                    else:
                        short_path = relative_path
                    
                    output_subfolder = output_folder_path / short_path
                    output_subfolder.mkdir(exist_ok=True, parents=True)
                except (OSError, IOError) as e:
                    print(f"⚠️ Erreur lors de la création du sous-dossier: {e}")
                    # Utiliser un dossier basé sur un hash en cas d'erreur
                    hash_obj = hashlib.md5(os.path.dirname(msg_file_path).encode())
                    output_subfolder = output_folder_path / hash_obj.hexdigest()[:8]
                    output_subfolder.mkdir(exist_ok=True, parents=True)
                
                # Extraire et traiter les PDF du fichier .msg
                try:
                    extracted_info = extract_and_process_pdfs_from_msg(msg_file_path, str(output_subfolder), str(results_folder_path))

                    # Ajouter les informations extraites au dictionnaire global
                    for pdf_name, info in extracted_info.items():
                        full_path = f"{msg_file_path}>{pdf_name}" if '>' in pdf_name else pdf_name
                        all_pdf_data[full_path] = info
                
                except Exception as e:
                    print(f"❌ Erreur critique lors du traitement de {msg_file_path}: {e}")
                    continue
                
                # Mettre à jour les statistiques
                pdf_count = sum(1 for key in extracted_info.keys() if not '>' in key and '.pdf' in key)
                nested_msg_count = sum(1 for key in extracted_info.keys() if '>' in key)
                
                total_pdf_files += pdf_count
                total_nested_msg += nested_msg_count
                
                # Créer un fichier de synthèse pour ce .msg
                safe_filename = sanitize_filename(filename)
                summary_file = results_folder_path / f"{safe_filename}_summary.txt"
                try:
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
                except (OSError, IOError) as e:
                    print(f"⚠️ Erreur lors de l'écriture du fichier de synthèse: {e}")

    # Sauvegarder toutes les données extraites dans un fichier texte
    consolidated_data_path = os.path.join(results_folder, "donnees_extraites_consolidees.txt")
    save_extracted_data_to_txt(all_pdf_data, consolidated_data_path)
    # Créer un rapport global
    global_report_path = results_folder_path / "rapport_global.txt"
    try:
        with open(global_report_path, "w", encoding="utf-8") as report:
            report.write(f"Rapport global d'extraction\n")
            report.write(f"==========================\n\n")
            report.write(f"Total de fichiers .msg traités: {total_msg_files}\n")
            report.write(f"Total de fichiers PDF extraits: {total_pdf_files}\n")
            report.write(f"Total de fichiers .msg imbriqués: {total_nested_msg}\n")
        
        print(f"\n✅ Traitement terminé!")
        print(f"Rapport global disponible à: {global_report_path}")
    except (OSError, IOError) as e:
        print(f"⚠️ Erreur lors de l'écriture du rapport global: {e}")
        print(f"\n✅ Traitement terminé, mais impossible d'écrire le rapport global!")

# Fonction pour vérifier si le dossier est accessible et s'il contient des fichiers .msg
def validate_input_folder(folder_path):
    if not os.path.exists(folder_path):
        print(f"❌ Le dossier {folder_path} n'existe pas!")
        return False
    
    msg_files = []
    for dirpath, _, filenames in os.walk(folder_path):
        for filename in filenames:
            if filename.lower().endswith(".msg"):
                msg_files.append(os.path.join(dirpath, filename))
                if len(msg_files) >= 5:  # Vérifier seulement les 5 premiers pour éviter de parcourir tout le dossier
                    break
        if len(msg_files) >= 5:
            break
    
    if not msg_files:
        print(f"⚠️ Aucun fichier .msg trouvé dans {folder_path}. Vérifiez le dossier!")
        return False
    
    return True

if __name__ == "__main__":
    # Paramètres configurables
    root_folder = "Virements vers 23 mails_2 ans/"  # Dossier racine contenant les fichiers .msg
    output_folder = "extracted_files"  # Dossier pour stocker les fichiers extraits et le texte brut
    results_folder = "resultats_extraction"  # Dossier pour stocker les informations extraites
    
    # Vérifier si les dossiers de sortie existent déjà et les nettoyer si nécessaire
    for folder in [output_folder, results_folder]:
        if os.path.exists(folder):
            try:
                shutil.rmtree(folder)
                print(f"🧹 Nettoyage du dossier existant: {folder}")
            except Exception as e:
                print(f"⚠️ Impossible de nettoyer le dossier {folder}: {e}")
                # Créer un nom alternatif avec timestamp
                import time
                folder = f"{folder}_{int(time.time())}"
                print(f"Utilisation d'un dossier alternatif: {folder}")
    
    # Valider le dossier d'entrée
    if validate_input_folder(root_folder):
        try:
            process_msg_files_recursively(root_folder, output_folder, results_folder)
        except Exception as e:
            print(f"❌ Erreur critique: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("⛔ Traitement annulé en raison d'erreurs dans la configuration.")

if __name__ == "__main__":
    root_folder = "Virements vers 23 mails_2 ans/"  # Dossier racine contenant les fichiers .msg
    output_folder = "extracted_files"  # Dossier pour stocker les fichiers extraits et le texte brut
    results_folder = "resultats_extraction"  # Dossier pour stocker les informations extraites
    
    process_msg_files_recursively(root_folder, output_folder, results_folder)
