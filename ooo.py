import os
import io
import fitz  # PyMuPDF
import extract_msg
import re
import hashlib
from pathlib import Path
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

def save_extracted_data_to_txt(all_extracted_info, output_filename="donnees_extraites_consolidees.txt"):
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
        "Entité", "Direction", "contact1AXA", "contact2AXA", "contact3AXA",
        "Destinataire", "Tel Destinataire", "Fax Destinataire", "Date Document",
        "Référence", "Compte à débiter", "SWIFT", "Titulaire de compte",
        "Montant décaissement", "Devise", "Date valeur compensée", "Bénéficiaire",
        "IBAN Bénéficiaire", "Banque Bénéficiaire", "Swift Bénéficiaire",
        "Motif du paiement", "Référence de l'opération", "Signataire1", "Signataire2", "PATH"
    ]
    
    # Créer le répertoire parent du fichier de sortie s'il n'existe pas
    os.makedirs(os.path.dirname(output_filename), exist_ok=True)
    
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

def extract_and_process_pdfs_from_msg(msg_path):
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
                # Appliquer les regex pour extraire des informations
                extracted_info = extract_information(pdf_text)
                all_extracted_info[filename] = extracted_info
            else:
                print(f"⚠️ Aucun texte extrait de {filename}. Le fichier peut être scanné ou vide.")
        
        elif filename.endswith('.msg'):
            print(f"📧 Fichier .msg imbriqué trouvé : {filename}")
            
            # Sauvegarder temporairement le fichier .msg imbriqué
            temp_msg_path = os.path.join("temp", filename)
            os.makedirs(os.path.dirname(temp_msg_path), exist_ok=True)
            
            try:
                with open(temp_msg_path, "wb") as nested_msg_file:
                    nested_msg_file.write(attachment.data)
            except (OSError, IOError) as e:
                print(f"⚠️ Erreur lors de l'écriture du fichier .msg imbriqué: {e}")
                continue
            
            # Traiter récursivement le fichier .msg imbriqué
            nested_results = extract_and_process_pdfs_from_msg(temp_msg_path)
            
            # Nettoyer le fichier temporaire après traitement
            try:
                os.remove(temp_msg_path)
            except:
                pass
            
            # Ajouter les résultats du .msg imbriqué aux résultats globaux
            for pdf_name, info in nested_results.items():
                all_extracted_info[f"{filename}>{pdf_name}"] = info  # Utiliser une notation pour indiquer l'imbrication
    
    return all_extracted_info

def process_msg_files_recursively(root_folder, results_filename):
    """
    Parcourt récursivement un dossier racine pour traiter tous les fichiers .msg,
    extraire les PDF et appliquer les regex.
    """
    # Créer le dossier temp pour les fichiers .msg imbriqués temporaires
    os.makedirs("temp", exist_ok=True)
    
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
                
                # Extraire et traiter les PDF du fichier .msg
                try:
                    extracted_info = extract_and_process_pdfs_from_msg(msg_file_path)

                    # Ajouter les informations extraites au dictionnaire global
                    for pdf_name, info in extracted_info.items():
                        full_path = f"{msg_file_path}>{pdf_name}" if '>' in pdf_name else f"{msg_file_path}>{pdf_name}"
                        all_pdf_data[full_path] = info
                
                except Exception as e:
                    print(f"❌ Erreur critique lors du traitement de {msg_file_path}: {e}")
                    continue
                
                # Mettre à jour les statistiques
                pdf_count = sum(1 for key in extracted_info.keys() if '.pdf' in key.lower())
                nested_msg_count = sum(1 for key in extracted_info.keys() if '>' in key)
                
                total_pdf_files += pdf_count
                total_nested_msg += nested_msg_count

    # Sauvegarder toutes les données extraites dans un fichier texte
    save_extracted_data_to_txt(all_pdf_data, results_filename)
    
    # Nettoyer le dossier temp
    try:
        import shutil
        shutil.rmtree("temp")
    except:
        pass
    
    print(f"\n✅ Traitement terminé!")
    print(f"Total de fichiers .msg traités: {total_msg_files}")
    print(f"Total de fichiers PDF extraits: {total_pdf_files}")
    print(f"Total de fichiers .msg imbriqués: {total_nested_msg}")
    print(f"Résultats sauvegardés dans: {results_filename}")

if __name__ == "__main__":
    # Paramètres configurables
    root_folder = "Virements vers 23 mails_2 ans/"  # Dossier racine contenant les fichiers .msg
    results_filename = "donnees_extraites_consolidees.txt"  # Nom du fichier de résultats
    
    # Exécuter le traitement
    process_msg_files_recursively(root_folder, results_filename)
