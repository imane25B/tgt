import os
import re

def extract_information(text):
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
        "Motif du paiement": [r"Référence à indiquer sur le\s+virement -Détail Réf de\s+l’opération\s*/\s*Transfer\s+reference\s*([A-Z0-9]+)",r"Motif du paiement\s*/\s*Payment purpose\s*/\s*Transfer reference\s*([^\s]+)",r"Détail Réf de l’opération\s*/\s*Transfer reference\s*(.*)",r"Détail Réf de l’opération\s*/\s*Transfer\s*reference\s*([^\s]+)"],
        "Référence de l'opération": [r"Détail Réf de l’opération\s*/\s*Transfer reference[\s\S]*?(Transfer id\s*\d+\s.*)",r"Transfer id[^\n]*"],
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
                    extracted_data[key] = match.group(1)
                except IndexError:
                    extracted_data[key] = match.group(0)
                break  # Arrête de tester une fois qu'une regex a fonctionné

    return extracted_data

def process_text_files_in_folder(root_folder, output_folder):
    """
    Parcourt récursivement le dossier racine pour traiter tous les fichiers .txt.
    """
    os.makedirs(output_folder, exist_ok=True)  # Créer le dossier de sortie s'il n'existe pas

    for dirpath, _, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.endswith(".txt"):
                input_file_path = os.path.join(dirpath, filename)
                print(f"📂 Traitement de {input_file_path}...")
                
                # Lire le contenu du fichier texte
                with open(input_file_path, "r", encoding="utf-8") as file:
                    text = file.read()
                
                # Extraire les informations
                extracted_info = extract_information(text)
                
                # Chemin pour le fichier de sortie
                relative_path = os.path.relpath(dirpath, root_folder)
                output_subfolder = os.path.join(output_folder, relative_path)
                os.makedirs(output_subfolder, exist_ok=True)
                
                output_file_path = os.path.join(output_subfolder, f"{filename}_extracted_info.txt")
                with open(output_file_path, "w", encoding="utf-8") as file:
                    for key, value in extracted_info.items():
                        file.write(f"{key}: {value}\n")
                
                print(f"✅ Informations extraites sauvegardées dans {output_file_path}.")

if __name__ == "__main__":
    root_folder = "tt"  # Remplacez par le chemin de votre dossier contenant les fichiers .txt
    output_folder = "resultats_extraction"  # Dossier où les résultats seront stockés
    
    process_text_files_in_folder(root_folder, output_folder)
