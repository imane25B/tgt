import re

def extract_information(text):
    extracted_data = {}

    # Dictionnaire de regex basées sur les titres
    patterns = {
        "Entité": r"4\s*44\s*=\s*([A-Z\s]+)\n[A-Z\s]+\n",
        "Direction": r"Direction\s+[^\n]*",  # Mise à jour pour récupérer toute la ligne Direction
        "contact1AXA": r"([A-Za-zÀ-ÿ]+(?: [A-Za-zÀ-ÿ]+)*)\s*((?:\d{2}\s*){4,5})",
        "Destinataire": r"(?:4\s*44\s*=)?\s*(?:[A-Z\s]+\n)*([A-Z\s]+)\n",
        "Tel Destinataire": r"Tel\s*\n([\d\s]+)",
        "Fax Destinataire": r"Fax\s*\n([\d\s]+)",
        "Mails": r"Mail\s*([\w\.-]+(?:\s[\w\.-]*)*@[\w\.-]+)",
        "Date Document": r"(.*?le\s*/?\s*on\s*\d{2}/\d{2}/\d{4})(?=\s*Notre\s+référence)",
        "Référence": r"Our\s+reference\s*:?\s*(\d+)",
        "Compte à débiter": r"(?:From\s*\n|Par le débit de notre compte n)\s*([A-Z]{2}\d{2}\s*\d{4}\s*\d{4}\s*\d{4}\s*\d{4}\s*\d{4}\s*\d{3})",
        "SWIFT": r"Swift:\s*([A-Z0-9]+)",
        "Titulaire de compte": r"Swift:\s*[A-Z0-9]+\s*(.*\n.*)",
        "Montant décaissement": r"Veuillez virer la somme de\s*(?:Please\s*)?([\d,]+\.\d{2})\s([A-Z]{3})",
        "Date valeur compensée": r"Date de valeur compensée\s*(\d{2}/\d{2}/\d{4})",
        "Bénéficiaire": r"Nom bénéficiaire\s+Beneficiary(?:\s+name)?\s+([A-ZÀ-ÿ\s]+?)(?=\s+(?:name|IBAN))",
        "IBAN Bénéficiaire": r"IBAN\s*IBAN\s*([A-Z0-9\s]+)\s*Banque bénéficiaire",
        "IBAN Bénéficiaire 2": r"([^\n]*)(?=\nIBAN)",
        "Banque Bénéficiaire": r"Banque bénéficiaire\s*Beneficiary\s*(?:bank)?\s*([A-Z]+)",
        "Swift Bénéficiaire": r"Code Swift\s*Swift code\s*([A-Z0-9]+)",
        "Swift Bénéficiaire 2": r"([A-Z0-9]+)\s*(?=.*Code\s*Swift\s*Swift\s*code)",
        "Commission": r"([^\n]+)(?=\nWithout account commission)",
        "Motif du paiement": r"Motif du paiement\s*Payment\s*(?:purpose)?\s*([^\n]+)",
        "Référence de l'opération": r"(id \d+)",
        "Signataire1": r"Signatures autorisées\s*Authorized signatures\s*([A-Za-zÀ-ÿ\s]+)",
        "Signataire2": r"Signatures autorisées\s*Authorized signatures\s*([A-Za-zÀ-ÿ\s]+)"
    }

    # Appliquer les regex et stocker les résultats
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.MULTILINE)
        if match:
            extracted_data[key] = match.group(1) if match.lastgroup else match.group(0)
        else:
            extracted_data[key] = "Non trouvé"

    return extracted_data


if __name__ == "__main__":
    with open("extracted_text.txt", "r", encoding="utf-8") as file:
        text = file.read()

    extracted_info = extract_information(text)

    with open("extracted_info.txt", "w", encoding="utf-8") as file:
        for key, value in extracted_info.items():
            file.write(f"{key}: {value}\n")

    print("Informations extraites sauvegardées dans 'extracted_info.txt'.")
