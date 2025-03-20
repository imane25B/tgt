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
        c for c in unicodedata.normalize('NFD', input_str) if unicodedata.category(c) != 'Mn'
    )

def write_data_to_txt(data_list, success_data, output_filename="data.pdf.tresorerie1.txt"):
    """
    Écrit les données extraites dans un fichier .txt avec | comme délimiteur.
    Les colonnes sont ordonnées selon les exigences fournies, le texte est en majuscules et sans accents.

    Args:
        data_list (list): Liste des données extraites, avec chaque élément représentant une page.
        success_data (list): Données des blocs "success" à fusionner avec les données extraites.
        output_filename (str): Nom du fichier de sortie.
    """
    # Ordre des colonnes requis
    column_order = [
        "OBJET", "fax_destinataire", "Mail_Expediteur","Expéditeur", "DATE HEURE ACCUSE DE RECEPTION", 
        "DATE HEURE ENVOI", "N page", "Duree envoi", "Titre de l'expéditeur", 
        "Entité", "Direction", "Contact AXA 1", "Contact AXA 2", "Contact AXA 3", 
        "Destinataire", "Adresse Destinataire", "Tel Destinataire", 
        "Fax Destinataire", "Mail Destinataire", "Date document", 
        "Référence", "Compte à débiter", "SWIFT", "Titulaire de compte", 
        "Montant décaissement", "Devise", "Date valeur compensée", "Bénéficiaire", 
        "IBAN Bénéficiaire", "Banque Bénéficiaire", "Swift Bénéficiaire", 
        "Commission", "Motif du paiement", "Référence de l'opération", 
        "Signataire1", "Signataire2","PATH"
    ]

    # Vérifier si le fichier est vide pour écrire l'en-tête
    file_exists = os.path.exists(output_filename)
    is_empty = os.stat(output_filename).st_size == 0 if file_exists else True

    with open(output_filename, 'a', encoding='utf-8') as f:
        # Écrire l'en-tête si le fichier est vide
        if is_empty:
            f.write("|".join(remove_accents(col).upper() for col in column_order) + "\n")

        # Parcourir les données extraites
        for i, page_data in enumerate(data_list):
            for page_name, row in page_data.items():
                # Fusionner avec les données de succès correspondantes (si disponibles)
                success_info = success_data[i] if i < len(success_data) and isinstance(success_data[i], dict) else {}
                row.update(success_info)

                # Générer une ligne de données en respectant l'ordre des colonnes
                line = [remove_accents(str(row.get(col, ""))).upper() for col in column_order]
                f.write("|".join(line) + "\n")

    print(f"Données écrites avec succès dans {output_filename}")
