from datetime import datetime

def filter_dates(input_file, output_file):
    """
    Filtre les lignes d'un fichier texte basé sur une date de référence.
    Ne garde que les dates postérieures au 01/11/2024.
    
    Args:
        input_file (str): Chemin du fichier d'entrée
        output_file (str): Chemin du fichier de sortie
    """
    # Date de référence
    date_reference = datetime.strptime("01/11/2024 00:00:00", "%d/%m/%Y %H:%M:%S")
    
    # Lecture et traitement du fichier
    with open(input_file, 'r', encoding='utf-8') as f_in, open(output_file, 'w', encoding='utf-8') as f_out:
        # Lire l'en-tête
        header = f_in.readline().strip()
        # Écrire l'en-tête dans le fichier de sortie
        f_out.write(header + '\n')
        
        # Trouver l'index de la colonne Date_envoi
        colonnes = header.split('|')
        try:
            date_index = colonnes.index('Date_envoi')
        except ValueError:
            raise ValueError("La colonne 'Date_envoi' n'a pas été trouvée dans le fichier")
        
        # Traiter chaque ligne
        for ligne in f_in:
            champs = ligne.strip().split('|')
            try:
                date_envoi = datetime.strptime(champs[date_index], "%d/%m/%Y %H:%M:%S")
                if date_envoi > date_reference:
                    f_out.write(ligne)
            except (ValueError, IndexError):
                print(f"Erreur de traitement pour la ligne: {ligne.strip()}")
                continue

# Exemple d'utilisation
if __name__ == "__main__":
    input_file = "donnees.txt"  # Remplacer par votre nom de fichier d'entrée
    output_file = "donnees_filtrees.txt"  # Fichier de sortie
    
    try:
        filter_dates(input_file, output_file)
        print("Filtrage terminé avec succès!")
    except Exception as e:
        print(f"Une erreur est survenue: {e}")
