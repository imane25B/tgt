from datetime import datetime
import re

def nettoyer_date(date_string):
    """
    Nettoie une chaîne de date en gérant plusieurs formats possibles.
    
    Args:
        date_string (str): Chaîne de date à nettoyer
        
    Returns:
        str: Date au format JJ/MM/AAAA HH:MM:SS
    """
    # Format JJ/MM/AAAA HH:MM:SS
    pattern1 = r'(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})'
    # Format AAAA-MM-JJ HH:MM:SS avec potentiellement des millisecondes et timezone
    pattern2 = r'(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})(?:\.\d+)?(?:[+-]\d{2}:?\d{2})?'
    
    # Essayer le premier format (JJ/MM/AAAA)
    match = re.search(pattern1, date_string)
    if match:
        return match.group(1)
    
    # Essayer le deuxième format (AAAA-MM-JJ)
    match = re.search(pattern2, date_string)
    if match:
        # Convertir le format AAAA-MM-JJ en JJ/MM/AAAA
        date_iso = match.group(1)
        try:
            date_obj = datetime.strptime(date_iso, '%Y-%m-%d %H:%M:%S')
            return date_obj.strftime('%d/%m/%Y %H:%M:%S')
        except ValueError:
            return None
            
    return None

def filter_dates(input_file, output_file):
    """
    Filtre les lignes d'un fichier texte basé sur une date de référence.
    Ne garde que les dates postérieures au 01/11/2024.
    Nettoie les dates qui contiennent des caractères supplémentaires.
    
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
            date_index = colonnes.index('DATE HEURE ENVOI')
        except ValueError:
            raise ValueError("La colonne 'Date_envoi' n'a pas été trouvée dans le fichier")
        
        # Traiter chaque ligne
        for ligne in f_in:
            champs = ligne.strip().split('|')
            try:
                # Nettoyer la date d'abord
                date_string = champs[date_index]
                date_propre = nettoyer_date(date_string)
                
                if date_propre:
                    date_envoi = datetime.strptime(date_propre, "%d/%m/%Y %H:%M:%S")
                    if date_envoi > date_reference:
                        # Mettre à jour le champ de date avec la version nettoyée
                        champs[date_index] = date_propre
                        # Reconstruire la ligne avec la date nettoyée
                        nouvelle_ligne = '|'.join(champs)
                        f_out.write(nouvelle_ligne + '\n')
            except (ValueError, IndexError) as e:
                print(f"Erreur de traitement pour la ligne: {ligne.strip()}")
                print(f"Erreur détaillée: {str(e)}")
                continue

# Exemple d'utilisation
if __name__ == "__main__":
    input_file = "a/output.txt"  # Remplacer par votre nom de fichier d'entrée
    output_file = "donnees_filtrees.txt"  # Fichier de sortie
    
    try:
        filter_dates(input_file, output_file)
        print("Filtrage terminé avec succès!")
    except Exception as e:
        print(f"Une erreur est survenue: {e}")
