import csv

def convert_txt_to_csv(input_file, output_file):
    """
    Convertit un fichier texte avec délimiteur | en fichier CSV.
    
    Args:
        input_file (str): Chemin du fichier texte d'entrée
        output_file (str): Chemin du fichier CSV de sortie
    """
    try:
        # Lecture du fichier texte
        with open(input_file, 'r', encoding='utf-8') as txt_file:
            # Lecture des lignes et split sur le délimiteur |
            lines = [line.strip().split('|') for line in txt_file if line.strip()]
        
        # Écriture dans le fichier CSV
        with open(output_file, 'w', newline='', encoding='utf-8') as csv_file:
            writer = csv.writer(csv_file)
            # Écriture de chaque ligne
            for line in lines:
                writer.writerow(line)
                
        print(f"Conversion réussie ! Le fichier {output_file} a été créé.")
        
    except FileNotFoundError:
        print(f"Erreur : Le fichier {input_file} n'a pas été trouvé.")
    except Exception as e:
        print(f"Une erreur est survenue : {str(e)}")

# Exemple d'utilisation
if __name__ == "__main__":
    input_file = "donnees.txt"
    output_file = "donnees.csv"
    convert_txt_to_csv(input_file, output_file)
