from functions import extract_tables,extract_matches, display_information

if __name__ == "__main__":
    fichier_pptx = "example FASEP.pptx"
    patterns = ["Montant du FASEP", 
                "Date de signature de la convention", 
                "Avis sur le versement intermédiaire"]
    tableaux = extract_tables(fichier_pptx)
    test = extract_matches(tableaux, patterns)
    display_information(test)