from functions import extract_tables, extract_matches, display_information

def main():
    pptx_file = "example FASEP.pptx"
    patterns = ["Montant du FASEP", 
                "Date de signature de la convention", 
                "Avis sur le versement intermédiaire"]
    tables = extract_tables(pptx_file)
    result = extract_matches(tables, patterns)
    display_information(result)

if __name__ == "__main__":
    main()