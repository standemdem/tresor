import pytest
import pptx
from functions import extract_tables


@pytest.fixture
def ppt_file():
    # Chemin vers un fichier PowerPoint contenant des tableaux pour les tests
    return "example FASEP.pptx"  

def test_extract_non_empty_tables(ppt_file):
    # Vérification que la liste de tableaux n'est pas vide
    tableaux = extract_tables(ppt_file)
    assert tableaux  

def test_extract_tables_types(ppt_file):
    # Vérification que chaque élément de la liste est bien un tableau
    tables = extract_tables(ppt_file)
    for table in tables:
        assert isinstance(table, pptx.table.Table)  
