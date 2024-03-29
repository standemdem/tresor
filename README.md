# Test Technique du Trésor Public

### Consignes

Ecrire un code python qui permet de:
* récupérer sur le powerpoint « Exemple FASEP.pptx » les dates de signature de la convention de don de la subvention du fonds d’étude et d’aide au secteur privé (FASEP): DONE
* le montant de la subvention du FASEP : DONE   
* l’avis du service économique de l’ambassade pour le premier terme intermédiaire de la subvention: DONE
* Assurez-vous que le code est robuste à des modifications possibles du Powerpoint.
* Si possible ajoutez des unit tests sur le code python que vous avez écrit.

### Program Setup
I am working on Linux with WSL2 using Ubuntu 22.04.3 LTS.
Follow these steps after cloning my repository.

1. Create a virtual environment
```python
python3 -m venv .venv
```
2. Activate the virtual environment
```
source .venv/bin/activate
```
3. Install the required packages
For this project, only pptx and pytest need to be installed, other packages are not needed, but I didn't have the time to clean up the requirements.txt.

If you want to make sure that it works, use the command bellow

```
pip install requirements.txt
```
4. Launch the program
```
python3 main.py
```

5. Start the tests
```
python3 -m pytest tests/
```

