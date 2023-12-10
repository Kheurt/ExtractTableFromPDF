# ExtractTableFromPDF
Extract tables from a PDF file or a Microsoft Word file with python

# Recapitulatif de l'extraction sur un pdf
Notre principale tâche ici est d’extraire les tableaux présents dans les documents pour les mettre sous une forme facilement maniable pour la data science : nous <b>optons pour le format csv</b>.</br>
Pour extraire les tableaux, nous avons tout d’abord convertis les fichiers `PDF` en fichier `DOCX (Windows Docs)`. Ensuite, nous avons écrit un programme `python` qui sélectionne les tableaux des différents fichiers, lis ces derniers ligne par ligne afin de construire le tableau csv correspondant. Tout ce processus est possible en se servant de l’api présent dans la bibliothèque `docx`.

[Mail Me](mailto:krysrkl@tswansite.com)


# Instructions
Avant de commencer, vous devez noter que ce travail a été entièrement réalisé sur Google Colab. Donc il est possible que reproduire les instructions qui vont suivre produisent des érreurs liées à la compatibilité de votre environnement avec les bibliothèques. Donc si votre environnement local n'est pas à jours ou présente des conflits pour certaines dépendences, je vous recommande d'utiliser [Google Colab](https://https://colab.research.google.com/).

* Commençons par installer les bibliothèques
 
 ```bash
  pip install docx
  pip install python-docx
 ```
* Ensuite importez le module `Document` de l'api
  
  ```python
  from docx.api import Document
  ```

* Chargez les fichiers docx
  ```python
  dataset1Path = "drive/MyDrive/Hackaton/CAYSTI/Annuaire2019_2020.docx"
  dataset2Path = "drive/MyDrive/Hackaton/CAYSTI/Annuaire2020_2021.docx"
  ```

* Initialisez l'objet document word
```python
  document1 = Document(dataset1Path)
  document2 = Document(dataset2Path)
```

* Extraction des tables du fichier

```python
  allTables1 = document1.tables
  allTables2 = document2.tables
```

* Initialisation de quelques variable pour contrôler notre curseur et le séparateur de colonnes des CSV
```python
  data_to_print = ""
  col_separator = ";"
  output_index = 0
```

* Maibtenant parcourons chaque tableau pour construire une tableau Word pour contruire nos tableaux CSV
```python
for single_table in document1.tables:
    # Browse table rows
    table_rows = single_table.rows
    # Browse cells in the row
    row_cells = None
    for table_row in table_rows:
      row_cells = table_row.cells # list of cells
      for cell in row_cells:
        data_to_print += cell.text + col_separator
      # print the end_of_line
      data_to_print += '\n'
      print(data_to_print)

 # Sauvegarde du texte dans un fichier texte
 with open(f'out1/dataTab{output_index}.csv', 'w', encoding='utf-8') as txt_file:
       txt_file.write(data_to_print)
 # reset the data file
 data_to_print = ""
 # go to next table_index
 output_index += 1
```


# Now zip the output to download the csv files
```python
import os
import zipfile

def zipdir(path, ziph):
    # ziph is zipfile handle
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file),
                       os.path.relpath(os.path.join(root, file),
                                       os.path.join(path, '..')))

```
## Zip out1 (Annuaire 2019_2020)
```python
with zipfile.ZipFile('out1.zip', 'w', zipfile.ZIP_DEFLATED) as zipf:
    zipdir('out1/', zipf)
```

## zip out2 (Annuaire 2020_2021)
```python
with zipfile.ZipFile('out2.zip', 'w', zipfile.ZIP_DEFLATED) as zipf:
    zipdir('out2/', zipf)
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first
to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License

[MIT](https://choosealicense.com/licenses/mit/)
