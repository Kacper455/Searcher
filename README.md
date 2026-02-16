# Searcher

This scrypt is helping you find given expression in specified folder, in a certain  .xlsx file.

## Installation

Importing following libraries for code to run properly.

```python
import glob
import os
import re
from openpyxl.reader.excel import load_workbook
```

## Usage

```python
#searching function
def search_xlsx(path, regex):
    try:
        wb = load_workbook(path, data_only=True)

        #iterating through all sheets, rows and cells in given file 
        for sheet in wb.worksheets:
            for r,row in enumerate(sheet.iter_rows(values_only=True), start=1):
                for c,cell in enumerate(row, start=1):
                    
                    #converting text in cell to string
                    if cell is None:
                        continue
                    text = str(cell)
                    
                    #printing exact location of founded regex to the console
                    if regex.search(text):
                        print(
                            f"📊 {path} | arkusz '{sheet.title}' "
                            f"| komórka {r}:{c}: {text}"
                        )
    except Exception as e:
        print(f" Błąd XLSX ({path}): {e}")
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first
to discuss what you would like to change.

Please make sure to update tests as appropriate.
