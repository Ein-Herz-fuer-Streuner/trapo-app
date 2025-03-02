import os
import pandas as pd
from docx import Document

def get_file():
    is_valid = False
    path = ""
    while not is_valid:
        path = str(input("Bitte gib den Pfad zu einer Datei (.xlsx, .csv, .docx) ein:"))
        if path.endswith(".xlsx") or path.endswith(".csv") or path.endswith(".docx"):
            is_valid = True
    path = os.path.abspath(path)
    return path

def read_file(path):
    df = []
    try:
        if path.endswith(".xlsx"):
            df = pd.read_excel(path)
        elif path.endswith(".csv"):
            df = pd.read_csv(path)
        elif  path.endswith(".docx"):
            df = read_docx(path)
    except ValueError:
        print("Datei ist ungültig")
        return None
    except FileNotFoundError:
        print(f"Datei {path} existiert nicht")
        return None
    except:
        print("Etwas anderes ist schief gelaufen")
        return None
    return df

def read_docx(path):
    document = Document(path)
    tables = []
    for table in document.tables:
        # Create a DataFrame structure with empty strings, sized by the number of rows and columns in the table
        df = [['' for _ in range(len(table.columns))] for _ in range(len(table.rows))]

        # Iterate through each row in the current table
        for i, row in enumerate(table.rows):
            # Iterate through each cell in the current row
            for j, cell in enumerate(row.cells):
                # If the cell has text, store it in the corresponding DataFrame position
                if cell.text:
                    df[i][j] = cell.text

        # Convert the list of lists (df) to a pandas DataFrame and add it to the tables list
        tables.append(pd.DataFrame(df))
    return tables[0]
