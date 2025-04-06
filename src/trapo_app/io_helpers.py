import os
from pathlib import Path
import shutil

import pandas as pd
from docx import Document

def get_files(dir_, ending):
    pdfs = []
    dir_ = os.path.abspath(dir_)
    for filename in os.listdir(dir_):
        if filename.endswith(ending):
            full_path = os.path.join(dir_, filename)
            pdfs.append(full_path)
        else:
            continue
    return pdfs

def get_file():
    is_valid = False
    path = ""
    while not is_valid:
        path = str(input("Bitte gib den Pfad zu einer Datei (.xlsx, .csv, .docx) ein:"))
        path = path.replace("\"", "")
        path = path.replace("'", "")
        if path.endswith(".xlsx") or path.endswith(".csv") or path.endswith(".docx"):
            is_valid = True
    path = os.path.abspath(path)
    return path

def get_path():
    is_valid = False
    path = ""
    while not is_valid:
        path = str(input("Bitte gib den Pfad zu einem Order ein:"))
        if os.path.exists(path):
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
        elif path.endswith(".docx"):
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


def read_files(files):
    res = []
    for file in files:
        tmp = read_file(file)
        if tmp:
            res.append(tmp)
    return res

def iter_unique_cells(cells):
    prior_cell = None
    for c in cells:
        if c == prior_cell:
            continue
        yield c
        prior_cell = c


def read_docx(path):
    document = Document(path)
    data = []
    for table in document.tables:
        for row in table.rows:
            data.append([cell.text.strip() for cell in list(iter_unique_cells(row.cells))])

    # Convert the data to a DataFrame
    df = pd.DataFrame(data)

    # Optional: If the first row is the header
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)
    return df

def rename_files(col_old, col_new):
    for old, new in zip(col_old, col_new):
        dir, file = os.path.split(old)
        new_path = os.path.join(dir, new)
        os.rename(old, new_path)

def create_folders(df):
    files_old = df.drop_duplicates(subset="Kennzeichen", keep="first")
    files_old = files_old["Datei"].values.tolist()
    files_old = [x for x in files_old if not x in ["", "?"]]
    for path in files_old:
        path = os.path.join(".", path)
        Path(path).mkdir(parents=True, exist_ok=True)

def move_files(df):
    for row in df.itertuples():
        file_ = row["Datei neu"]
        dir, file = os.path.split(file_)
        new_path = os.path.join(".", row["Kennzeichen"], file)
        shutil.move(file_, new_path)

def filter_stopps(files):
    return [x for x in files if s in str.lower(x) for s in ["nord", "süd", "sued", "südwest", "suedwest", "mitte"]]