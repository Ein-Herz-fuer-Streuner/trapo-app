import os
import shutil
import tkinter as tk
import unicodedata
from pathlib import Path
from tkinter import filedialog
import glob

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
    df = pd.DataFrame()
    try:
        if path.endswith(".xlsx"):
            df = pd.read_excel(path, dtype=str, keep_default_na=False)
        elif path.endswith(".csv"):
            df = pd.read_csv(path, dtype=str, keep_default_na=False)
        elif path.endswith(".docx"):
            df = read_docx(path)
    except ValueError:
        print("Datei ist ungültig")
        return None
    except FileNotFoundError:
        print(f"Datei {path} existiert nicht")
        return None
    except Exception as err:
        print("Etwas anderes ist schief gelaufen", err)
        return None
    return df


def get_several_files_ui():
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(
        title="Wähle mehrere Word- oder Exceltabellen aus",
        filetypes=[("Word- und Excel-Dateien", "*.docx *.xlsx *.xls")]
    )
    return list(file_paths)


def get_file_ui():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Wähle eine Datei aus",
        filetypes=[("Word-,CSV- oder Excel-Datei", "*.docx *.xlsx *.xls *.csv")]
    )
    return file_path


def read_files(files):
    res = []
    for file in files:
        tmp = read_file(file)
        if not tmp.empty:
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
    df = pd.DataFrame(data=data, dtype=str)

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
    files_old = files_old["Kennzeichen"].values.tolist()
    files_old = [x for x in files_old if not x in ["", "?"]]
    for path in files_old:
        path = os.path.join(".", path)
        Path(path).mkdir(parents=True, exist_ok=True)
    return files_old


def move_files(df):
    seen = []
    for index, row in df.iterrows():
        file_ = row["Datei neu"]
        if file_ in ["", "?"]:
            continue
        file_old = row["Datei"]
        dir, _ = os.path.split(file_old)
        file_path = os.path.join(dir, file_)
        new_path = os.path.join(".", row["Kennzeichen"], file_)
        if file_ not in seen:
            shutil.move(file_path, new_path)
            seen.append(file_)

def simple_normalize(s):
    return (unicodedata.normalize("NFKD", s)
            .encode("ASCII", "ignore")
            .decode()
            .lower())


def filter_stopps(files):
    return [x for x in files for s in ["nord", "süd", "sued", "sud", "südwest", "suedwest", "sudwest", "mitte"] if s in simple_normalize(x)]

def move_and_rename(tuples):
    for plate, stop, file in tuples:
        path = os.path.join(".", plate)
        new_path = os.path.join(".", plate + "_" + stop)
        # move word doc to new folder
        shutil.move(file, path)
        # rename
        if os.path.exists(path):
            if not os.path.exists(new_path):
                os.rename(path, new_path)
            else:
                print("Target folder name already exists.")
        else:
            print("Original folder does not exist.")

def get_all_files_from_folder(glob_path):
    return glob.glob(glob_path)