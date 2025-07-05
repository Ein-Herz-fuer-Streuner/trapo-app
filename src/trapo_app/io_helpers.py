import glob
import os
import shutil
import tkinter as tk
import unicodedata
from io import BytesIO
from pathlib import Path
from tkinter import filedialog

import pandas as pd
from PIL import Image
from docx import Document
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage


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


def read_file(path, images):
    df = pd.DataFrame()
    imgs = []
    try:
        if path.endswith(".xlsx"):
            df = read_excel(path)
        elif path.endswith(".csv"):
            df = pd.read_csv(path, dtype=str, keep_default_na=False)
        elif path.endswith(".docx"):
            if images:
                df, imgs = read_docx_images(path)
            else:
                df = read_docx(path)
    except ValueError:
        print("Datei ist ungültig")
        return df, imgs
    except FileNotFoundError:
        print(f"Datei {path} existiert nicht")
        return df, imgs
    except Exception as err:
        print("Etwas anderes ist schief gelaufen", err)
        return df, imgs
    return df, imgs


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


def read_files(files, images):
    res = []
    imgs = []
    for file in files:
        tmp, imgs = read_file(file, images)
        if not tmp.empty:
            res.append(tmp)
        else:
            print("Konnte Datei ", file, "nicht lesen")
    return res, imgs


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

def read_excel(file):
    # Read first 20 rows without headers to detect the header row
    preview = pd.read_excel(file, header=None, nrows=20)

    # Find first row index where there's at least one non-empty cell (likely header)
    header_row_idx = preview.apply(lambda row: row.notna().any(), axis=1).idxmax()

    # Read again using that row as header, skip everything before
    df = pd.read_excel(file, dtype=str, keep_default_na=False, header=header_row_idx)

    # Optional: drop fully empty rows below the table if any
    df = df.dropna(how='all')
    return df

def read_docx_images(path):
    document = Document(path)
    data = []
    images_map = []
    for table in document.tables:
        for row in table.rows:
            row_data = []
            row_images = []
            for cell_idx, cell in enumerate(list(iter_unique_cells(row.cells))):
                row_data.append(cell.text.strip())
                cell_images = []
                for para in cell.paragraphs:
                    for run in para.runs:
                        for rel in run.part.rels.values():
                            if "image" in rel.reltype:
                                img_blob = rel.target_part.blob
                                img_stream = BytesIO(img_blob)
                                img_stream.name = f"r{len(data)}_c{cell_idx}.png"
                                cell_images.append(img_stream)
                row_images.append(cell_images)
            data.append(row_data)
            images_map.append(row_images)

    # Convert the data to a DataFrame
    df = pd.DataFrame(data=data, dtype=str)

    # Optional: If the first row is the header
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)
    images_map = images_map[1:]
    return df, images_map


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
    return [x for x in files for s in ["nord", "süd", "sued", "sud", "südwest", "suedwest", "sudwest", "mitte"] if
            s in simple_normalize(x)]


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


def save_distance_sheets(paths, dfs, imgs, image_resize=(80, 80)):
    for file, df, pics in zip(paths, dfs, imgs):
        path, name = os.path.split(file)
        name, _ = os.path.splitext(name)
        # Save to Excel with formatting
        output_file = name + "_Entfernung.xlsx"
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Entfernung')

            workbook = writer.book
            worksheet = writer.sheets['Entfernung']

            # Define format: dark blue background, white bold font
            header_format = workbook.add_format({
                'bold': True,
                'font_color': 'white',
                'bg_color': '#294879',  # Dark blue
                'align': 'center'
            })

            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            # Apply formatting to rows where row equals column headers
            for row_num, row_data in df.iterrows():
                if row_data['Nr.'] == 'Nr.':
                    for col_num, col in enumerate(df.columns):
                        if str(row_data[col]) == col:
                            worksheet.write(row_num + 1, col_num, row_data[col], header_format)
            for column in df:
                column_length = max(df[column].astype(str).map(len).max(), len(column))
                col_idx = df.columns.get_loc(column)
                writer.sheets['Entfernung'].set_column(col_idx, col_idx, column_length)

        # Step 2: Re-open workbook to insert images
        '''
        wb = load_workbook(output_file)
        ws = wb['Entfernung']
        # Step 3: Add images
        for r_idx in range(len(pics)):
            for c_idx in range(len(pics[r_idx])):
                cell_images = pics[r_idx][c_idx]
                excel_row = r_idx + 2  # +2 because Excel is 1-indexed and row 1 is the header
                excel_col = c_idx + 1

                for img_stream in cell_images:
                    try:
                        img_stream.seek(0)
                        pil_img = Image.open(img_stream)
                        pil_img.thumbnail(image_resize)

                        img_buffer = BytesIO()
                        pil_img.save(img_buffer, format="PNG")
                        img_buffer.seek(0)
                        img_buffer.name = img_stream.name  # optional

                        xl_img = XLImage(img_buffer)
                        ws.add_image(xl_img, ws.cell(row=excel_row, column=excel_col).coordinate)

                    except Exception as e:
                        print(f"Failed to embed image: {e}")

        wb.save(output_file)
        '''
