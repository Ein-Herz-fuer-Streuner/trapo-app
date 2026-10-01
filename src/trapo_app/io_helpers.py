"""Einlesen von Tabellen (Word, Excel, CSV) und Verwalten der Traces-Dateien."""
import glob
import os
import shutil
import unicodedata
from pathlib import Path

import pandas as pd
from docx import Document
from docx.oxml.ns import qn

STOP_KEYWORDS = ("nord", "mitte", "sud", "sued")  # matched against the ASCII-normalized file name


def read_file(path, images):
    """Liest eine .xlsx-, .csv- oder .docx-Datei; gibt (DataFrame, Bilder) zurück, bei Fehlern ein leeres DataFrame."""
    df = pd.DataFrame()
    imgs = {}
    try:
        if path.endswith(".xlsx"):
            df = read_excel(path)
        elif path.endswith(".csv"):
            df = pd.read_csv(path, dtype=str, keep_default_na=False)
        elif path.endswith(".docx"):
            if images:
                df, imgs = read_docx_with_images(path)
            else:
                df = read_docx(path)
    except ValueError:
        print("Datei ist ungültig")
    except FileNotFoundError:
        print(f"Datei {path} existiert nicht")
    except Exception as err:
        print("Etwas anderes ist schief gelaufen", err)
    return df, imgs


def read_files(files, images):
    """Liest mehrere Dateien; nicht lesbare Dateien werden übersprungen."""
    dfs = []
    imgs = []
    for file in files:
        df, img_dict = read_file(file, images)
        if df.empty:
            print("Konnte Datei ", file, "nicht lesen")
            continue
        dfs.append(df)
        imgs.append(img_dict)
    return dfs, imgs


def iter_unique_cells(cells):
    """Überspringt direkt aufeinanderfolgende Duplikate (verbundene Zellen)."""
    prior_cell = None
    for cell in cells:
        if cell == prior_cell:
            continue
        yield cell
        prior_cell = cell


def read_docx(path):
    """Liest die Tabellen eines Word-Dokuments; die erste Zeile ist die Kopfzeile."""
    document = Document(path)
    data = [
        [cell.text.strip() for cell in iter_unique_cells(row.cells)]
        for table in document.tables
        for row in table.rows
    ]
    return _first_row_as_header(pd.DataFrame(data=data, dtype=str))


def read_excel(file):
    """Liest eine Excel-Datei; die Kopfzeile ist die erste nicht-leere Zeile."""
    preview = pd.read_excel(file, header=None, nrows=20)
    header_row_idx = preview.notna().any(axis=1).idxmax()
    df = pd.read_excel(file, dtype=str, keep_default_na=False, header=header_row_idx)
    return df.dropna(how='all')


def _first_row_as_header(df):
    df.columns = df.iloc[0]
    return df[1:].reset_index(drop=True)


def _extract_cell(cell):
    """Gibt (Text, Bilder als Bytes) einer Word-Zelle zurück."""
    paragraph_texts = []
    for para in cell.paragraphs:
        fragments = []
        for run in para.runs:
            for node in run._element.iter():
                if node.tag.endswith('}t') and node.text:
                    fragments.append(node.text)
                elif node.tag.endswith('}br'):
                    fragments.append('\n')  # soft line break
        paragraph_texts.append(''.join(fragments))

    pictures = []
    for node in cell._element.iter():
        if node.tag.endswith('}blip'):
            r_id = node.get(qn('r:embed'))
            if r_id:
                pictures.append(cell.part.related_parts[r_id].blob)

    return '\n'.join(paragraph_texts).strip(), pictures


def read_docx_with_images(path, img_column='Photo'):
    """Wie `read_docx`, aber Zellen mit Bild enthalten einen Schlüssel für die zurückgegebene Bild-Map."""
    doc = Document(path)
    rows, img_registry = [], {}
    img_col_index = None  # first column that actually contains a picture

    for table in doc.tables:
        for row in table.rows:
            current = []
            for c_idx, cell in enumerate(row.cells):
                text, pictures = _extract_cell(cell)
                if pictures:  # keep only the first image per cell
                    key = f"img_{len(img_registry)}"
                    img_registry[key] = pictures[0]
                    if img_col_index is None:
                        img_col_index = c_idx
                    current.append(key)
                else:
                    current.append(text)
            rows.append(current)

    df = _first_row_as_header(pd.DataFrame(rows, dtype=str))
    if img_col_index is not None:
        df = df.rename(columns={df.columns[img_col_index]: img_column})
    return df, img_registry


def rename_files(old_paths, new_names):
    """Benennt jede Datei innerhalb ihres Ordners um."""
    for old, new in zip(old_paths, new_names):
        os.rename(old, os.path.join(os.path.dirname(old), new))


def create_folders(df):
    """Legt für jedes Kennzeichen einen Ordner an und gibt die Kennzeichen zurück."""
    plates = [p for p in df["Kennzeichen"].drop_duplicates() if p not in ("", "?")]
    for plate in plates:
        Path(".", plate).mkdir(parents=True, exist_ok=True)
    return plates


def move_files(df):
    """Verschiebt die umbenannten Traces-Dateien in den Ordner ihres Kennzeichens."""
    moved = set()
    for _, row in df.iterrows():
        new_name = row["Datei neu"]
        if new_name in ("", "?") or new_name in moved:
            continue
        source = os.path.join(os.path.dirname(row["Datei"]), new_name)
        shutil.move(source, os.path.join(".", row["Kennzeichen"], new_name))
        moved.add(new_name)


def simple_normalize(s):
    """Kleinbuchstaben ohne Akzente/Umlaute (ü -> u)."""
    return (unicodedata.normalize("NFKD", s)
            .encode("ASCII", "ignore")
            .decode()
            .lower())


def filter_stopps(files):
    """Behält nur Dateien, deren Name einen Trapo-Stopp (Nord, Mitte, Süd, Südwest) enthält."""
    return [f for f in files if any(keyword in simple_normalize(f) for keyword in STOP_KEYWORDS)]


def move_and_rename(tuples):
    """Verschiebt je (Kennzeichen, Stopp, Word-Datei) die Datei in den Ordner und benennt ihn 'Kennzeichen_Stopp'."""
    for plate, stop, file in tuples:
        folder = os.path.join(".", plate)
        new_folder = os.path.join(".", f"{plate}_{stop}")
        if not os.path.exists(folder):
            print("Ursprünglicher Ordner existiert nicht.")
            continue
        shutil.move(file, folder)
        if os.path.exists(new_folder):
            print("Zielordner existiert bereits.")
        else:
            os.rename(folder, new_folder)


def get_all_files_from_folder(glob_path):
    return glob.glob(glob_path)


def _row_cell_text(tr, column):
    """Text der Zelle `column` einer Word-Tabellenzeile (leer, wenn die Zeile kürzer ist)."""
    cells = tr.findall(qn('w:tc'))
    if column >= len(cells):
        return ""
    return ''.join(t.text for t in cells[column].iter(qn('w:t')) if t.text).strip()


def sort_word_table(input_path, output_path, sort_column, sort_order):
    """
    Sortiert die erste Word-Tabelle nach vorgegebener Reihenfolge.
    Alle Zellformatierungen bleiben 1:1 erhalten, die Kopfzeile bleibt oben.

    sort_column: Index der Spalte nach der sortiert wird (0-basiert)
    sort_order:  gewünschte Reihenfolge als Liste, z.B. ["München", "Berlin", "Hamburg"];
                 Zeilen, die nicht darin vorkommen, landen am Ende.
    """
    doc = Document(input_path)
    tbl = doc.tables[0]._tbl
    data_rows = tbl.findall(qn('w:tr'))[1:]  # keep the header row untouched
    rank = {value: position for position, value in enumerate(sort_order)}

    def row_rank(tr):
        return rank.get(_row_cell_text(tr, sort_column), len(sort_order))

    for tr in data_rows:
        tbl.remove(tr)
    for tr in sorted(data_rows, key=row_rank):
        tbl.append(tr)

    doc.save(output_path)


def split_word_table(input_path, output_base, split_column, parts):
    """
    Teilt die erste Word-Tabelle in mehrere Dateien `<output_base>_<Teil>.docx`.
    Jede Datei ist eine Kopie des Originals (Formatierung bleibt erhalten), in der nur die
    Zeilen der Teilliste übrig bleiben; die Kopfzeile bleibt immer stehen.

    split_column: Index der Spalte mit dem Treffpunkt (0-basiert)
    parts:        {Teilname: [Treffpunkte]}
    Gibt {Teilname: (Pfad, Anzahl Zeilen)} zurück; Teile ohne Zeilen werden nicht gespeichert.
    """
    written = {}
    for name, meeting_points in parts.items():
        doc = Document(input_path)
        tbl = doc.tables[0]._tbl
        wanted = set(meeting_points)
        kept = 0
        for tr in tbl.findall(qn('w:tr'))[1:]:
            if _row_cell_text(tr, split_column) in wanted:
                kept += 1
            else:
                tbl.remove(tr)
        if not kept:
            continue
        path = f"{output_base}_{_safe_file_part(name)}.docx"
        doc.save(path)
        written[name] = (path, kept)
    return written


def _safe_file_part(name):
    """Ersetzt Zeichen, die in Dateinamen nicht erlaubt sind."""
    return "".join("-" if c in '\\/:*?"<>|' else c for c in name).strip()
