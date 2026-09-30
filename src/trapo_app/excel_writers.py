"""Alle Excel-Ausgaben der Trapo-App."""
import os
from io import BytesIO

import pandas as pd
from PIL import Image

RO_TITLE = "CHIPLIST EUROPA"
RO_FOOTER_LEFT = "Data si ora plecarii"
RO_FOOTER_RIGHT = "Numele transportatorului"

DISTANCE_SHEET = "Entfernung"
HEADER_COLOR = "#294879"
MAX_IMAGE_SIZE = (100, 100)
PIXELS_PER_POINT = 0.75
PIXELS_PER_CHAR = 7


def _excel_writer(path):
    return pd.ExcelWriter(path, engine="xlsxwriter")


def _column_width(df, column):
    longest = max((len(str(value)) for value in df[column] if not pd.isna(value)), default=0)
    return max(longest, len(column))


def write_df_to_excel(df, path, sheet_name):
    """Schreibt `df` mit automatisch angepassten Spaltenbreiten in eine neue Excel-Datei."""
    with _excel_writer(path) as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]
        for idx, column in enumerate(df.columns):
            if column == "":
                continue
            worksheet.set_column(idx, idx, _column_width(df, column))


def _output_name(source_path, suffix=""):
    base, _ = os.path.splitext(os.path.basename(source_path))
    return f"{base}{suffix}.xlsx"


# ──────────────────────────────────────────────
#  Entfernungstabellen (trapo-km)
# ──────────────────────────────────────────────

def save_distance_sheets(paths, dfs, img_banks, img_column="Photo", max_img_size=MAX_IMAGE_SIZE):
    """Speichert je Eingabedatei eine formatierte Excel-Datei `<Name>_Entfernung.xlsx`.

    Die Tabellen enthalten Kopfzeilen-Markierungen (Spalte 'Nr.' == 'Nr.'), vor die ein
    Treffpunkt-Banner und die Spaltenüberschriften geschrieben werden.
    """
    for file_path, df, img_registry in zip(paths, dfs, img_banks):
        _save_distance_sheet(_output_name(file_path, "_Entfernung"), df, img_registry, img_column, max_img_size)


def _save_distance_sheet(out_path, df, img_registry, img_column, max_img_size):
    df_xls = df.copy()
    img_keys = None
    if img_column in df_xls.columns:
        img_keys = df_xls[img_column].astype(object).where(df_xls[img_column].notna(), None).tolist()
        df_xls[img_column] = ""  # only the pictures should show up in this column

    with _excel_writer(out_path) as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet(DISTANCE_SHEET)
        writer.sheets[DISTANCE_SHEET] = worksheet
        formats = {
            "header": workbook.add_format({
                'bold': True, 'font_color': 'white', 'bg_color': HEADER_COLOR,
                'align': 'center', 'valign': 'vcenter'}),
            "cell": workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'}),
            "meeting_point": workbook.add_format({
                'bold': True, 'font_color': 'white', 'bg_color': HEADER_COLOR,
                'align': 'center', 'valign': 'vcenter', 'font_size': 12}),
        }

        excel_rows = _write_rows(worksheet, df, df_xls, formats)
        _autosize_columns(worksheet, df_xls, skip=img_column)
        if img_keys is not None:
            _insert_images(worksheet, df_xls.columns.get_loc(img_column), img_keys, excel_rows,
                           img_registry, max_img_size)


def _write_rows(worksheet, df, df_xls, formats):
    """Schreibt Datenzeilen und Treffpunkt-Kopfzeilen; gibt die Excel-Zeile je df-Zeile zurück."""
    is_marker = (df['Nr.'] == 'Nr.').tolist()
    values = df_xls.to_numpy(dtype=object)
    excel_rows = {}
    row = 0
    for idx in range(len(df)):
        if is_marker[idx]:
            row = _write_group_header(worksheet, row, df, idx, formats)
            continue
        for col_num, value in enumerate(values[idx]):
            worksheet.write(row, col_num, "" if pd.isna(value) else value, formats["cell"])
        excel_rows[idx] = row
        row += 1
    return excel_rows


def _write_group_header(worksheet, row, df, marker_idx, formats):
    """Schreibt Treffpunkt-Banner (falls bekannt) und Spaltenüberschriften; gibt die nächste freie Zeile zurück."""
    next_idx = marker_idx + 1
    meeting_point = ""
    if next_idx < len(df) and 'Treffpunkt' in df.columns:
        meeting_point = str(df.iloc[next_idx]['Treffpunkt']).strip()
    if meeting_point:
        worksheet.merge_range(row, 0, row, len(df.columns) - 1, meeting_point, formats["meeting_point"])
        row += 1
    for col_num, name in enumerate(df.columns):
        worksheet.write(row, col_num, name, formats["header"])
    return row + 1


def _autosize_columns(worksheet, df, skip):
    for idx, column in enumerate(df.columns):
        if column != skip:
            worksheet.set_column(idx, idx, _column_width(df, column))


def _insert_images(worksheet, img_col_idx, img_keys, excel_rows, img_registry, max_img_size):
    max_col_width = 0
    max_w, max_h = max_img_size
    for idx, key in enumerate(img_keys):
        if key is None or key not in img_registry or idx not in excel_rows:
            continue
        buffer = BytesIO(img_registry[key])
        w_px, h_px = Image.open(buffer).size
        scale = min(1, max_w / w_px, max_h / h_px)  # only scale down
        scaled_w, scaled_h = int(w_px * scale), int(h_px * scale)
        max_col_width = max(max_col_width, scaled_w / PIXELS_PER_CHAR)

        excel_row = excel_rows[idx]
        worksheet.set_row(excel_row, scaled_h * PIXELS_PER_POINT)
        buffer.seek(0)
        worksheet.insert_image(excel_row, img_col_idx, "", {
            "image_data": buffer,
            "x_scale": scale,
            "y_scale": scale,
            "x_offset": 0,
            "y_offset": 0,
            "positioning": 1,  # moves/sizes with cells
        })
    worksheet.set_column(img_col_idx, img_col_idx, max_col_width)


# ──────────────────────────────────────────────
#  Rumänische Liste (trapo-ro)
# ──────────────────────────────────────────────

def save_ro_excel(dfs, files):
    """Speichert je Eingabedatei eine Excel-Datei mit Titel- und Fußzeile für Rumänien."""
    for df, file_path in zip(dfs, files):
        _save_ro_sheet(_output_name(file_path), df)


def _save_ro_sheet(out_path, df):
    sheet_name = 'Sheet1'
    with _excel_writer(out_path) as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = worksheet
        centered = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        title = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})

        last_col = len(df.columns)
        worksheet.merge_range(0, 0, 0, last_col - 1, RO_TITLE, title)

        first_data_row = 1  # directly below the title
        df.to_excel(writer, sheet_name=sheet_name, startrow=first_data_row, index=False)
        for idx, column in enumerate(df.columns):
            worksheet.set_column(idx, idx, _column_width(df, column), centered)

        footer_row = first_data_row + len(df) + 1
        if last_col >= 3:
            worksheet.merge_range(footer_row, 0, footer_row, 2, RO_FOOTER_LEFT, centered)
        if last_col >= 6:
            worksheet.merge_range(footer_row, last_col - 3, footer_row, last_col - 1, RO_FOOTER_RIGHT, centered)
        elif last_col > 3:  # fewer than 6 columns: right footer starts after the left one
            worksheet.merge_range(footer_row, 3, footer_row, last_col - 1, RO_FOOTER_RIGHT, centered)
