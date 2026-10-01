"""Extrahiert Daten aus den Traces-PDFs."""
import re

import camelot
import ftfy
import pandas as pd

from trapo_app.errors import TrapoError

CELLS_TO_KEEP = (
    "IMSOC",
    "Bestimmungsort ",
    "Identifikationsnummer",
    "| Rumänien",
)
DESTINATION_REGEX = re.compile(r"(.+);(.*\d+\s?[a-zA-Z]?)\s(\d{4,}\s.+)")


def get_table_data(files):
    """Gibt je Datei die relevanten Zellentexte (siehe CELLS_TO_KEEP) der ersten beiden Seiten zurück."""
    rows_per_file = []
    for file in files:
        try:
            tables = camelot.read_pdf(file, pages="1-2", flavor="lattice", backend="pdfium", line_scale=20)
        except Exception as err:
            raise TrapoError(f"Etwas ist beim PDF einlesen schief gegangen: {err}") from err

        rows = []
        for table in tables:
            for _, row in table.df.iterrows():
                for cell in row:
                    cell = ftfy.fix_text(cell)  # remove known extraction errors
                    cell = cell.replace('\n', ' ')  # for lattice mode with line-spanning entries
                    cell = re.sub(" +", ' ', cell)  # multiple whitespaces to only one
                    if any(keyword in cell for keyword in CELLS_TO_KEEP):
                        rows.append(cell)
        rows_per_file.append(rows)
    return rows_per_file


def _parse_intra(row, file):
    parts = row.split("Bezugsnummer ")
    if len(parts) < 2:
        print("Fehler: Datei hat keine Bezugsnummer:", file)
        return None
    intra = parts[1]
    return "I" + intra if intra.startswith("N") else intra


def _parse_destination(row, file):
    row = row.split("Name ")[1].split(" ISO")[0].replace(" Adresse ", ";")
    match = DESTINATION_REGEX.match(row)
    if not match:
        print("Fehler: Kein Bestimmungsort bei Datei:", file)
        return None
    return ", ".join(match.groups())


def _parse_chip(row, file):
    parts = row.split("Identifikationsnummer ")
    if len(parts) < 2:
        print("Fehler: Datei hat keine Identifikationsnummer:", file)
        return None
    return parts[1].replace("Microchip ", "")


def _parse_plate(row):
    return row.split(" |")[0].split(" ")[-1]


def extract_table_data(files):
    """Gibt je Chip eine Zeile [Datei, Intra, Chip, Kontakt, Kennzeichen] zurück."""
    results = []
    for file, rows in zip(files, get_table_data(files)):
        intra = contact = plate = ""
        chips = []
        current_row = ""
        try:
            for current_row in rows:
                if "IMSOC" in current_row:
                    intra = _parse_intra(current_row, file) or intra
                elif "Bestimmungsort" in current_row:
                    contact = _parse_destination(current_row, file) or contact
                elif "Identifikationsnummer" in current_row:
                    chip = _parse_chip(current_row, file)
                    if chip is not None:
                        chips.append(chip)
                elif "| Rumänien" in current_row:
                    plate = _parse_plate(current_row)
        except Exception as err:
            print("Fehler bei Datei und Zeile:", file, "\n", current_row, "\n", err)
            continue

        results.extend([file, intra, str(chip), contact, plate] for chip in chips)
    return results


def extract_traces(files):
    """Extrahiert alle Traces-Daten als DataFrame, sortiert nach Kontakt und Chip."""
    columns = ["Datei", "Intra", "Chip", "Kontakt", "Kennzeichen"]
    df = pd.DataFrame(extract_table_data(files), columns=columns)
    return df.sort_values(['Kontakt', 'Chip'])
