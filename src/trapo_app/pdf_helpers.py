import re
import sys

import camelot
import ftfy
import pandas as pd

cells_to_keep = [
    "IMSOC",
    "Bestimmungsort ",
    "Identifikationsnummer",
    "| Rumänien"
]

def get_table_data(files):
    rows_per_file = []
    for file in files:
        rows = []
        tbls = []
        try:
            tbls = camelot.read_pdf(file, pages="1-2", flavor="lattice", backend="pdfium", line_scale=20)
        except Exception as err:
            print("Etwas ist beim PDF einlesen schief gegangen:", err)
            sys.exit(1)
        for tab in tbls:
            tab = tab.df
            for i, row in tab.iterrows():
                for cell in row:
                    cell = ftfy.fix_text(cell)  # remove known extraction errors
                    cell = cell.replace('\n', ' ')  # for lattice mode with line-spanning entries
                    cell = re.sub(" +", ' ', cell)  # multiple whitespaces to only one
                    for w in cells_to_keep:
                        if w in cell:
                            rows.append(cell)
        rows_per_file.append(rows)

    return rows_per_file


def extract_table_data(files):
    results = []
    raw = get_table_data(files)

    for file, rows in zip(files,raw):
        intra = ""
        contact = ""
        chips = []
        kennzeichen = ""
        for row in rows:
            if "IMSOC" in row:
                row = row.split("Bezugsnummer ")[1]
                if row.startswith("N"):
                    row = "I" + row
                intra = row
            elif "Bestimmungsort" in row:
                row = row.split("Name ")[1]
                row = row.split(" ISO")[0]
                row = row.replace(" Adresse ", ";")
                match = re.match(r"(.+);(.*\d+\s?[a-zA-Z]?)\s(\d{4,}\s.+)", row)
                if not match:
                    print("Fehler bei Datei und Zeile:", file, row)
                    continue
                row = match.group(1) + ", " + match.group(2) + ", " + match.group(3)
                contact = row
            elif "Identifikationsnummer" in row:
                row = row.split("Identifikationsnummer ")[1]
                row = row.replace("Microchip ", "")
                chips.append(row)
            elif "| Rumänien" in row:
                row = row.split(" |")[0]
                kennzeichen = row.split(" ")[-1]

        for chip in chips:
            results.append([file, intra, str(chip), contact, kennzeichen])

    return results


def extract_traces(files):
    cols = ["Datei", "Intra", "Chip", "Kontakt", "Kennzeichen"]
    results = extract_table_data(files)
    df_result = pd.DataFrame(results, columns=cols)
    df_result = df_result.sort_values(['Kontakt', 'Chip'])
    return df_result
