#!/bin/python3.13
"""Konsolen-Befehle der Trapo-App; jede öffentliche Funktion ist ein Entry-Point in setup.cfg."""
import functools
import sys
from pathlib import Path

import trapo_app.excel_writers as excel_writers
import trapo_app.gui as gui
import trapo_app.io_helpers as io_helpers
import trapo_app.pdf_helpers as pdf_helpers
import trapo_app.table_helpers as table_helpers
from trapo_app.errors import TrapoError

ADDRESS_COLUMN_HINT = "WICHTIG: BITTE GIB DER ADRESSENSPALTE DEN NAMEN 'KONTAKT'!"
CHAT_FILE_PROMPT = "Gib als erstes den Pfad zur Datei aus dem Messenger ein, z.B. './data/chat.docx'"
MULTIPLE_FILES_HINT = "Wähle im sich gleich öffnenden Fenster alle Dateien aus, {}. Kehre danach hierhin zurück."


def cli_command(func):
    """Zeigt Fehler mit verständlicher Meldung an, statt einen Stacktrace auszugeben."""

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except TrapoError as err:
            print(err)
            sys.exit(1)

    return wrapper


def _ask_for_file(prompt):
    print(prompt)
    return gui.get_file_ui()


def _read_table_ui(prompt):
    """Fragt mit `prompt` nach einer Datei und liest sie als Tabelle ein."""
    df, _ = io_helpers.read_file(_ask_for_file(prompt), False)
    return df


def _output_path(file_name):
    return Path.cwd() / file_name


def main():
    print("Willkommen bei der Trapo App! 🐶🚐")
    print()
    print("Du kannst folgende Konsolen-Befehle benutzen:")
    print("- trapo-vergleich: Vergleicht 2 Tabellen")
    print("- trapo-extrakt: Extrahiert Daten aus den Traces Dokumenten")
    print("- trapo-traces-vergleich: Vergleicht die extrahierten Traces-Daten sie mit einer Tabelle")
    print("- trapo-traces: Benennt Traces Dokumente um")
    print("- trapo-km: Das Rosi-Spezial :) Sortiert eine Tabelle nach Entfernungen der angegeben Adressen")
    print("- trapo-kombi: Kombiniert mehrere Tabellen zu einer")
    print("- trapo-ro: Erstellt eine Excel-Liste mit rumänischen Titeln aus der Trapotabelle")
    print("- trapo-sort: Sortiert die Tabelle basierend auf der eingegebenen Reihenfolge der Treffpunkte")
    print("- trapo-split: Teilt die Tabelle anhand der Treffpunkte in einzelne Word-Dateien (z.B. Nord, Südwest) auf")


@cli_command
def compare():
    df1 = _read_table_ui(f"{CHAT_FILE_PROMPT}\n{ADDRESS_COLUMN_HINT}")
    df2 = _read_table_ui("Gib nun den Pfad zur Datei aus PetOffice ein, z.B. './data/po.docx'")
    print("Vergleiche...")
    df = table_helpers.compare(df1, df2)
    path = _output_path("Trapo_Vergleich.xlsx")
    excel_writers.write_df_to_excel(df, path, 'Auto-Vergleich')
    print(f"Fertig! Die Vergleichsdatei liegt unter '{path}'")


@cli_command
def extract():
    print("Gib den Pfad zum Ordner an, in dem alle Traces-Dokumente liegen, z.B. './data/traces'")
    pdfs = gui.get_several_files_ui(".pdf")
    if not pdfs:
        raise TrapoError("Keine PDFs gefunden")
    print("Extrahiere Informationen... Das kann etwas dauern...")
    df = pdf_helpers.extract_traces(pdfs)
    path = _output_path("Traces_Extrakt.xlsx")
    excel_writers.write_df_to_excel(df, path, 'Traces')
    print(f"Fertig! Die Datei liegt unter '{path}'")


@cli_command
def compare_with_traces():
    df1 = _read_table_ui("Gib als erstes den Pfad zur Trapo_Vergleich-Tabelle ein, z.B. './Trapo_Vergleich.xlsx'")
    df2 = _read_table_ui("Gib nun den Pfad zu Traces_Extrakt-Tabelle ein, z.B. './Traces_Extrakt.xlsx'")
    print("Vergleiche...")
    df = table_helpers.compare_traces(df1, df2)
    path = _output_path("Trapo_Traces_Vergleich.xlsx")
    excel_writers.write_df_to_excel(df, path, 'Traces-Vergleich')
    print(f"Fertig! Die Vergleichsdatei liegt unter '{path}'")


@cli_command
def rename():
    df = _read_table_ui("Gib nun den Pfad zur Trapo_Vergleich-Tabelle ein, z.B. './Trapo_Traces_Vergleich.xlsx'")
    print("Baue neuen Dateinamen...")
    df, old, new = table_helpers.write_new_file_names(df)
    print(f"Benenne {len(new)} Dateien um...")
    io_helpers.rename_files(old, new)
    print("Erstelle Kennzeichen-Ordner und ordne Traces zu...")
    folders = io_helpers.create_folders(df)
    io_helpers.move_files(df)
    print("Wähle nun alle Word-Dokumente zu den Trapo-Stopps aus, z.B. '04.05.25-NORD-V1-Name.docx'")
    files = io_helpers.filter_stopps(gui.get_several_files_ui(".docx"))
    if not files:
        raise TrapoError("Keine .docx-Dateien gefunden")
    dfs, _ = io_helpers.read_files(files, False)
    print("Benenne Ordner um nach Stopp und verschiebe Word-Datei...")
    io_helpers.move_and_rename(table_helpers.find_stopp_for_plate(files, dfs, folders))
    print("Fertig, alle Traces-Dateien wurden umbenannt und verschoben.")


@cli_command
def distance():
    print(ADDRESS_COLUMN_HINT)
    print(MULTIPLE_FILES_HINT.format("für die du die Entfernung berechnen willst"))
    files = gui.get_several_files_ui()
    dfs, imgs = io_helpers.read_files(files, True)
    plates = _read_table_ui("Gib als jetzt den Pfad zur Kennzeichen-Datei ein, z.B. './data/kennzeichen.csv'")
    stopps = _read_table_ui("Gib nun den Pfad zur Trapo-Stopp-Liste ein, z.B. 'Trapo-Adressen.xlsx'")
    print("Verkleinere Tabellen...")
    dfs = table_helpers.clean_plate_dfs(table_helpers.shrink_tables(dfs))
    print("Füge Kennzeichen hinzu...")
    dfs = table_helpers.add_plates(dfs, plates)
    print("Berechne Anfahrten & sortiere nach Entfernung...")
    dfs = table_helpers.add_distance(dfs, stopps)
    print("Fertig!\nSpeichere...")
    excel_writers.save_distance_sheets(files, dfs, imgs)
    print(f"Fertig! Die Datei liegt im Ordner '{Path.cwd()}'")


@cli_command
def combine():
    print(MULTIPLE_FILES_HINT.format("die du kombinieren willst"))
    files = gui.get_several_files_ui()
    dfs, _ = io_helpers.read_files(files, False)
    df = table_helpers.combine_dfs(dfs)
    path = _output_path("Trapo_Kombiniert.xlsx")
    excel_writers.write_df_to_excel(df, path, 'Kombi')
    print(f"Fertig! Die Datei liegt unter '{path}'")


@cli_command
def translate():
    print(CHAT_FILE_PROMPT)
    print(ADDRESS_COLUMN_HINT)
    files = gui.get_several_files_ui()
    dfs, _ = io_helpers.read_files(files, False)
    excel_writers.save_ro_excel(table_helpers.translate_headers(dfs), files)
    print(f"Fertig! Die Datei liegt im Ordner '{Path.cwd()}'")


@cli_command
def sort_by_tp():
    source = _ask_for_file(f"{CHAT_FILE_PROMPT}\n{ADDRESS_COLUMN_HINT}")
    df, _ = io_helpers.read_file(source, False)
    if "Treffpunkt" not in df.columns:
        raise TrapoError("Die Tabelle hat keine Spalte 'Treffpunkt'.")
    sorted_tps = gui.get_sorted_tps(df['Treffpunkt'].unique())
    output = _output_path(f"{Path(source).stem}_sortiert.docx")
    io_helpers.sort_word_table(source, str(output), list(df.columns).index("Treffpunkt"), sorted_tps)
    print(f"Fertig! Die Datei liegt im Ordner '{Path.cwd()}'")


@cli_command
def split_by_tp():
    source = _ask_for_file(f"{CHAT_FILE_PROMPT}\n{ADDRESS_COLUMN_HINT}")
    if not source.endswith(".docx"):
        raise TrapoError("Zum Aufteilen wird eine Word-Datei (.docx) benötigt.")
    df, _ = io_helpers.read_file(source, False)
    if "Treffpunkt" not in df.columns:
        raise TrapoError("Die Tabelle hat keine Spalte 'Treffpunkt'.")
    print("Gib im sich öffnenden Fenster die Namen der Teillisten ein, z.B. Nord, Südwest, Conivet-Mitte.")
    names = gui.get_part_names()
    if not names:
        raise TrapoError("Abgebrochen: Es wurden keine Teillisten angelegt.")
    print("Ordne nun jedem Treffpunkt eine Teilliste zu.")
    parts = gui.get_part_assignments(names, list(df['Treffpunkt'].unique()))
    if not parts:
        raise TrapoError("Abgebrochen: Die Treffpunkte wurden nicht zugeordnet.")
    written = io_helpers.split_word_table(
        source, str(_output_path(Path(source).stem)), list(df.columns).index("Treffpunkt"), parts)
    for name in parts:
        if name in written:
            path, rows = written[name]
            print(f"- {Path(path).name}: {rows} Zeilen")
        else:
            print(f"- {name}: keine Zeilen, Datei wurde nicht erstellt")
    print(f"Fertig! Die Dateien liegen im Ordner '{Path.cwd()}'")


if __name__ == "__main__":
    main()
