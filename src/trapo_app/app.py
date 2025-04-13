#!/bin/python3.13
import os.path
import sys

import pandas as pd

import trapo_app.io_helpers as io_helpers
import trapo_app.pdf_helpers as pdf_helpers
import trapo_app.table_helpers as table_helpers


def main():
    print("Willkommen bei der Trapo App! 🐶🚐")
    print()
    print("Du kannst folgende Konsolen-Befehle benutzen:")
    print("- trapo-vergleich: Vergleicht 2 Tabellen")
    print("- trapo-extrakt: Extrahiert Daten aus den Traces Dokumenten")
    print("- trapo-traces-vergleich: Vergleicht die extrahierten Traces-Daten sie mit einer Tabelle")
    print("- trapo-traces: Benennt Traces Dokumente um")
    print("- trapo-kennzeichen: Sortiert eine Tabelle nach Entfernungen der angegeben Adressen")
    print("- trapo-kombi: Kombiniert mehrere Tabellen zu einer")
    print("- trapo-komplett: Macht alles auf einmal! Ein Start, eine Datei am Ende, alles erledigt!")


def compare():
    print("Gib als erstes den Pfad zur Datei aus dem Messenger ein, z.B. './data/chat.docx'")
    print("WICHTIG: BITTE GIB DER ADRESSENSPALTE DEN NAMEN 'KONTAKT'!")
    file1 = io_helpers.get_file_ui()
    df1 = io_helpers.read_file(file1)
    print("Gib nun den Pfad zur Datei aus PetOffice ein, z.B. './data/po.docx'")
    file2 = io_helpers.get_file_ui()
    df2 = io_helpers.read_file(file2)
    print("Vergleiche...")
    df = table_helpers.compare(df1, df2)
    path = os.path.abspath(os.path.join(".", "Trapo_Vergleich.xlsx"))
    writer = pd.ExcelWriter(path)
    df.to_excel(writer, sheet_name='Auto-Vergleich', index=False)
    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets['Auto-Vergleich'].set_column(col_idx, col_idx, column_length)
    writer.close()
    print(f"Fertig! Die Vergleichsdatei liegt unter '{path}'")


def extract():
    print("Gib den Pfad zum Ordner an, in dem alle Traces-Dokumente liegen, z.B. './data/traces'")
    path = io_helpers.get_path()
    pdfs = io_helpers.get_files(path, ".pdf")
    if len(pdfs) == 0:
        print("Keine PDFs gefunden")
        sys.exit(1)
    print("Extrahiere Informationen... Das kann etwas dauern...")
    df = pdf_helpers.extract_traces(pdfs)
    path = os.path.abspath(os.path.join(".", "Traces_Extrakt.xlsx"))
    writer = pd.ExcelWriter(path)
    df.to_excel(writer, sheet_name='Traces', index=False)
    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets['Traces'].set_column(col_idx, col_idx, column_length)
    writer.close()
    print(f"Fertig! Die Datei liegt unter '{path}'")


def compare_with_traces():
    print("Gib als erstes den Pfad zur Trapo_Vergleich-Tabelle ein, z.B. './Trapo_Vergleich.xlsx'")
    file1 = io_helpers.get_file_ui()
    df1 = io_helpers.read_file(file1)
    print("Gib nun den Pfad zu Traces_Extrakt-Tabelle ein, z.B. './Traces_Extrakt.xlsx'")
    file2 = io_helpers.get_file_ui()
    df2 = io_helpers.read_file(file2)
    print("Vergleiche...")
    df = table_helpers.compare_traces(df1, df2)
    path = os.path.abspath(os.path.join(".", "Trapo_Traces_Vergleich.xlsx"))
    writer = pd.ExcelWriter(path)
    df.to_excel(writer, sheet_name='Traces-Vergleich', index=False)
    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets['Traces-Vergleich'].set_column(col_idx, col_idx, column_length)
    writer.close()
    print(f"Fertig! Die Vergleichsdatei liegt unter '{path}'")


def rename():
    print("Gib nun den Pfad zur Trapo_Vergleich-Tabelle ein, z.B. './Trapo_Traces_Vergleich.xlsx'")
    file1 = io_helpers.get_file_ui()
    df1 = io_helpers.read_file(file1)
    print("Baue neuen Dateinamen...")
    df1, old, new = table_helpers.write_new_file_names(df1)
    print(f"Benenne {len(new)} Dateien um...")
    io_helpers.rename_files(old, new)
    print("Erstelle Kennzeichen-Ordner und ordne Traces zu...")
    io_helpers.create_folders(df1)
    io_helpers.move_files(df1)
    # print("Gib den Pfad zum Ordner an, in dem die Trapo-Stopp-Dokumente liegen, z.B. './data'")
    # path = io_helpers.get_path()
    # files = io_helpers.get_files(path, ".docx")
    # files = io_helpers.filter_stopps(files)
    # if len(files) == 0:
    #    print("Keine .docx-Dateien gefunden")
    #    sys.exit(1)
    # all_files = io_helpers.read_files(files)
    # print("Benenne Ordner um nach Stopp...")
    # TODO
    print("Fertig, alle Traces-Dateien wurden umbenannt und verschoben.")


def distance():
    print("Gib als erstes den Pfad zur Kennzeichen-Datei ein, z.B. './data/kennzeichen.xlsx'")
    file1 = io_helpers.get_file_ui()
    df1 = io_helpers.read_file(file1)
    print("Berechne Anfahrten...")
    print("Dies ist nur ein Dummy, hier passiert noch nichts.")
    path = os.path.abspath(os.path.join(".", "Trapo_Kennzeichen.xlsx"))
    print(f"Fertig! Die Datei liegt unter '{path}'")


def combine():
    print(
        "Wähle im sich gleich öffnenden Fenster alle Dateien aus, die du kombinieren willst. Kehre danach hierhin zurück.")
    files = io_helpers.get_several_files_ui()
    dfs = io_helpers.read_files(files)
    df = table_helpers.combine_dfs(dfs)
    path = os.path.abspath(os.path.join(".", "Trapo_Kombiniert.xlsx"))
    writer = pd.ExcelWriter(path)
    df.to_excel(writer, sheet_name='Kombi', index=False)
    for column in df:
        if column == "":
            continue
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets['Kombi'].set_column(col_idx, col_idx, column_length)
    writer.close()
    print(f"Fertig! Die Datei liegt unter '{path}'")


def do_all():
    print("Dies ist nur ein Dummy, hier passiert noch nichts.")


if __name__ == "__main__":
    main()
