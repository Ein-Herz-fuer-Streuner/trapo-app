#!/bin/python3.13
import pandas as pd

import trapo_app.io_helpers as io_helpers
import trapo_app.table_helpers as table_helpers


def main():
    print("Willkommen bei der Trapo App! 🐶🚐")
    print()
    print("Du kannst folgende Konsolen-Befehle benutzen:")
    print("- trapo-vergleich: Vergleicht 2 Tabellen")
    print("- trapo-traces: Benennt Traces Dokumente um")
    print("- trapo-traces-vergleich: Vergleicht Traces Dokumente mit einer Tabelle")
    print("- trapo-kennzeichen: Sortiert eine Tabelle nach Entfernungen der angegeben Adressen")
    print("- trapo-komplett: Macht alles aufeinmal! Ein Start, eine Datei am Ende, alles erledigt!")


def compare():
    print("Gib als erstes den Pfad zur Datei aus dem Messenger ein, z.B. './data/chat.docx'")
    print("WICHTIG: BITTE GIB DER ADRESSENSPALTE DEN NAMEN 'KONTAKT'!")
    file1 = io_helpers.get_file()
    df1 = io_helpers.read_file(file1)
    # df1 = io_helpers.read_file("..\\..\\data\\chat.docx")
    print("Gib nun den Pfad zur Datei aus PetOffice ein, z.B. './data/po.docx'")
    file2 = io_helpers.get_file()
    df2 = io_helpers.read_file(file2)
    # df2 = io_helpers.read_file("..\\..\\data\\po.docx")

    print("Vergleiche...")
    df = table_helpers.compare(df1, df2)

    writer = pd.ExcelWriter('./Trapo_Vergleich.xlsx')
    df.to_excel(writer, sheet_name='Auto-Vergleich', index=False)
    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets['Auto-Vergleich'].set_column(col_idx, col_idx, column_length)
    writer.close()
    print("Fertig! Die Vergleichsdatei liegt unter './Trapo_Vergleich.xlsx'")

def rename():
    print("Gib den Pfad zum Ordner an, in dem alle Traces-Dokumente liegen, z.B. './data/traces'")
    print("Dies ist nur ein Dummy, hier passiert noch nichts.")
    print("Fertig, alle Traces-Dateien wurden umbennant.")


def compare_with_traces():
    print("Gib als erstes den Pfad zur Trapo-Tabelle ein, z.B. './data/Trapo_Vergleich.docx'")
    file1 = io_helpers.get_file()
    df1 = io_helpers.read_file(file1)
    print("Gib nun den Pfad zum Ordner an, in dem alle Traces-Dokumente liegen, z.B. './data/traces'")
    path = io_helpers.get_path()
    print("Suche Tracesnummern. Das kann dauern. Bitte warten...")

    print("Fertig. Vergleiche jetzt mit der Tabelle...")

    print("Dies ist nur ein Dummy, hier passiert noch nichts.")
    print("Fertig! Die Vergleichsdatei liegt unter './Trapo_Vergleich_Traces.xlsx'")


def distance():
    print("Gib als erstes den Pfad zur Kennzeichen-Datei ein, z.B. './data/kennzeichen.xlsx'")
    file1 = io_helpers.get_file()
    df1 = io_helpers.read_file(file1)
    print("Berechne Anfahrten...")
    print("Dies ist nur ein Dummy, hier passiert noch nichts.")
    print("Fertig! Die Datei liegt unter './Kennzeichen_sortiert.xlsx'")


def do_all():
    print("Dies ist nur ein Dummy, hier passiert noch nichts.")

if __name__ == "__main__":
    main()