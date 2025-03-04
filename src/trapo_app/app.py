#!/bin/python3
import io_helpers
import table_helpers

def main():
    print("Willkommen bei der Trapo App!")
    print()
    print("Du kannst folgende Konsolen-Befehle benutzen:")
    print("- trapo-vergleich: Vergleicht 2 Tabellen")
    print("- trapo-traces: Benennt Traces Dokumente um")
    print("- trapo-traces-vergleich: Vergleicht Traces Dokumente mit einer Tabelle")
    print("- trapo-kennzeichen: Sortiert eine Tabelle nach Entfernungen der angegeben Adressen")

def compare():
    #print("Gib als erstes den Pfad zur Datei aus dem Messenger ein, z.B. './data/chat.docx'")
    # file1 = io_helpers.get_file()
    df1 = io_helpers.read_file("data/chat2.docx")
    #print("Gib nun den Pfad zur Datei aus PetOffice ein")
    # file2 = io_helpers.get_file()
    df2 = io_helpers.read_file("data/docx2.docx")
   # print("Vergleiche...")
    table_helpers.compare(df1, df2)

def rename():
    print("Benenne Traces um...")
    print("Dies ist nur ein Dummy, hier passiert noch nichts.")

def compare_with_traces():
    print("Benenne Traces um...")
    print("Dies ist nur ein Dummy, hier passiert noch nichts.")

def distance():
    print("Berechne Anfahrten...")
    print("Dies ist nur ein Dummy, hier passiert noch nichts.")

if __name__ == "__main__":
    #main()
    compare()