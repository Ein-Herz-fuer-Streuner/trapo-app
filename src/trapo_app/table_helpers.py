import pandas as pd
import string
from dateutil import parser

def clean_compare_table(df):
    df = df.rename(columns= {"Daten ES/PS/PSO": "Kontakt", "Microchip": "Chip"}, errors= "ignore")
    wanted_cols = ['Name', 'Ort', 'Chip', 'Kontakt', 'DOB']
    tab_cols = list(df.columns)
    got_cols = []
    for col in tab_cols:
        if col in wanted_cols:
            got_cols.append(col)

    new_df = df[got_cols]
    new_df = new_df.sort_values('Name')
    return new_df

def clean_name(name):
    name = name.lower().replace("box change", "").replace("box delivery", "").strip()
    return string.capwords(name)

def clean_dob(dob):
    if dob == "":
        return ""
    return parser.parse(dob, dayfirst=True).strftime("%d.%m.%Y")

def clean_contact(contact):
    parts = contact.split('\n')
    result = []
    for part in parts:
        part = part.title()
        part = part.replace(", Deutschland", "")
        part = part.replace(",Deutschland", "")
        part = part.replace("(Deutschland)", "")
        part = part.replace("( Deutschland)", "")
        part = part.replace(", Schweiz", "")
        part = part.replace(",Schweiz", "")
        part = part.replace("(Schweiz)", "")
        part = part.replace("( Schweiz)", "")
        part = part.replace(", Österreich", "")
        part = part.replace(",Österreich", "")
        part = part.replace("(Österreich)", "")
        part = part.replace("( Österreich)", "")
        if "@" in part or part.isdigit() or "Mobil" in part or "Telefon" in part or "+4" in part:
            continue
        result.append(part)
    return ", ".join(result)

def clean_location(loc):
    return loc.replace("\n", " ")

def prep_work(df1, df2):
    df1 = clean_compare_table(df1)
    df1['Name'] = df1['Name'].apply(clean_name)
    df1['DOB'] = df1['DOB'].apply(clean_dob)
    df1['Kontakt'] = df1['Kontakt'].apply(clean_contact)
    if 'Ort' in df1.columns:
        df1["Ort"] = df1["Ort"].apply(clean_location)

    df2 = clean_compare_table(df2)
    df2['Name'] = df2['Name'].apply(clean_name)
    df2['DOB'] = df2['DOB'].apply(clean_dob)
    df2['Kontakt'] = df2['Kontakt'].apply(clean_contact)
    if 'Ort' in df2.columns:
        df2["Ort"] = df2["Ort"].apply(clean_location)
    return df1, df2

def compare(df1, df2):
    df1, df2 = prep_work(df1, df2)
    len_df1 = df1.size
    len_df2 = df2.size
    max_len = max(len_df1, len_df2)
    min_len = min(len_df1, len_df2)
    if len_df1 != len_df2:
        print("Oh oh, die Listen haben nicht die gleiche Länge!")
    df1["Differenzen"] = ""
    # TODO actual comparing
    return df1