import pandas as pd
import string
from dateutil import parser
import re

phone_regex = r'[\+0-9\/\s-]{8,}'
email_regex = r'[a-zA-Z0-9._%+-]+@?[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'

def clean_compare_table(df):
    df = df.rename(columns={"Daten ES/PS/PSO": "Kontakt", "Microchip": "Chip"}, errors="ignore")
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
        part = part.strip()
        match_mail = re.search(email_regex, part)
        match_phone = re.search(phone_regex, part)
        if part == "" or part.isdigit() or "Mail" in part or "Tel." in part or "Telefon" in part or "Mobil" in part or "Vvk" in part or match_mail or match_phone:
            continue
        result.append(part)
    finished = ", ".join(result)
    finished = finished.replace("  ", " ")
    finished = finished.replace(",,", ",")
    finished = finished.strip()
    return finished


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

    differences = []
    # Check df1 against df2
    for index, row in df1.iterrows():
        if row["Name"] in df2["Name"].values:
            matched_row = df2[df2["Name"] == row["Name"]].iloc[0]
            diffs = []
            if row["Kontakt"] != matched_row["Kontakt"]:
                diffs.append(f"Kontakt: {row['Kontakt'] if row["Kontakt"] != "" else "''"} -> {matched_row['Kontakt'] if matched_row["Kontakt"] != "" else "''"}")
            if row["DOB"] != matched_row["DOB"]:
                diffs.append(f"DOB: {row['DOB'] if row["DOB"] != "" else "''"} -> {matched_row['DOB'] if matched_row["DOB"] != "" else "''"}")
            if row["Chip"] != matched_row["Chip"]:
                diffs.append(f"Chip: {row['Chip'] if row["Chip"] != "" else "''"} -> {matched_row['Chip'] if matched_row['Chip'] != "" else "''"}")

            difference = "\n".join(diffs) if diffs else "\u2705"
            differences.append({"Name": row["Name"], "Ort": row["Ort"], "Chip": row["Chip"], "DOB": row["DOB"],  "Kontakt": row["Kontakt"],
                                "Differenz": difference})
        else:
            differences.append(
                {"Name": row["Name"], "Chip": row["Chip"], "Kontakt": row["Kontakt"], "Differenz": "Fehlt in Datei 2"})

    # Compare df2 against df1 for missing names
    for index, row in df2.iterrows():
        if row["Name"] not in df1["Name"].values:
            differences.append({"Name": row["Name"], "Ort": "?", "Chip": row["Chip"], "DOB": row["DOB"], "Kontakt": row["Kontakt"],
                                "Differenz": "Fehlt in Datei 1"})

    # Result DataFrame
    df_result = pd.DataFrame(differences)
    df_result = df_result.sort_values('Name')
    return df_result
