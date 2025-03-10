import re
import string

import pandas as pd
from dateutil import parser
from thefuzz import fuzz

phone_regex = r'[\+0-9\/\s-]{8,}'
email_regex = r'[a-zA-Z0-9._%+-]+@?[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
sim_thresh = 97

replace_dict = {
    ", Deutschland": "",
    ",Deutschland": "",
    "(Deutschland)": "",
    "( Deutschland)": "",
    ", Schweiz": "",
    ",Schweiz": "",
    "(Schweiz)": "",
    "( Schweiz)": "",
    ", Österreich": "",
    ",Österreich": "",
    "(Österreich)": "",
    "( Österreich)": "",
    "  ": " ",
}

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
    parts = contact.split('\n')[:4]
    result = []
    for part in parts:
        part = part.title()
        for key, value in replace_dict.items():
            part = part.replace(key, value)

        part = re.sub(r'(Tel\.?|Telefon|Mobil)?:?\s?[+0-9/\s-]{8,}', "", part)
        part = re.sub(r'(Str\b\.?|Strasse)', "Straße", part)
        part = re.sub(r'(str\b\.?|strasse)', "straße", part)
        part = part.strip()
        # remove duplicates e.g. 2x streetname + no
        part = " ".join(dict.fromkeys(part.split(" ")))
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

def compare_contact(cont1, cont2):
    split1 = cont1.split(",")
    split2 = cont2.split(",")
    if len(split1) != len(split2):
        return False
    for i, (s1, s2) in enumerate(zip(split1, split2)):
        s1 = s1.strip()
        s2 = s2.strip()
        match i:
            case 0:
                if fuzz.ratio(s1, s2) < 99:
                    # Tom Man and Tina Woman would match
                    # Tom and Martina Man would not
                    if not s1 in s2 and not s2 in s1:
                        return False
            case 1:
                ss1 = s1.split(" ")
                ss2 = s2.split(" ")
                str1 = " ".join(ss1[:-1]).strip()
                str2 = " ".join(ss2[:-1]).strip()
                if len(ss1) != len(ss2):
                    return False
                if len(ss1) == 1:
                    continue
                no1 = re.search(r'\d+', s1)
                no2 = re.search(r'\d+', s2)
                # Street has to match
                if str1 != str2:
                    return False
                # Hausnummer has to match (9a vs 9 is ok)
                if not no1 or not no2:
                    return False
                if no1.group(0) != no2.group(0):
                    return False

            case 2:
                ss1 = s1.split(" ")
                ss2 = s2.split(" ")
                plz1 =  ss1[0].strip()
                plz2 =  ss2[0].strip()
                city1 = ss1[1].strip()
                city2 = ss2[1].strip()
                # Post code has to be an exact match
                if plz1 != plz2:
                    return False
                # Rothenburg & Rothenburg ob der Tauber have to match
                short_city1 = city1.split(" ")
                short_city2 = city2.split(" ")
                if short_city1[0].strip() != short_city2[0].strip():
                    return False
            case _:
                print("Oh oh, hier ist was schief gelaufen")
    return True

def compare(df1, df2):
    df1, df2 = prep_work(df1, df2)

    differences = []
    # Check df1 against df2
    for index, row in df1.iterrows():
        if row["Name"] in df2["Name"].values:
            matched_row = df2[df2["Name"] == row["Name"]].iloc[0]
            diffs = []
            if not compare_contact(row["Kontakt"], matched_row["Kontakt"]):
                        diffs.append(
                            f"Kontakt: {row['Kontakt'] if row['Kontakt'] != '' else '\'\''} \u2192 {matched_row['Kontakt'] if matched_row['Kontakt'] != '' else ''}")

            if row["DOB"] != matched_row["DOB"]:
                diffs.append(
                    f"DOB: {row['DOB'] if row['DOB'] != '' else '\'\''} \u2192 {matched_row['DOB'] if matched_row['DOB'] != '' else '\'\''}")

            if row["Chip"] != matched_row["Chip"]:
                diffs.append(
                    f"Chip: {row['Chip'] if row['Chip'] != '' else '\'\''} \u2192 {matched_row['Chip'] if matched_row['Chip'] != '' else '\'\''}")

            difference = ", ".join(diffs) if diffs else "\u2713"
            differences.append({"Name": row["Name"], "Ort": row["Ort"], "Chip": row["Chip"], "DOB": row["DOB"],
                                "Kontakt": row["Kontakt"],
                                "Differenz (Chat \u2192 PetOffice)": difference})
        else:
            differences.append(
                {"Name": row["Name"], "Chip": row["Chip"], "Kontakt": row["Kontakt"], "Differenz (Chat \u2192 PetOffice)": "Fehlt in PetOffice-Datei"})

    # Compare df2 against df1 for missing names
    for index, row in df2.iterrows():
        if row["Name"] not in df1["Name"].values:
            differences.append(
                {"Name": row["Name"], "Ort": "?", "Chip": row["Chip"], "DOB": row["DOB"], "Kontakt": row["Kontakt"],
                 "Differenz (Chat \u2192 PetOffice)": "Fehlt in Chat-Datei"})

    # Result DataFrame
    df_result = pd.DataFrame(differences)
    df_result = df_result.sort_values('Name')
    return df_result
