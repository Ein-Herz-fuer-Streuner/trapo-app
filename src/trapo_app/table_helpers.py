import os.path
import re
import string
import sys
import time
from datetime import datetime

import pandas as pd
from thefuzz import fuzz as thefuzz
from rapidfuzz import process, fuzz
from trapo_app import io_helpers, math_helpers

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
    "( Schweden)": "",
    ", Schweden": "",
    "  ": " ",
    "und": "&",
    "Und": "&",
    "u.": "&",
    ",": "-",
}

german_char_map = {ord('ä'): 'ae', ord('ü'): 'ue', ord('ö'): 'oe', ord('ß'): 'ss'}


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


def clean_german(cell):
    return cell.translate(german_char_map)


def clean_name(name):
    name = name.lower()
    name = re.sub(r"box\s?(change)?!?", "", name).strip()
    return string.capwords(name)


def clean_dob(dob):
    lower = dob.lower()
    if dob == "" or dob == " " or lower == "nan":
        return ""
    # try different formats before giving up
    for fmt in (
            '%d.%m.%Y', '%d.%m.%y', '%d.%m %Y', '%d.%m %y', '%d %m.%Y', '%d %m %Y', '%d %m.%y', '%d %m %y', '%d/%m/%Y'):
        try:
            tmp = datetime.strptime(dob, fmt)
            return tmp.strftime('%d.%m.%Y')
        except ValueError:
            pass
        except TypeError:
            print("Oh oh, da ist Fehler beim Formatieren eines Geburtsdatums aufgetreten:", dob)
            print("Bitte korrigieren und diesen Befehl neustarten.")
            sys.exit(1)
    # none matched
    print("Oh oh, da ist ein Datum nicht richtig formatiert:", dob)
    print("Bitte korrigieren und diesen Befehl neustarten.")
    sys.exit(1)


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
    finished = re.sub(r'\W+$', '', finished)
    finished = finished.strip()
    return finished


def clean_location(loc):
    return loc.replace("\n", " ")


def chip_to_str(chip):
    return str(chip).replace(".0", "")


def prep_work(df1, df2):
    # avoid bs results
    if not "Kontakt" in list(df1.columns):
        print("Bitte die Spalte mit den Adoptantendaten in 'Kontakt' umbenennen und diesen Befehl neustarten.")
        sys.exit(1)

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
    # avoid mix up if two animals have the same name
    df1 = df1.sort_values('Name')
    df1 = df1.reset_index(drop=True)
    df2 = df2.sort_values('Name')
    df2 = df2.reset_index(drop=True)
    return df1, df2


def compare_contact(cont1, cont2):
    reasons = []
    is_same = True
    split1 = cont1.split(",")
    split2 = cont2.split(",")
    if len(split1) != len(split2):
        # TODO if Tierheim -> make ok
        return False, ["Länge"]
    if len(split1) > 3 and "Tierheim" in split1[0]:
        del split1[1]
        del split2[1]
    for i, (s1, s2) in enumerate(zip(split1, split2)):
        s1 = s1.strip()
        s2 = s2.strip()
        match i:
            case 0:
                if thefuzz.ratio(s1, s2) < 99:
                    # Tom Man and Tina Woman would match
                    # Tom and Martina Man would not
                    # TODO split and check if all in other
                    if not s1 in s2 and not s2 in s1:
                        is_same = False
                        reasons.append("Name")
                        continue
            case 1:
                ss1 = re.match(r"(\D+)\s*(\d+)?", s1)
                if not ss1:
                    print("Chat-Datei: Konnte Straße", s1, "nicht matchen")
                    return False, ["Regex-Fehler"]
                ss1 = ss1.group(1).strip()
                ss2 = re.match(r"(\D+)\s*(\d+)?", s2)
                if not ss2:
                    print("PO-Datei: Konnte Straße", s2, "nicht matchen")
                    return False, ["Regex-Fehler"]
                ss2 = ss2.group(1).strip()
                if len(ss1) != len(ss2):
                    is_same = False
                    reasons.append("Straße")
                    continue
                if len(ss1) == 1:
                    continue
                no1 = re.search(r'\d+', s1)
                no2 = re.search(r'\d+', s2)
                # Street has to match
                if ss1 != ss2:
                    is_same = False
                    reasons.append("Straße")
                    continue
                # Hausnummer has to match (9a vs 9 is ok)
                if not no1 or not no2:
                    is_same = False
                    reasons.append("HNr")
                    continue
                if no1.group(0) != no2.group(0):
                    is_same = False
                    reasons.append("HNr")
                    continue

            case 2:
                ss1 = s1.split(" ")
                ss2 = s2.split(" ")
                plz1 = ss1[0].strip()
                plz2 = ss2[0].strip()
                city1 = ss1[1].strip()
                city2 = ss2[1].strip()
                # Post code has to be an exact match
                if plz1 != plz2:
                    is_same = False
                    reasons.append("Plz")
                # Rothenburg & Rothenburg ob der Tauber have to match
                short_city1 = city1.split(" ")
                short_city2 = city2.split(" ")
                if short_city1[0].strip() != short_city2[0].strip():
                    is_same = False
                    reasons.append("Stadt")
                    continue
            case _:
                print("Oh oh, hier ist was schief gelaufen, bitte überprüfen und neustarten:", cont1, cont2)
                sys.exit(0)
    return is_same, reasons


def match_pet(row, df):
    name = row["Name"]
    count = 0
    # first check by name
    for name_comp in df["Name"].values:
        if name == name_comp:
            count += 1
    if count == 1:
        return df[df["Name"] == name].iloc[0]
    # if more than one or none, check by chip
    chip = row["Chip"]
    for chip_comp in df["Chip"].values:
        if chip == chip_comp:
            return df[df["Chip"] == chip].iloc[0]
    return pd.Series()


def compare(df1, df2):
    df1, df2 = prep_work(df1, df2)

    differences = []
    try:
        # Check df1 against df2
        for index, row in df1.iterrows():
            matched_row = match_pet(row, df2)
            if not matched_row.empty:
                diffs = []
                if row["Chip"] != matched_row["Chip"]:
                    diffs.append(
                        f"Chip: {row['Chip'] if row['Chip'] != '' else '\'\''} \u2192 {matched_row['Chip'] if matched_row['Chip'] != '' else '\'\''}")
                if row["DOB"] != matched_row["DOB"]:
                    diffs.append(
                        f"DOB: {row['DOB'] if row['DOB'] != '' else '\'\''} \u2192 {matched_row['DOB'] if matched_row['DOB'] != '' else '\'\''}")
                is_same, reasons = compare_contact(row["Kontakt"], matched_row["Kontakt"])
                if not is_same:
                    diffs.append(
                        f"Kontakt ({", ".join(reasons)}): {row['Kontakt'] if row['Kontakt'] != '' else '\'\''} \u2192 {matched_row['Kontakt'] if matched_row['Kontakt'] != '' else ''}")

                difference = ", ".join(diffs) if diffs else "\u2713"
                differences.append({"Name": row["Name"], "Ort": row["Ort"], "Chip": row["Chip"], "DOB": row["DOB"],
                                    "Kontakt": row["Kontakt"],
                                    "Differenz (Chat \u2192 PetOffice)": difference})
            else:
                differences.append(
                    {"Name": row["Name"], "Chip": row["Chip"], "Kontakt": row["Kontakt"],
                     "Differenz (Chat \u2192 PetOffice)": "Fehlt in PetOffice-Datei"})
    except KeyError as err:
        print("Spalte nicht in Tabelle gefunden:", err)
        sys.exit(0)

    try:
        # Compare df2 against df1 for missing names
        for index, row in df2.iterrows():
            matched_row = match_pet(row, df1)
            if matched_row.empty:
                differences.append(
                    {"Name": row["Name"], "Ort": "?", "Chip": row["Chip"], "DOB": row["DOB"], "Kontakt": row["Kontakt"],
                     "Differenz (Chat \u2192 PetOffice)": "Fehlt in Chat-Datei"})
    except KeyError as err:
        print("Spalte nicht in Tabelle gefunden:", err)
        sys.exit(0)
    # Result DataFrame
    df_result = pd.DataFrame(differences)
    df_result = df_result.sort_values(['Name', 'Chip'])
    return df_result


def get_diff_col_name(cols):
    col = ""
    for c in cols:
        if "Differenz" in c:
            return c
    return col


def compare_traces(df1, df2):
    df1["Chip"] = df1["Chip"].apply(chip_to_str)
    df2["Chip"] = df2["Chip"].apply(chip_to_str)
    differences = []
    diff_column = get_diff_col_name(list(df1.columns))
    if diff_column == "":
        print("Kein Spaltennamen mit 'Differenz' gefunden")
        sys.exit(0)
    # Check df1 against df2
    # df1: Name, Ort, Chip, DOB, Kontakt, Differenz (Chat \u2192 PetOffice)
    # df2: Datei, Intra, Chip, Kontakt
    try:
        for index, row in df1.iterrows():
            matched_row = pd.Series()
            if row["Chip"] in df2["Chip"].values:
                matched_row = df2[df2["Chip"] == row["Chip"]].iloc[0]
            elif row[diff_column] in df2["Chip"].values:
                matched_row = df2[df2["Chip"] == row[diff_column]].iloc[0]
            if not matched_row.empty:
                diffs = []
                if not compare_contact(row["Kontakt"], matched_row["Kontakt"]):
                    diffs.append(
                        f"Kontakt: {row['Kontakt'] if row['Kontakt'] != '' else '\'\''} \u2192 {matched_row['Kontakt'] if matched_row['Kontakt'] != '' else ''}")

                difference = ", ".join(diffs) if diffs else "\u2713"
                differences.append({"Name": row["Name"], "Ort": row["Ort"], "Chip": row["Chip"], "DOB": row["DOB"],
                                    "Kontakt": row["Kontakt"], "Intra": matched_row["Intra"],
                                    "Datei": matched_row["Datei"],
                                    "Kennzeichen": matched_row["Kennzeichen"],
                                    "Differenz (Chat \u2192 Traces)": difference})
            else:
                differences.append(
                    {"Name": row["Name"], "Chip": str(row["Chip"]), "Kontakt": row["Kontakt"], "Intra": "?",
                     "Datei": "?", "Kennzeichen": "?",
                     "Differenz (Chat \u2192 Traces)": "Fehlt in Traces-Dokumenten"})
    except KeyError as err:
        print("Spalte nicht in Tabelle gefunden:", err)
        sys.exit(0)
    # Compare df2 against df1 for missing names
    try:
        for index, row in df2.iterrows():
            if row["Chip"] not in df1["Chip"].values and row["Chip"] not in df1[diff_column].values:
                differences.append(
                    {"Name": "?", "Ort": "?", "Chip": str(row["Chip"]), "DOB": "?", "Kontakt": row["Kontakt"],
                     "Intra": row["Intra"], "Datei": row["Datei"], "Kennzeichen": row["Kennzeichen"],
                     "Differenz (Chat \u2192 Traces)": "Fehlt in Chat-Datei"})
    except KeyError as err:
        print("Spalte nicht in Tabelle gefunden:", err)
        sys.exit(0)
    # Result DataFrame
    df_result = pd.DataFrame(differences)
    df_result = df_result.sort_values(['Name', 'Chip'])
    return df_result


def build_file_name(df):
    files_old = df.drop_duplicates(subset="Datei", keep="first")
    files_old = files_old["Datei"].values.tolist()
    files_old = [x for x in files_old if not x in ["", "?"]]
    files_new = []
    df["Datei neu"] = ""
    for file in files_old:
        animals = []
        contact = ""
        intra = ""
        # get all animals and the Contact info
        for index, row in df.iterrows():
            if row["Intra"] != "?" and row["Intra"] in file:
                animals.append(row["Name"])
                if contact == "":
                    contact = row["Kontakt"]
                    intra = row["Intra"]

        if contact == "" or len(animals) == 0 or intra == "":
            continue
        all_animals = "_".join(animals)
        name = contact.split(",")[0]
        # Intranummer_Tiername_Vorname Nachname
        new = f"{intra}_{all_animals}_{name}.pdf"
        files_new.append(new)
    # save new file name for moving
    df["Datei neu"] = df["Datei"].apply(match_old_new, args=(files_old, files_new))
    return df, files_old, files_new


def match_old_new(cell, old, new):
    for o, n in zip(old, new):
        if cell == o:
            return n
    return ""


def write_new_file_names(df):
    df, old, new = build_file_name(df)
    # ( nice to have : Dateiname ohne ä,ü,ö und ß -> ersetzen durch ue, ae oe, ss)
    new = [clean_german(x) for x in new]
    df["Datei neu"] = df["Datei neu"].apply(clean_german)
    return df, old, new


def combine_dfs(dfs):
    dfs = [df.reset_index(drop=True) for df in dfs]
    df = pd.DataFrame()
    try:
        df = pd.concat(dfs, ignore_index=True)
    except Exception as err:
        print("Ups, da passt was nicht", err)
        sys.exit(0)
    return df


def find_stopp_for_plate(files, dfs, folders):
    results = []
    seen = []
    # go through folders
    for plate in folders:
        count = 0
        found = False
        # go through zip files, dfs
        for file, df in zip(files, dfs):
            if found:
                break
            # clean all names in df
            df['Name'] = df['Name'].apply(clean_name)
            # extract stop location
            stop = extract_stop(file)
            if stop == "":
                break
            # get all file names of folder
            files_ = io_helpers.get_all_files_from_folder(os.path.join(".", plate) + "/*.pdf")
            # go through all names in the word document
            for name in df["Name"]:
                if found:
                    break
                for file_ in files_:
                    if name in file_ and not stop in seen:
                        # in case of double names, wait for 5 matches
                        count += 1
                        if count > 4:
                            # append tuple (folder, stop, file)
                            results.append((plate, stop, file))
                            seen.append(stop)
                            found = True
                            break
    return results


def extract_stop(file):
    file = io_helpers.simple_normalize(file)
    if "nord" in file:
        return "NORD"
    if "mitte" in file:
        return "MITTE"
    if "südwest" in file or "suedwest" in file or "sudwest" in file:
        return "SÜDWEST"
    if "süd" in file or "sued" in file or "sud" in file:
        return "SÜD"
    return ""

def replace_empty(cell):
    return '?' if pd.isna(cell) or str(cell).strip() == '' else cell

def shrink_tables(dfs):
    results = []
    for df in dfs:
        wanted_cols = ['Photo', 'Name', 'Kontakt', 'Treffpunkt']
        tab_cols = list(df.columns)
        got_cols = []
        for col in tab_cols:
            if col in wanted_cols:
                got_cols.append(col)
        new_df = df[got_cols].copy()
        new_df["Treffpunkt"] = new_df["Treffpunkt"].apply(replace_empty)
        results.append(new_df)
    return results

def clean_plate_dfs(dfs):
    results = []
    for df in dfs:
        # clean input
        df['Name'] = df['Name'].apply(clean_name)
        df['Kontakt'] = df['Kontakt'].apply(clean_contact)
        if 'Ort' in df.columns:
            df["Ort"] = df["Ort"].apply(clean_location)
        df = df.sort_values('Name')
        df = df.reset_index(drop=True)
        results.append(df)

    return results

def extract_all_plates(text):
    text = text.upper()
    plates = []
    matched_spans = []

    def is_overlapping(span):
        return any(max(start, span[0]) < min(end, span[1]) for start, end in matched_spans)

    # --- German plate pattern with optional 'E'
    german_pattern = re.compile(r'\b([A-ZÄÖÜ]{1,3})\s*[- :]?\s*([A-Z]{1,2})\s*[- ]?\s*(\d{1,4})(\s*E)?\b')
    for match in german_pattern.finditer(text):
        if not is_overlapping(match.span()):
            city, letters, numbers, e = match.groups()
            e = ' E' if e else ''
            plates.append(f"{city}-{letters} {numbers}{e}")
            matched_spans.append(match.span())

    # --- Austrian pattern (simplified)
    austrian_pattern = re.compile(r'\b([A-ZÄÖÜ]{1,3})\s*(\d{1,4})([A-Z]{1,2})\b')
    for match in austrian_pattern.finditer(text):
        if not is_overlapping(match.span()):
            district, numbers, letters = match.groups()
            plate = f"{district}-{numbers}{letters} (AT)"
            plates.append(plate)
            matched_spans.append(match.span())

    # --- Swiss pattern
    swiss_pattern = re.compile(r'\b([A-Z]{2})\s*(\d{1,6})\b')
    for match in swiss_pattern.finditer(text):
        if not is_overlapping(match.span()):
            canton, numbers = match.groups()
            plate = f"{canton}-{numbers} (CH)"
            plates.append(plate)
            matched_spans.append(match.span())

    if len(plates) == 0:
        return ""
    return ", ".join(plates)


def find_plate(name, names_only, row_lookup, df):
    #print(name)
    norm = name.lower()
    match, score, _ = process.extractOne(
        norm,
        names_only,
        scorer=fuzz.ratio,
        score_cutoff=80
    ) or (None, None, None)

    if match is None:
        return ""  # nothing above cutoff
    row = row_lookup[match]  # map back to df2 row
    r = df.iloc[row]
    return extract_all_plates(r["Kennzeichen"])


def add_plates(dfs, df1):
    results = []
    df1.columns = (
        df1.columns
        .str.replace(r'^Name.*', 'Name', regex=True)
        .str.replace(r'^Kennzeichen.*', 'Kennzeichen', regex=True,flags=re.DOTALL)
    )
    # Step 1: Extract potential names from df2
    name_index_pairs = []

    for idx, row in df1.iterrows():
        text = str(row['Name'])
        # Find words that look like names (basic regex for words with letters only, min 2 chars)
        candidates = re.findall(r'\b[A-ZÄÖÜa-zäöüß]{2,}\b', text)
        for name in candidates:
            name_index_pairs.append((name.strip().lower(), idx))
    #print(name_index_pairs)
    names_only = [n for n, _ in name_index_pairs]  # list[str]
    row_lookup = {n: row for n, row in name_index_pairs}
    for df in dfs:
        # Compute the new column values using apply
        new_column_values = df["Name"].apply(find_plate, args=(names_only, row_lookup ,df1))
        # Insert the new column at position 2
        df.insert(loc=2, column='Kennzeichen', value=new_column_values)
        results.append(df)
    return results

def add_distance(dfs, stopps):
    results = []
    print("Anfragen an OpenStreemMap sind begrenzt auf eine pro Sekunde, dieser Schritt kann dauern...")
    for i, df in enumerate(dfs):
        print(f"Berechne Entfernungen für Dokument {i+1} von {len(dfs)}...")
        # Compute the new column values using apply
        distances = []
        for _, row in df.iterrows():
            distances.append(math_helpers.calculate_distance(row, stopps))
            time.sleep(1)
        # Insert the new column at position 5
        df.insert(loc=5, column='Entfernung', value=distances)
        # sort by distance PER LOCATION
        df = df.sort_values(by=['Treffpunkt', 'Entfernung'], ascending=[True,False])
        df = df.reset_index(drop=True)
        # insert extra header
        df = insert_headers(df)
        results.append(df)

    return results

def insert_headers(df):
    # Prepare new rows with repeated headers and numbering
    rows = []
    #columns = ['Nr.'] + df.columns.tolist()  # Add new 'Row Number' column
    current_meeting = "Treffpunkt"
    row_counter = 0

    for idx, row in df.iterrows():
        if (idx == 0 and row['Treffpunkt'] != "Treffpunkt") or (not row['Treffpunkt'] in current_meeting and not current_meeting in row['Treffpunkt']):
            current_meeting = row['Treffpunkt']
            row_counter = 0  # Reset counter for new group
            # Insert header row (with 'Nr.' as header too)
            rows.append({'Nr.': 'Nr.', **{col: col for col in df.columns}})
        row_counter += 1
        rows.append({'Nr.': row_counter, **row.to_dict()})

    # Convert to new DataFrame
    new_df = pd.DataFrame(rows)
    new_df.columns = new_df.iloc[0]
    new_df = new_df[1:].reset_index(drop=True)
    return new_df