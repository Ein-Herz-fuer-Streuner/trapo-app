"""DataFrame-Logik: Tabellen bereinigen, vergleichen, Kennzeichen zuordnen und Entfernungen berechnen."""
import os.path
import re
import string
from datetime import datetime

import pandas as pd
from rapidfuzz import fuzz, process

from trapo_app import io_helpers, math_helpers
from trapo_app.errors import TrapoError

PHONE_REGEX = re.compile(r'[\+0-9\/\s-]{8,}')
EMAIL_REGEX = re.compile(r'[a-zA-Z0-9._%+-]+@?[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
PHONE_PREFIX_REGEX = re.compile(r'(Tel\.?|Telefon|Mobil)?:?\s?[+0-9/\s-]{8,}')
BOX_CHANGE_REGEX = re.compile(r"box\s?(change)?!?", re.I)
NON_CONTACT_MARKERS = ("Mail", "Tel.", "Telefon", "Mobil", "Vvk")

_COUNTRIES = "Deutschland|Schweiz|Österreich|Schweden"
# Applied in this order to every line of a contact cell
CONTACT_REPLACEMENTS = [
    (re.compile(rf"\(\s?(?:{_COUNTRIES})\)|,\s?(?:{_COUNTRIES})"), ""),
    (re.compile(r" {2,}"), " "),
    (re.compile(r"\bund\b", re.I), "&"),
    (re.compile(","), "-"),
]
STREET_REPLACEMENTS = [
    (re.compile(r'(Str\b\.?|Strasse)'), "Straße"),
    (re.compile(r'(str\b\.?|strasse)'), "straße"),
]

GERMAN_CHAR_MAP = {ord('ä'): 'ae', ord('ü'): 'ue', ord('ö'): 'oe', ord('ß'): 'ss'}
DOB_FORMATS = ('%d.%m.%Y', '%d.%m.%y', '%d.%m %Y', '%d.%m %y', '%d %m.%Y', '%d %m %Y', '%d %m.%y', '%d %m %y',
               '%d/%m/%Y')
COMPARE_COLUMNS = ('Name', 'Ort', 'Chip', 'Kontakt', 'DOB')
DISTANCE_COLUMNS = ('Photo', 'Name', 'Kontakt', 'Treffpunkt')

CHAT_DIFF_COLUMN = "Differenz (Chat → PetOffice)"
TRACES_DIFF_COLUMN = "Differenz (Chat → Traces)"
OK_MARK = "✓"
ARROW = "→"

TRANSLATE_MAP = {
    "No": "Nr Crt",
    "Specie": "Specie",
    "Microchip": "Microcip",
    "Pass A": "Pasaport",
    "Health Book": "Nr carnet\nsanatate",
    "Age": "Varsta\n(in luni)",
    "Name": "Nume",
    "kg": "Kg",
    "Sex/\ncastr": "Sex\n(Femela/Mascul/Castrat)",
    "TRACES": "Importator\nNume, Adresa\ntrace nr.",
    "Kontakt": "Adoptator\nNume, Adresa\nTelefon",
    "Data\nchip": "Data microcip",
    "Data\nrabic": "Data rabic",
}

RO_HEADERS = [
    "Nr Crt", "Specie", "Microcip", "Pasaport", "Nr carnet\nsanatate", "Varsta\n(in luni)", "Nume", "Kg",
    "Sex\n(Femela/Mascul/Castrat)", "Importator\nNume, Adresa\ntrace nr.", "Adoptator\nNume, Adresa\nTelefon",
    "Nr auto iesire", "Data microcip", "Data rabic",
    "Locul de origine Nr de Inregistrare\nsau de autorizare unic al unitatii\nde origine a animalelor",
    "Nr inmatriculare\nAuto Sosire", "Nr proces verbal\ndezinfectie",
]


# ──────────────────────────────────────────────
#  Zellen bereinigen
# ──────────────────────────────────────────────

def clean_german(cell):
    """Ersetzt ä, ö, ü und ß (für Dateinamen)."""
    return cell.translate(GERMAN_CHAR_MAP)


def clean_name(name):
    name = BOX_CHANGE_REGEX.sub("", name.lower()).strip()
    return string.capwords(name)


def clean_name_with_box_change(name):
    name = name.lower()
    if BOX_CHANGE_REGEX.match(name):
        name = BOX_CHANGE_REGEX.sub("", name).strip() + " (Box Change)"
    else:
        name = name.strip()
    return string.capwords(name)


def clean_dob(dob):
    """Vereinheitlicht ein Geburtsdatum zu TT.MM.JJJJ; leere Werte bleiben leer."""
    if dob.strip() == "" or dob.lower() == "nan":
        return ""
    for fmt in DOB_FORMATS:
        try:
            return datetime.strptime(dob, fmt).strftime('%d.%m.%Y')
        except ValueError:
            pass
    raise TrapoError(f"Oh oh, da ist ein Datum nicht richtig formatiert: {dob}\n"
                     "Bitte korrigieren und diesen Befehl neustarten.")


def _is_contact_noise(part):
    return (part == "" or part.isdigit()
            or any(marker in part for marker in NON_CONTACT_MARKERS)
            or EMAIL_REGEX.search(part) or PHONE_REGEX.search(part))


def clean_contact(contact):
    """Reduziert eine Kontaktzelle auf 'Name, Straße Nr, PLZ Ort' (ohne Telefon/Mail)."""
    parts = []
    for line in contact.split('\n')[:4]:
        line = line.title()
        for pattern, replacement in CONTACT_REPLACEMENTS:
            line = pattern.sub(replacement, line)
        line = PHONE_PREFIX_REGEX.sub("", line)
        for pattern, replacement in STREET_REPLACEMENTS:
            line = pattern.sub(replacement, line)
        line = line.strip()
        line = " ".join(dict.fromkeys(line.split(" ")))  # remove duplicates e.g. 2x streetname + no
        if not _is_contact_noise(line):
            parts.append(line)
    cleaned = ", ".join(parts).replace("  ", " ").replace(",,", ",")
    return re.sub(r'\W+$', '', cleaned).strip()


def clean_location(loc):
    return loc.replace("\n", " ")


def chip_to_str(chip):
    return str(chip).replace(".0", "")


# ──────────────────────────────────────────────
#  Vergleich Chat <-> PetOffice
# ──────────────────────────────────────────────

def clean_compare_table(df):
    """Behält nur die Vergleichsspalten und sortiert nach Name."""
    df = df.rename(columns={"Daten ES/PS/PSO": "Kontakt", "Microchip": "Chip"}, errors="ignore")
    return df[[col for col in df.columns if col in COMPARE_COLUMNS]].sort_values('Name')


def _prepare_compare_df(df):
    df = clean_compare_table(df)
    if "Kontakt" not in df.columns:
        raise TrapoError("Bitte die Spalte mit den Adoptantendaten in 'Kontakt' umbenennen "
                         "und diesen Befehl neustarten.")
    df['Name'] = df['Name'].apply(clean_name)
    df['DOB'] = df['DOB'].apply(clean_dob)
    df['Kontakt'] = df['Kontakt'].apply(clean_contact)
    if 'Ort' in df.columns:
        df["Ort"] = df["Ort"].apply(clean_location)
    # sorted by name to avoid mix ups if two animals have the same name
    return df.sort_values('Name').reset_index(drop=True)


def prep_work(df1, df2):
    return _prepare_compare_df(df1), _prepare_compare_df(df2)


def _compare_name(name1, name2):
    # Tom Man and Tina Woman would match, Tom and Martina Man would not
    if fuzz.ratio(name1, name2) >= 99 or name1 in name2 or name2 in name1:
        return []
    return ["Name"]


_STREET_REGEX = re.compile(r"(\D+)\s*(\d+)?")


def _compare_street(street1, street2):
    """Gibt die Abweichungen zurück oder None, wenn eine Straße nicht gelesen werden kann."""
    match1 = _STREET_REGEX.match(street1)
    if not match1:
        print("Chat-Datei: Konnte Straße", street1, "nicht matchen")
        return None
    match2 = _STREET_REGEX.match(street2)
    if not match2:
        print("PO-Datei: Konnte Straße", street2, "nicht matchen")
        return None
    name1, name2 = match1.group(1).strip(), match2.group(1).strip()
    if len(name1) != len(name2):
        return ["Straße"]
    if len(name1) == 1:
        return []
    if name1 != name2:
        return ["Straße"]
    # the house number has to match (9a vs 9 is ok)
    number1, number2 = re.search(r'\d+', street1), re.search(r'\d+', street2)
    if not number1 or not number2 or number1.group(0) != number2.group(0):
        return ["HNr"]
    return []


def _compare_city(city_line1, city_line2):
    parts1, parts2 = city_line1.split(), city_line2.split()
    if len(parts1) < 2:
        raise TrapoError(f"Datei 1: Oh oh, hier ist was schief gelaufen, bitte überprüfen und neustarten: "
                         f"{city_line1}")
    if len(parts2) < 2:
        raise TrapoError(f"Datei 2: Oh oh, hier ist was schief gelaufen, bitte überprüfen und neustarten: "
                         f"{city_line2}")
    reasons = []
    if parts1[0] != parts2[0]:  # post code has to be an exact match
        reasons.append("Plz")
    if parts1[1] != parts2[1]:  # Rothenburg & Rothenburg ob der Tauber have to match
        reasons.append("Stadt")
    return reasons


def compare_contact(cont1, cont2):
    """Vergleicht zwei bereinigte Kontakte; gibt (ist_gleich, Abweichungen) zurück."""
    parts1, parts2 = cont1.split(","), cont2.split(",")
    if len(parts1) != len(parts2):
        # TODO if Tierheim -> make ok
        return False, ["Länge"]
    if len(parts1) > 3 and "Tierheim" in parts1[0]:
        del parts1[1]
        del parts2[1]

    reasons = []
    for index, (part1, part2) in enumerate(zip(parts1, parts2)):
        part1, part2 = part1.strip(), part2.strip()
        if index == 0:
            reasons += _compare_name(part1, part2)
        elif index == 1:
            street_reasons = _compare_street(part1, part2)
            if street_reasons is None:
                return False, ["Regex-Fehler"]
            reasons += street_reasons
        elif index == 2:
            reasons += _compare_city(part1, part2)
        else:
            raise TrapoError(f"Oh oh, hier ist was schief gelaufen, bitte überprüfen und neustarten: "
                             f"{cont1} {cont2}")
    return not reasons, reasons


def match_pet(row, df):
    """Sucht das Tier zuerst über den (eindeutigen) Namen, sonst über den Chip; leere Serie wenn nichts passt."""
    by_name = df[df["Name"] == row["Name"]]
    if len(by_name) == 1:
        return by_name.iloc[0]
    if row["Chip"] != "":
        by_chip = df[df["Chip"] == row["Chip"]]
        if not by_chip.empty:
            return by_chip.iloc[0]
    return pd.Series()


def _show(value):
    return value if value != "" else "''"


def _diff(label, old, new):
    return f"{label}: {_show(old)} {ARROW} {_show(new)}"


def _pet_differences(row, matched):
    diffs = []
    for column in ("Chip", "DOB"):
        if row[column] != matched[column]:
            diffs.append(_diff(column, row[column], matched[column]))
    is_same, reasons = compare_contact(row["Kontakt"], matched["Kontakt"])
    if not is_same:
        diffs.append(_diff(f"Kontakt ({', '.join(reasons)})", row["Kontakt"], matched["Kontakt"]))
    return diffs


def compare(df1, df2):
    """Vergleicht die Chat-Tabelle (df1) mit der PetOffice-Tabelle (df2)."""
    df1, df2 = prep_work(df1, df2)

    differences = []
    try:
        for _, row in df1.iterrows():
            matched = match_pet(row, df2)
            if matched.empty:
                differences.append({"Name": row["Name"], "Chip": row["Chip"], "Kontakt": row["Kontakt"],
                                    CHAT_DIFF_COLUMN: "Fehlt in PetOffice-Datei"})
                continue
            diffs = _pet_differences(row, matched)
            differences.append({"Name": row["Name"], "Ort": row["Ort"], "Chip": row["Chip"], "DOB": row["DOB"],
                                "Kontakt": row["Kontakt"],
                                CHAT_DIFF_COLUMN: ", ".join(diffs) if diffs else OK_MARK})

        for _, row in df2.iterrows():
            if match_pet(row, df1).empty:
                differences.append({"Name": row["Name"], "Ort": "?", "Chip": row["Chip"], "DOB": row["DOB"],
                                    "Kontakt": row["Kontakt"], CHAT_DIFF_COLUMN: "Fehlt in Chat-Datei"})
    except KeyError as err:
        raise TrapoError(f"Spalte nicht in Tabelle gefunden: {err}") from err

    return pd.DataFrame(differences).sort_values(['Name', 'Chip'])


# ──────────────────────────────────────────────
#  Vergleich mit Traces
# ──────────────────────────────────────────────

def compare_traces(df1, df2):
    """Vergleicht die Trapo_Vergleich-Tabelle (df1) mit den extrahierten Traces-Daten (df2)."""
    df1["Chip"] = df1["Chip"].apply(chip_to_str)
    df2["Chip"] = df2["Chip"].apply(chip_to_str)
    # the chip can also be part of the difference text, e.g. "Chip: 123 -> 456"
    diff_column = next((col for col in df1.columns if "Differenz" in col), None)
    if diff_column is None:
        raise TrapoError("Kein Spaltennamen mit 'Differenz' gefunden")

    differences = []
    try:
        for _, row in df1.iterrows():
            traces_rows = df2[df2["Chip"] == row["Chip"]]
            if traces_rows.empty:
                traces_rows = df2[df2["Chip"] == row[diff_column]]
            if traces_rows.empty:
                differences.append({"Name": row["Name"], "Chip": str(row["Chip"]), "Kontakt": row["Kontakt"],
                                    "Intra": "?", "Datei": "?", "Kennzeichen": "?",
                                    TRACES_DIFF_COLUMN: "Fehlt in Traces-Dokumenten"})
                continue
            matched = traces_rows.iloc[0]
            is_same, reasons = compare_contact(row["Kontakt"], matched["Kontakt"])
            difference = OK_MARK if is_same else _diff(f"Kontakt ({', '.join(reasons)})",
                                                       row["Kontakt"], matched["Kontakt"])
            differences.append({"Name": row["Name"], "Ort": row["Ort"], "Chip": row["Chip"], "DOB": row["DOB"],
                                "Kontakt": row["Kontakt"], "Intra": matched["Intra"], "Datei": matched["Datei"],
                                "Kennzeichen": matched["Kennzeichen"], TRACES_DIFF_COLUMN: difference})

        known_chips = set(df1["Chip"]) | set(df1[diff_column])
        for _, row in df2.iterrows():
            if row["Chip"] not in known_chips:
                differences.append({"Name": "?", "Ort": "?", "Chip": str(row["Chip"]), "DOB": "?",
                                    "Kontakt": row["Kontakt"], "Intra": row["Intra"], "Datei": row["Datei"],
                                    "Kennzeichen": row["Kennzeichen"], TRACES_DIFF_COLUMN: "Fehlt in Chat-Datei"})
    except KeyError as err:
        raise TrapoError(f"Spalte nicht in Tabelle gefunden: {err}") from err

    return pd.DataFrame(differences).sort_values(['Name', 'Chip'])


def build_file_name(df):
    """Baut 'Intra_Tiernamen_Adoptant.pdf' je Traces-Datei.

    Gibt (df mit Spalte 'Datei neu', alte Dateipfade, neue Dateinamen) zurück; die beiden Listen
    enthalten nur Dateien, für die ein Name gebaut werden konnte, und sind paarweise zugeordnet.
    """
    old_files = [f for f in df["Datei"].drop_duplicates() if f not in ("", "?")]
    candidates = list(zip(df["Intra"], df["Name"], df["Kontakt"]))
    new_names = {}
    for file in old_files:
        animals = []
        contact = intra = ""
        # every animal whose Intra number is part of the file path belongs to this file
        for row_intra, name, row_contact in candidates:
            if row_intra != "?" and row_intra in file:
                animals.append(name)
                if contact == "":
                    contact, intra = row_contact, row_intra
        if contact == "" or intra == "":
            continue
        # Intranummer_Tiername_Vorname Nachname
        new_names[file] = f"{intra}_{'_'.join(animals)}_{contact.split(',')[0]}.pdf"

    df["Datei neu"] = df["Datei"].map(new_names).fillna("")
    return df, list(new_names), list(new_names.values())


def write_new_file_names(df):
    """Wie `build_file_name`, aber ohne ä, ö, ü und ß in den neuen Namen."""
    df, old, new = build_file_name(df)
    df["Datei neu"] = df["Datei neu"].apply(clean_german)
    return df, old, [clean_german(name) for name in new]


def combine_dfs(dfs):
    try:
        return pd.concat([df.reset_index(drop=True) for df in dfs], ignore_index=True)
    except Exception as err:
        raise TrapoError(f"Ups, da passt was nicht {err}") from err


# ──────────────────────────────────────────────
#  Traces-Ordner den Stopps zuordnen
# ──────────────────────────────────────────────

MIN_NAME_MATCHES = 5  # in case of double names, wait for 5 matches


def _find_document(traces_files, documents, seen_stops):
    """Gibt (Word-Datei, Stopp) des ersten Dokuments zurück, dessen Tiernamen oft genug in den Traces-Dateien stehen."""
    matches = 0
    for file, stop, names in documents:
        if stop == "" or stop in seen_stops:
            continue
        for name in names:
            matches += sum(name in traces_file for traces_file in traces_files)
            if matches >= MIN_NAME_MATCHES:
                return file, stop
    return None


def find_stopp_for_plate(files, dfs, folders):
    """Findet je Kennzeichen-Ordner das Word-Dokument (Stopp), dessen Tiernamen in den Traces-Dateien vorkommen.

    Gibt eine Liste von (Kennzeichen, Stopp, Word-Datei) zurück.
    """
    documents = [(file, extract_stop(file), df["Name"].apply(clean_name)) for file, df in zip(files, dfs)]
    results = []
    seen_stops = set()
    for plate in folders:
        traces_files = io_helpers.get_all_files_from_folder(os.path.join(".", plate) + "/*.pdf")
        found = _find_document(traces_files, documents, seen_stops)
        if found:
            file, stop = found
            results.append((plate, stop, file))
            seen_stops.add(stop)
    return results


def extract_stop(file):
    """Erkennt den Stopp (NORD, MITTE, SÜDWEST, SÜD) am Dateinamen."""
    file = io_helpers.simple_normalize(file)
    if "nord" in file:
        return "NORD"
    if "mitte" in file:
        return "MITTE"
    if "sudwest" in file or "suedwest" in file:
        return "SÜDWEST"
    if "sud" in file or "sued" in file:
        return "SÜD"
    return ""


# ──────────────────────────────────────────────
#  Entfernungstabellen (trapo-km)
# ──────────────────────────────────────────────

def _replace_empty(cell):
    return '?' if pd.isna(cell) or str(cell).strip() == '' else cell


def shrink_tables(dfs):
    """Behält nur Foto, Name, Kontakt und Treffpunkt."""
    results = []
    for df in dfs:
        new_df = df[[col for col in df.columns if col in DISTANCE_COLUMNS]].copy()
        new_df["Treffpunkt"] = new_df["Treffpunkt"].apply(_replace_empty)
        results.append(new_df)
    return results


def clean_plate_dfs(dfs):
    results = []
    for df in dfs:
        df['Name'] = df['Name'].apply(clean_name_with_box_change)
        df['Kontakt'] = df['Kontakt'].apply(clean_contact)
        if 'Ort' in df.columns:
            df["Ort"] = df["Ort"].apply(clean_location)
        results.append(df.sort_values('Name').reset_index(drop=True))
    return results


RENTAL_REGEX = re.compile(r"Miet[\s-]?(wagen|auto)", re.I)
GERMAN_PLATE_REGEX = re.compile(r'\b([A-ZÄÖÜ]{1,3})\s*[- :]?\s*([A-Z]{1,2})\s*[- ]?\s*(\d{1,4})(\s*E)?\b')
AUSTRIAN_PLATE_REGEX = re.compile(r'\b([A-ZÄÖÜ]{1,3})\s*(\d{1,4})([A-Z]{1,2})\b')
SWISS_PLATE_REGEX = re.compile(r'\b([A-Z]{2})\s*(\d{1,6})\b')


def _german_plate(city, letters, numbers, electric):
    return f"{city}-{letters} {numbers}{' E' if electric else ''}"


def _austrian_plate(district, numbers, letters):
    return f"{district}-{numbers}{letters} (AT)"


def _swiss_plate(canton, numbers):
    return f"{canton}-{numbers} (CH)"


def extract_all_plates(text):
    """Findet deutsche, österreichische und Schweizer Kennzeichen in `text` und gibt sie als Text zurück."""
    text = text.upper()
    plates = []
    matched_spans = []

    def is_overlapping(span):
        return any(max(start, span[0]) < min(end, span[1]) for start, end in matched_spans)

    # earlier patterns win when two patterns match the same characters
    for regex, build_plate in ((GERMAN_PLATE_REGEX, _german_plate),
                               (AUSTRIAN_PLATE_REGEX, _austrian_plate),
                               (SWISS_PLATE_REGEX, _swiss_plate)):
        for match in regex.finditer(text):
            if not is_overlapping(match.span()):
                plates.append(build_plate(*match.groups()))
                matched_spans.append(match.span())

    result = ", ".join(plates) if plates else "----"
    if RENTAL_REGEX.search(text):
        result += " (Mietwagen)"
    return result


def find_plate(name, names_only, row_lookup, df):
    """Sucht den ähnlichsten Namen (>= 80 %) in der Kennzeichen-Tabelle und gibt dessen Kennzeichen zurück."""
    found = process.extractOne(name.lower(), names_only, scorer=fuzz.ratio, score_cutoff=80)
    if found is None:
        return ""
    return extract_all_plates(df.loc[row_lookup[found[0]], "Kennzeichen"])


def add_plates(dfs, df1):
    """Fügt jeder Tabelle die Spalte 'Kennzeichen' (aus der Kennzeichen-Tabelle df1) hinzu."""
    df1.columns = (
        df1.columns
        .str.replace(r'^Name.*', 'Name', regex=True)
        .str.replace(r'^Kennzeichen.*', 'Kennzeichen', regex=True, flags=re.DOTALL)
    )
    # every word of the name cell is a candidate, e.g. "Rex & Bello" -> rex, bello
    row_lookup = {}
    for idx, row in df1.iterrows():
        for word in re.findall(r'\b[A-ZÄÖÜa-zäöüß]{2,}\b', str(row['Name'])):
            row_lookup[word.lower()] = idx
    names_only = list(row_lookup)

    for df in dfs:
        df.insert(loc=2, column='Kennzeichen',
                  value=df["Name"].apply(find_plate, args=(names_only, row_lookup, df1)))
    return dfs


def add_distance(dfs, stopps):
    """Berechnet die Entfernung zum Treffpunkt und sortiert je Treffpunkt (in Reihenfolge des Auftretens)."""
    print("Anfragen an OpenStreemMap sind begrenzt auf eine pro Sekunde, dieser Schritt kann dauern...")
    results = []
    for i, df in enumerate(dfs):
        print(f"Berechne Entfernungen für Dokument {i + 1} von {len(dfs)}...")
        distances = [math_helpers.calculate_distance(row, stopps) for _, row in df.iterrows()]
        df.insert(loc=min(5, len(df.columns)), column='Entfernung', value=distances)

        meeting_point_order = {tp: pos for pos, tp in enumerate(df['Treffpunkt'].unique())}
        df = (df.assign(_TreffOrder=df['Treffpunkt'].map(meeting_point_order),
                        _BoxChange=df['Name'].astype(str).str.contains(r'\(Box Change\)', case=False, na=False))
              # 1. meeting point as it appeared, 2. "(Box Change)" rows last, 3. farthest first
              .sort_values(by=['_TreffOrder', '_BoxChange', 'Entfernung'], ascending=[True, True, False])
              .drop(columns=['_TreffOrder', '_BoxChange'])
              .reset_index(drop=True))
        results.append(insert_headers(df))
    return results


def _is_same_meeting_point(current, other):
    return other in current or current in other


def insert_headers(df):
    """Fügt vor jeder Treffpunkt-Gruppe eine Kopfzeile ein und nummeriert die Zeilen je Gruppe (Spalte 'Nr.')."""
    rows = []
    current_meeting = "Treffpunkt"
    counter = 0
    for idx, row in df.iterrows():
        meeting_point = row['Treffpunkt']
        starts_group = (meeting_point != "Treffpunkt" if idx == 0
                        else not _is_same_meeting_point(current_meeting, meeting_point))
        if starts_group:
            current_meeting = meeting_point
            counter = 0
            rows.append({'Nr.': 'Nr.', **{col: col for col in df.columns}})
        counter += 1
        rows.append({'Nr.': counter, **row.to_dict()})
    return pd.DataFrame(rows)


# ──────────────────────────────────────────────
#  Rumänische Liste (trapo-ro)
# ──────────────────────────────────────────────

def translate_headers(dfs):
    """Überträgt die Tabellen in das rumänische Format (Spalten RO_HEADERS)."""
    return [_translate_table(df) for df in dfs]


def _translate_table(df):
    df = df.assign(Name=df['Name'].apply(clean_name))
    rows = []
    for number, (_, row) in enumerate(df.iterrows(), start=1):
        new_row = dict.fromkeys(RO_HEADERS)  # columns without a counterpart stay empty
        for old_col, new_col in TRANSLATE_MAP.items():
            if old_col in df.columns:
                value = row[old_col]
                new_row[new_col] = number if old_col == "No" and value == "" else value
        rows.append(new_row)
    return pd.DataFrame(rows, columns=RO_HEADERS)
