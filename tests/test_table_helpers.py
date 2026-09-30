import pandas as pd
import pytest

from trapo_app import table_helpers as th


class TestCleanName:
    def test_capitalizes_and_strips_box_change(self):
        assert th.clean_name("BOX CHANGE!  rex maX") == "Rex Max"

    def test_plain_name(self):
        assert th.clean_name("bello") == "Bello"

    def test_box_change_marker_is_appended(self):
        assert th.clean_name_with_box_change("Box Change Bello") == "Bello (box Change)"

    def test_without_box_change_only_strips(self):
        assert th.clean_name_with_box_change(" bello ") == "Bello"


class TestCleanDob:
    @pytest.mark.parametrize("raw, expected", [
        ("01.02.23", "01.02.2023"),
        ("01.02.2023", "01.02.2023"),
        ("1 2 2023", "01.02.2023"),
        ("01/02/2023", "01.02.2023"),
        ("", ""),
        (" ", ""),
        ("nan", ""),
    ])
    def test_formats(self, raw, expected):
        assert th.clean_dob(raw) == expected

    def test_invalid_date_aborts(self):
        with pytest.raises(SystemExit):
            th.clean_dob("someday")


class TestCleanContact:
    def test_removes_phone_and_mail_and_normalizes_street(self):
        raw = "Max Mustermann und Erika\nHauptstr. 5\n12345 Berlin\nTel: 0171 1234567\nmax@web.de"
        assert th.clean_contact(raw) == "Max Mustermann & Erika, Hauptstraße 5, 12345 Berlin"

    def test_keeps_words_containing_und(self):
        assert th.clean_contact("Anna Grund\nGrundstraße 4\n12345 Hund") == "Anna Grund, Grundstraße 4, 12345 Hund"

    def test_replaces_standalone_und_and_u_dot(self):
        assert th.clean_contact("Max u. Erika\nHauptstraße 5\n12345 Berlin").startswith("Max & Erika")


class TestCompareContact:
    def test_same_contact_with_house_number_suffix(self):
        assert th.compare_contact("Max Mustermann, Hauptstraße 5, 12345 Berlin",
                                  "Max Mustermann, Hauptstraße 5a, 12345 Berlin") == (True, [])

    def test_reports_all_differences(self):
        same, reasons = th.compare_contact("Max Mustermann, Hauptstraße 5, 12345 Berlin",
                                           "Max Muster, Hauptstraße 6, 12346 Bonn")
        assert not same
        assert reasons == ["HNr", "Plz", "Stadt"]

    def test_different_length(self):
        assert th.compare_contact("A, B", "A, B, C") == (False, ["Länge"])

    def test_different_name(self):
        same, reasons = th.compare_contact("Tom Man, Hauptstraße 5, 12345 Berlin",
                                           "Lisa Frau, Hauptstraße 5, 12345 Berlin")
        assert not same and reasons == ["Name"]


class TestExtractAllPlates:
    def test_german_plate_with_e(self):
        assert th.extract_all_plates("B-AB 1234 E") == "B-AB 1234 E"

    def test_no_plate(self):
        assert th.extract_all_plates("nichts") == "----"

    def test_austrian_and_swiss(self):
        assert th.extract_all_plates("W 1234AB und ZH 12345") == "W-1234AB (AT), ZH-12345 (CH)"

    @pytest.mark.parametrize("text", ["Mietwagen", "Miet-Auto", "Mietwagen B-AB 123"])
    def test_rental_car_is_flagged(self, text):
        assert th.extract_all_plates(text).endswith("(Mietwagen)")


class TestExtractStop:
    @pytest.mark.parametrize("name, stop", [
        ("31.01.26-NORD-V2.docx", "NORD"),
        ("31.01.26-MITTE-V2.docx", "MITTE"),
        ("31.01.26-SÜDWEST-V2.docx", "SÜDWEST"),
        ("x-SUED-y", "SÜD"),
        ("foo", ""),
    ])
    def test_stop(self, name, stop):
        assert th.extract_stop(name) == stop


class TestInsertHeaders:
    def test_header_row_per_meeting_point_and_numbering(self):
        df = pd.DataFrame({"Name": ["A", "B", "C"], "Treffpunkt": ["X", "X", "Y"]})
        out = th.insert_headers(df)
        assert out["Nr."].tolist() == ["Nr.", 1, 2, "Nr.", 1]
        assert out["Name"].tolist() == ["Name", "A", "B", "Name", "C"]


class TestBuildFileName:
    def test_names_files_by_intra_animals_and_contact(self):
        df = pd.DataFrame({
            "Name": ["Rex", "Bello"],
            "Kontakt": ["Max Mustermann, Weg 1, 12345 Berlin"] * 2,
            "Intra": ["INT1", "INT1"],
            "Datei": ["/x/INT1_scan.pdf", "/x/INT1_scan.pdf"],
        })
        out, old, new = th.build_file_name(df)
        assert old == ["/x/INT1_scan.pdf"]
        assert new == ["INT1_Rex_Bello_Max Mustermann.pdf"]
        assert out["Datei neu"].tolist() == new * 2


class TestTranslateHeaders:
    def test_maps_columns_and_numbers_rows(self):
        df = pd.DataFrame({"No": ["", ""], "Name": ["rex", "bello"], "Age": ["3", "4"]})
        out = th.translate_headers([df])[0]
        assert out["Nume"].tolist() == ["Rex", "Bello"]
        assert out["Nr Crt"].tolist() == [1, 2]
        assert list(out.columns) == th.ro_headers
