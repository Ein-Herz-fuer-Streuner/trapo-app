from trapo_app import math_helpers as mh


def test_clean_address_joins_street_and_zip_city():
    assert mh.clean_address("Hauptstr. 5\n12345 Berlin") == "Hauptstr. 5, 12345 Berlin"


def test_clean_address_without_street_is_none():
    assert mh.clean_address("nix") is None
