import pandas as pd
import pytest
import requests

from trapo_app import math_helpers as mh


def test_clean_address_joins_street_and_zip_city():
    assert mh.clean_address("Hauptstr. 5\n12345 Berlin") == "Hauptstr. 5, 12345 Berlin"


def test_clean_address_without_street_is_none():
    assert mh.clean_address("nix") is None


class FakeResponse:
    def __init__(self, payload):
        self.payload = payload

    def raise_for_status(self):
        pass

    def json(self):
        return self.payload


@pytest.fixture(autouse=True)
def clean_state(monkeypatch):
    monkeypatch.setattr(mh, "_coordinate_cache", {})
    monkeypatch.setattr(mh, "_distance_cache", {})


def test_coordinates_are_cached(monkeypatch):
    calls = []

    def fake_get(url, **kwargs):
        calls.append(url)
        return FakeResponse({"features": [{"geometry": {"coordinates": [13.4, 52.5]}}]})

    monkeypatch.setattr(mh, "_throttled_get", fake_get)
    assert mh.get_coordinates("Weg 1, 12345 Berlin") == (52.5, 13.4)
    assert mh.get_coordinates("Weg 1, 12345 Berlin") == (52.5, 13.4)
    assert len(calls) == 1


def test_unknown_address_is_cached_too(monkeypatch):
    calls = []
    monkeypatch.setattr(mh, "_throttled_get", lambda url, **kw: calls.append(url) or FakeResponse({"features": []}))
    assert mh.get_coordinates("Nirgendwo 1, 00000 Nichts") is None
    assert mh.get_coordinates("Nirgendwo 1, 00000 Nichts") is None
    assert len(calls) == 1


def test_network_error_is_not_cached(monkeypatch):
    def failing_get(url, **kwargs):
        raise requests.ConnectionError("offline")

    monkeypatch.setattr(mh, "_throttled_get", failing_get)
    assert mh.get_coordinates("Weg 1, 12345 Berlin") is None
    assert mh._coordinate_cache == {}


def test_driving_distance_in_whole_km(monkeypatch):
    monkeypatch.setattr(mh, "_throttled_get", lambda url, **kw: FakeResponse({"routes": [{"distance": 12345}]}))
    assert mh.get_driving_distance((1, 2), (3, 4)) == 12


def test_driving_distance_error_returns_none(monkeypatch):
    def failing_get(url, **kwargs):
        raise requests.Timeout()

    monkeypatch.setattr(mh, "_throttled_get", failing_get)
    assert mh.get_driving_distance((1, 2), (3, 4)) is None


def _rows():
    row = {"Kontakt": "Max, Weg 1, 12345 Berlin", "Treffpunkt": "Nord"}
    stopps = pd.DataFrame({"Treffpunkt": ["Nord"], "Adresse": ["Ring 2\n54321 Bonn"]})
    return pd.Series(row), stopps


def test_distance_under_one_km_is_zero_not_a_failure(monkeypatch):
    monkeypatch.setattr(mh, "get_coordinates", lambda address: (1.0, 2.0) if "Weg" in address else (1.0, 2.001))
    monkeypatch.setattr(mh, "get_driving_distance", lambda a, b: 0)
    assert mh.calculate_distance(*_rows()) == 0


def test_unresolvable_address_uses_fallback_distance(monkeypatch):
    monkeypatch.setattr(mh, "get_coordinates", lambda address: None)
    assert mh.calculate_distance(*_rows()) == mh.UNKNOWN_DISTANCE_KM
