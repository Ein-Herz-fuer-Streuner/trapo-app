"""Adressen, Geocoding (Photon) und Fahrstrecken (OSRM) für die Entfernungsberechnung."""
import re
import time

import requests

from trapo_app import __version__

PHOTON_URL = "https://photon.komoot.io/api/"
OSRM_URL = "http://router.project-osrm.org/route/v1/driving/{lon1},{lat1};{lon2},{lat2}?overview=false"
REQUEST_TIMEOUT_SECONDS = 5
MIN_SECONDS_BETWEEN_REQUESTS = 1  # the public OSM services allow about one request per second
UNKNOWN_DISTANCE_KM = 500  # used when an address cannot be resolved, sorts those rows first

_session = requests.Session()
_session.headers["User-Agent"] = f"TrapoApp/{__version__}"
_last_request_time = 0.0
_coordinate_cache = {}
_distance_cache = {}


def _throttled_get(url, **kwargs):
    """GET that keeps at least MIN_SECONDS_BETWEEN_REQUESTS between two requests."""
    global _last_request_time
    wait = MIN_SECONDS_BETWEEN_REQUESTS - (time.monotonic() - _last_request_time)
    if wait > 0:
        time.sleep(wait)
    try:
        return _session.get(url, timeout=REQUEST_TIMEOUT_SECONDS, **kwargs)
    finally:
        _last_request_time = time.monotonic()


def clean_address(raw_address):
    """Gibt 'Straße Nr, PLZ Ort' zurück oder None, wenn Straße oder PLZ/Ort fehlen."""
    parts = [line.strip() for line in re.split(r'[,\n]', raw_address) if line.strip()]

    street = ""
    zip_city = ""
    for part in parts:
        if re.match(r'\d{4,}', part):  # zip code followed by the city
            zip_city = part
        elif re.search(r'\d+', part):  # street with house number
            street = part

    if street and zip_city:
        return f"{street}, {zip_city}"
    return None


def get_stopp_address(tp, df):
    """Sucht die Adresse des Treffpunkts `tp` in der Stopp-Liste."""
    for _, row in df.iterrows():
        name = row['Treffpunkt']
        if name in tp or tp in name:
            return ", ".join(row['Adresse'].split("\n")).strip()
    return ""


def get_coordinates(address):
    """Wandelt eine Adresse per Photon in (Breitengrad, Längengrad) um, None wenn nichts gefunden wurde."""
    if address in _coordinate_cache:  # the API blocks repeated requests
        return _coordinate_cache[address]

    try:
        res = _throttled_get(PHOTON_URL, params={"q": address, "limit": 1})
        res.raise_for_status()
        features = res.json().get("features")
    except Exception as e:
        print(f"Geocoding failed for '{address}': {e}")
        return None  # not cached, the request may work next time

    coordinates = None
    if features:
        lon, lat = features[0]["geometry"]["coordinates"]
        coordinates = (lat, lon)
    _coordinate_cache[address] = coordinates
    return coordinates


def get_driving_distance(coord1, coord2):
    """Fahrstrecke in ganzen Kilometern zwischen zwei (lat, lon)-Koordinaten per OSRM, None bei Fehlern."""
    cache_key = (coord1, coord2)
    if cache_key in _distance_cache:
        return _distance_cache[cache_key]

    url = OSRM_URL.format(lat1=coord1[0], lon1=coord1[1], lat2=coord2[0], lon2=coord2[1])
    try:
        routes = _throttled_get(url).json().get("routes")
    except Exception as e:
        print(f"Routenberechnung fehlgeschlagen: {e}")
        return None
    if not routes:
        return None
    distance_km = int(routes[0]["distance"] / 1000)
    _distance_cache[cache_key] = distance_km
    return distance_km


def calculate_distance(row, stopps_df):
    """Fahrstrecke in km vom Adoptanten (`Kontakt`) zum Treffpunkt, UNKNOWN_DISTANCE_KM bei Fehlern."""
    address_from = clean_address(row['Kontakt'])
    if not address_from:
        print("Adresse ist nicht formatierbar", row['Kontakt'])
        return UNKNOWN_DISTANCE_KM
    stopp = row['Treffpunkt']
    address_to = clean_address(get_stopp_address(stopp, stopps_df))
    if not address_to:
        print("Adresse ist nicht formatierbar", stopp)
        return UNKNOWN_DISTANCE_KM

    coord_from = get_coordinates(address_from)
    if coord_from is None:
        print("Konnte keine Koordinaten für den Adoptanten finden: ", address_from)
        return UNKNOWN_DISTANCE_KM
    coord_to = get_coordinates(address_to)
    if coord_to is None:
        print("Konnte keine Koordinaten für den Treffpunkt finden", address_to)
        return UNKNOWN_DISTANCE_KM
    if coord_from == coord_to:
        return 0

    distance = get_driving_distance(coord_from, coord_to)
    if distance is None:
        print(f"Konnte keine Distanz von {address_from} zu {address_to} finden")
        return UNKNOWN_DISTANCE_KM
    return distance
