import re
import requests

cache = {}
dist_cache = {}

def clean_address(raw_address):
    # Split into lines, remove empty ones and trim whitespace
    parts = [line.strip() for line in re.split(r'[,\n]', raw_address) if line.strip()]

    street = ""
    zip_city = ""

    for part in parts:
        # Match zip and city: 5-digit zip followed by city name
        if re.match(r'\d{4,}', part):
            zip_city = part.strip()
        # Match street + number (number usually at the end or beginning)
        elif re.search(r'\d+', part):
            street = part.strip()

    if street and zip_city:
        return f"{street}, {zip_city}"
    else:
        return None

def get_stopp_address(tp, df):
    for _, row in df.iterrows():
        t_ = row['Treffpunkt']
        if t_ in tp or tp in t_:
            tmp = row['Adresse']
            tmp = ", ".join(tmp.split("\n")).strip()
            return tmp
    return ""

def get_coordinates(address):
    """Convert an address into latitude and longitude using Nominatim."""
    # cache or else API blocks duplicate requests
    if address in cache:
        (lat, lang) = cache[address]
        return lat, lang

    url = "https://photon.komoot.io/api/"
    params = {"q": address, "limit": 1}
    try:
        res = requests.get(url, params=params, timeout=5)
        res.raise_for_status()
        data = res.json()
        if data.get("features"):
            lon, lat = data["features"][0]["geometry"]["coordinates"]
            return (lat, lon)
    except Exception as e:
        print(f"Geocoding failed for '{address}': {e}")
    return None

def get_driving_distance(coord1, coord2):
    cache_key = f"{coord1}_{coord2}"
    if cache_key in dist_cache:
        return dist_cache[cache_key]
    """Use OSRM to calculate the driving distance between two coordinates."""
    osrm_url = f"http://router.project-osrm.org/route/v1/driving/{coord1[1]},{coord1[0]};{coord2[1]},{coord2[0]}?overview=false"
    response = requests.get(osrm_url)
    data = response.json()

    if "routes" in data and data["routes"]:
        distance_meters = data["routes"][0]["distance"]
        distance_km = int(distance_meters / 1000)  # Convert meters to kilometers
        dist_cache[cache_key] = distance_km
        return distance_km
    else:
        return None

def calculate_distance(row, stopps_df):
    address_from = clean_address(row['Kontakt'])
    if address_from is None:
        print("Adresse ist nicht formatierbar", row['Kontakt'])
        return 500
    stopp = row['Treffpunkt']
    address_to = clean_address(get_stopp_address(stopp, stopps_df))
    if address_to is None:
        print("Adresse ist nicht formatierbar", stopp)
        return 500
    if address_to == "":
        return 500
    coord_from = get_coordinates(address_from)
    if coord_from is None:
        print("Konnte keine Koordinaten für den Adopter finden: ", address_from )
        return 500
    coord_to = get_coordinates(address_to)
    if coord_to is None:
        print("Konnte keine Koordinaten für den Treffpunkt finden", address_to)
        return 500
    if coord_to[0] == coord_from[0] and coord_to[1] == coord_from[1]:
        return 0
    distance = get_driving_distance(coord_from, coord_to)
    if not distance:
        print(f"Konnte keine Distanz von {address_from} zu {address_from} finden")
        return 500
    return distance
