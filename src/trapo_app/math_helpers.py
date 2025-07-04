import random
import geopy.geocoders as geopy
import certifi
import ssl
import requests

def get_stopp_address(tp, df):
    for _, row in df.iterrows():
        t_ = row['Treffpunkt']
        if t_ in tp or tp in t_:
            return row['Adresse']
    return ""

def get_coordinates(address):
    """Convert an address into latitude and longitude using Nominatim."""
    ctx = ssl.create_default_context(cafile=certifi.where())
    geopy.options.default_ssl_context = ctx
    geolocator = geopy.Nominatim(user_agent="trapo-app", scheme='https')
    location = geolocator.geocode(address)
    if location:
        return location.latitude, location.longitude
    else:
        return None

def get_driving_distance(coord1, coord2):
    """Use OSRM to calculate the driving distance between two coordinates."""
    osrm_url = f"http://router.project-osrm.org/route/v1/driving/{coord1[1]},{coord1[0]};{coord2[1]},{coord2[0]}?overview=false"
    response = requests.get(osrm_url)
    data = response.json()

    if "routes" in data and data["routes"]:
        distance_meters = data["routes"][0]["distance"]
        distance_km = distance_meters / 1000  # Convert meters to kilometers
        return distance_km
    else:
        return None

def calculate_distance(row, stopps_df):
    '''
    address_from = row['Kontakt']
    stopp = row['Treffpunkt']
    address_to = get_stopp_address(stopp, stopps_df)
    if address_to is "":
        return 500
    coord_from = get_coordinates(address_from)
    coord_to = get_coordinates(address_to)
    distance = get_driving_distance(coord_from, coord_to)
    if not distance:
        return 500
    return distance
    '''
    return random.randint(0,400)