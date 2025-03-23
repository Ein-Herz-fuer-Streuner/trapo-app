import geopy.geocoders as geopy
import certifi
import ssl
import requests

'''
ctx = ssl.create_default_context(cafile=certifi.where())
geopy.options.default_ssl_context = ctx
geolocator = geopy.Nominatim(user_agent="trapo-app",scheme='https')
location_to = geolocator.geocode("Moosacher Straße 3 85614 Kirchseeon", language='de')
print((location_to.latitude, location_to.longitude))
location_from = geolocator.geocode("Effnerstraße 119 81925 München", language='de')
print((location_from.latitude, location_from.longitude))
'''


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


# Example: Addresses in Germany
address1 = "Cranachstraße 53, 99423 Weimar"
address2 = "Nordhäuser Str. 6, 37412 Herzberg"

coord1 = get_coordinates(address1)
coord2 = get_coordinates(address2)

if coord1 and coord2:
    distance = get_driving_distance(coord1, coord2)
    print(f"Die Strecke von\n{address1}\nnach\n{address2}\nbeträgt {distance:.2f}km.")
else:
    print("Could not determine coordinates for one or both addresses.")
