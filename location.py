import win32com.client
from geopy.geocoders import Nominatim

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def get_user_location():
    try:
        # Initialize Nominatim geocoder with a custom user agent
        geolocator = Nominatim(user_agent="my_user_agent")

        # Fetching location details using geocoder
        location = geolocator.geocode('me')

        if location:
            address_components = location.raw.get('address', {})
            city = address_components.get('city', '')
            region = address_components.get('state', '')
            country = address_components.get('country', '')
            postal_code = address_components.get('postcode', '')

            # Print location details
            print("Your current location details:")
            print(f"City: {city}")
            print(f"Region: {region}")
            print(f"Country: {country}")
            print(f"Postal Code: {postal_code}")
        else:
            print("Location data not available.")

    except Exception as e:
        print(f"Error: {e}")

