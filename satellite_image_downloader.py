import os
import requests
from dotenv import load_dotenv

# Set the path to the .env file
dotenv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'StructFlow-Automator-Private', '.env')

def get_coordinates(address, api_key):
    base_url = "https://maps.googleapis.com/maps/api/geocode/json"
    params = {
        "address": address,
        "key": api_key
    }
    response = requests.get(base_url, params=params)
    data = response.json()
    
    if data['status'] == 'OK':
        location = data['results'][0]['geometry']['location']
        return location['lat'], location['lng']
    else:
        raise ValueError(f"Could not find coordinates for the given address. Status: {data['status']}")

def get_satellite_image(address, api_key, zoom=18, size="608x325"):
    # Get coordinates from address
    lat, lng = get_coordinates(address, api_key)
    
    # Calculate bounding box (this is an approximation and may need adjustment based on zoom level)
    offset = 0.001 * (21 - zoom) # Adjust this value to change the size of the border
    box = f"{lat-offset},{lng-offset}|{lat-offset},{lng+offset}|{lat+offset},{lng+offset}|{lat+offset},{lng-offset}|{lat-offset},{lng-offset}"
    
    # Construct the URL for Google Static Maps API
    base_url = "https://maps.googleapis.com/maps/api/staticmap?"
    params = {
        "center": f"{lat},{lng}",
        "zoom": zoom,
        "size": size,
        "maptype": "satellite",
        "key": api_key,
        "markers": f"color:red|label:A|{lat},{lng}",  # Add a labeled red marker
        "path": f"color:0xFFFF00FF|weight:5|{box}"  # Add a yellow border around the address
    }
    
    url = base_url + "&".join(f"{k}={v}" for k, v in params.items())
    
    # Send request to Google Static Maps API
    response = requests.get(url)
    if response.status_code != 200:
        raise Exception(f"Failed to download image. Status code: {response.status_code}")
    
    # Create temp directory if it doesn't exist
    temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp")
    os.makedirs(temp_dir, exist_ok=True)
    
    # Save the image
    filename = os.path.join(temp_dir, f"satellite_image.jpg")
    with open(filename, "wb") as f:
        f.write(response.content)
    
    print(f"Satellite image saved to: {filename}")

def get_valid_size(size_input):
    default_size = "608x325"
    if not size_input:
        return default_size
    try:
        width, height = map(int, size_input.lower().split('x'))
        if width <= 0 or height <= 0:
            raise ValueError
        return f"{width}x{height}"
    except:
        print(f"Invalid size input. Using default size: {default_size}")
        return default_size

def get_valid_zoom(zoom_input):
    default_zoom = 18
    if not zoom_input:
        return default_zoom
    try:
        zoom = int(zoom_input)
        if 0 <= zoom <= 21:
            return zoom
        else:
            raise ValueError
    except:
        print(f"Invalid zoom level. Using default zoom: {default_zoom}")
        return default_zoom

def main():
    # Load environment variables from the specified .env file
    load_dotenv(dotenv_path)
    
    api_key = os.getenv("GOOGLE_MAPS_API_KEY")
    if not api_key:
        raise ValueError("GOOGLE_MAPS_API_KEY not found in environment variables")
    
    address = input("Enter the address: ")
    size_input = input("Enter the image size (width x height, e.g., 608x325) or press Enter for default: ")
    zoom_input = input("Enter the zoom level (0-21, where 21 is the closest) or press Enter for default: ")
    
    size = get_valid_size(size_input)
    zoom = get_valid_zoom(zoom_input)
    
    get_satellite_image(address, api_key, zoom=zoom, size=size)

if __name__ == "__main__":
    main()