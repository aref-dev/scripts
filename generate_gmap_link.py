def generate_google_maps_link(latitude, longitude):
    return f"https://www.google.com/maps?q={latitude},{longitude}"

# Example usage
latitude = 37.4221
longitude = -122.0841
print(generate_google_maps_link(latitude, longitude))
