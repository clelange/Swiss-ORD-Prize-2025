import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from shapely.geometry import shape
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from adjustText import adjust_text
import json
import urllib.request

# List of institutions
institutions = ['Université de Genève', 'University of Geneva', 'University of Basel', 'ETH Zürich', 'University Zurich', 'University of Lausanne', 'Université de Genève', 'University of Geneva', 'University of Lugano (USI)', 'University of Lausanne', 'Università della Svizzera italiana', 'Eawag', 'Eastern Switzerland University of Applied Sciences', 'University of Zurich', 'Swiss Federal Institute for Forest, Snow and Landscape Research WSL', 'ETH Zürich', 'EPFL', 'Ecole Polytechnique Fédérale de Lausanne', 'University of Applied Sciences and Arts Western Switzerland (HES-SO)', 'University of Lucerne', 'University of Geneva', 'Università della Svizzera italiana', 'ETH Zürich', 'University of Lausanne', 'Institute of Earth Surface Dynamics (IDYST)', 'University of Zurich', 'University of Applied Sciences and Arts of Southern Switzerland', 'University of St. Gallen', 'University of Lausanne', 'Bern University of Applied Sciences', 'EPFL', 'SUPSI', 'École polytechnique fédérale de Lausanne (EPFL)', 'University of Geneva', 'Dodis (Diplomatic Documents of Switzerland)', 'Université de Fribourg', 'Empa', 'ETH Zurich', 'EPFL', 'EAWAG', 'Ecole Polytechnique Fédérale de Lausanne', 'Università della Svizzera italiana', 'SIB Swiss Institute of Bioinformatics', 'Hochschule der Künste Bern', 'Lucerne School of Design, Film and Art']

institute_translations = {
    'Université de Genève': 'University of Geneva',
    'University Zurich': 'University of Zurich',
    'University of Lugano (USI)': 'Università della Svizzera italiana',
    'Ecole Polytechnique Fédérale de Lausanne': 'EPFL',
    'École polytechnique fédérale de Lausanne (EPFL)': 'EPFL',
    'Institute of Earth Surface Dynamics (IDYST)': 'University of Lausanne',
    'Eawag': 'EAWAG',
    'ETH Zürich': 'ETHZ',
    'ETH Zurich': 'ETHZ',
    'University of Applied Sciences and Arts of Southern Switzerland': 'University of Applied Sciences and Arts of Southern Switzerland (SUPSI)',
    'SUPSI': 'University of Applied Sciences and Arts of Southern Switzerland (SUPSI)',
    'Swiss Federal Institute for Forest, Snow and Landscape Research WSL': 'WSL',
    'Dodis (Diplomatic Documents of Switzerland)': 'Swiss Academy of Humanities and Social Sciences',
    'Empa': 'EMPA',
    'Université de Fribourg': 'University of Fribourg',
}
# Normalize institution names
institutions = [institute_translations.get(inst, inst) for inst in institutions]

# alternate names to be used in search
alt_names = {
    'Bern University of Applied Sciences': 'Murtenstrasse 10, 3008 Bern',
    'Eastern Switzerland University of Applied Sciences': 'Ostschweizer Fachhochschule',
    'University of Applied Sciences and Arts Western Switzerland (HES-SO)': 'Sion',
    'Lucerne School of Design, Film and Art': 'Hochschule Luzern',
    'SIB Swiss Institute of Bioinformatics': 'Basel',
    'WSL': 'Davos',
    'University of Zurich': 'Rämistrasse 71, 8006 Zürich',
    'Swiss Academy of Humanities and Social Sciences': 'Laupenstrasse 7, 3001 Bern',
    'University of Applied Sciences and Arts of Southern Switzerland (SUPSI)': 'Le Gerre, Via Pobiette 11, 6928 Manno',
    'University of Lucerne': 'Frohburgstrasse 3, 6002 Luzern',
}

# Deduplicate and clean
unique_institutions = sorted(list(set(institutions)))
print(f"Found {len(unique_institutions)} unique institutions.")

# Geocoding with caching
cache_file = 'geocoding_cache.json'

# Load existing cache if available
try:
    with open(cache_file, 'r') as f:
        locations = json.load(f)
        # Convert lists back to tuples
        locations = {k: tuple(v) for k, v in locations.items()}
    print(f"Loaded {len(locations)} institutions from cache.")
except FileNotFoundError:
    locations = {}
    print("No cache file found, will query all institutions.")

# Initialize geocoder
geolocator = Nominatim(user_agent="swiss_ord_prize_map_v1")
geocode = RateLimiter(geolocator.geocode, min_delay_seconds=2)

# Geocode missing institutions
missing_institutions = [inst for inst in unique_institutions if inst not in locations]
if missing_institutions:
    print(f"Geocoding {len(missing_institutions)} new institutions...")
    for inst in missing_institutions:
        try:
            # Append "Switzerland" to ensure we get Swiss locations
            if inst in alt_names:
                query = f"{alt_names[inst]}, Switzerland"
            else:
                query = f"{inst}, Switzerland"
            location = geocode(query)
            if location:
                locations[inst] = (location.latitude, location.longitude)
                print(f"Found: {inst} -> {location.address}")
            else:
                print(f"Not found: {inst}")
        except Exception as e:
            print(f"Error geocoding {inst}: {e}")

    # Save updated cache
    cache_data = {k: list(v) for k, v in locations.items()}
    with open(cache_file, 'w') as f:
        json.dump(cache_data, f, indent=2)
    print(f"Cache updated with {len(locations)} institutions.")
else:
    print("All institutions found in cache, no geocoding needed.")

print(f"Total geocoded institutions: {len(locations)}")

# Download Switzerland GeoJSON (High Resolution)
# Source: https://labs.karavia.ch/swiss-boundaries-geojson/
geojson_url = "https://raw.githubusercontent.com/ZHB/switzerland-geojson/master/country/switzerland.geojson"

print("Downloading high-resolution Switzerland map data...")
with urllib.request.urlopen(geojson_url) as url:
    swiss_geojson = json.loads(url.read().decode())

# Load custom font
font_path = 'Geneva.ttf'
custom_font = FontProperties(fname=font_path)

# Plotting
fig, ax = plt.subplots(figsize=(20, 15))

# Plot Switzerland border
for feature in swiss_geojson['features']:
    geom = shape(feature['geometry'])
    if geom.geom_type == 'Polygon':
        x, y = geom.exterior.xy
        ax.plot(x, y, color='#333333', linewidth=1)
        ax.fill(x, y, color='#f0f0f0')
    elif geom.geom_type == 'MultiPolygon':
        for poly in geom.geoms:
            x, y = poly.exterior.xy
            ax.plot(x, y, color='#333333', linewidth=1)
            ax.fill(x, y, color='#f0f0f0')

# Plot institutions
lats = [loc[0] for loc in locations.values()]
lons = [loc[1] for loc in locations.values()]
names = list(locations.keys())

# Increased marker size
ax.scatter(lons, lats, color='red', marker='o', s=200, label='Institutions', zorder=5)

# Add labels with custom font
texts = []
for i, txt in enumerate(names):
    texts.append(ax.text(lons[i], lats[i], txt, fontsize=14, fontproperties=custom_font))

# Adjust text to avoid overlap, removed arrows
print("Adjusting text labels...")
adjust_text(texts)

# Add title and remove axis
ax.set_title('Institutions in Switzerland', fontsize=30, fontproperties=custom_font)
ax.axis('off')

# Save
output_file = 'institutions_map.png'
plt.savefig(output_file, dpi=300, bbox_inches='tight')
print(f"Map saved to {output_file}")
