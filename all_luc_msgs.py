import requests
import json

PARCEL_SERVICE = (
    "https://assessormap.bernco.gov/server/rest/services/"
    "GIS/ASROnline_Public_Map/MapServer/22/query"
)

params = {
    "where": "1=1",
    "outFields": "LUC_MSG",
    "returnDistinctValues": "true",
    "returnGeometry": "false",
    "f": "json",
}

resp = requests.get(PARCEL_SERVICE, params=params, timeout=60)
resp.raise_for_status()
data = resp.json()

# The unique values are in the features list
unique_values = sorted(
    {f["attributes"]["LUC_MSG"] for f in data.get("features", []) if f["attributes"].get("LUC_MSG")}
)

print(f"Found {len(unique_values)} unique LUC_MSG values:")
for v in unique_values:
    # Highlight ones that look like single-family residential
    marker = " <-- SINGLE FAMILY" if "SINGLE" in v.upper() or "SFR" in v.upper() else ""
    print(f"  {v}{marker}")