"""
Bernalillo County Assessor — Countywide Parcel Search by LUC_MSG
==================================================================
Pulls ALL parcels matching an exact Land Use Code description
(LUC_MSG field, e.g. 'MOBILE HOME PARK') from the County Assessor's
own parcel layer — NOT the City of Albuquerque mirror, so this
covers the WHOLE county, not just parcels inside city limits.

Optionally restrict the search to a rectangular area using NW / SE
corner lat/lon coordinates (our usual protocol). Leave both blank
for a full countywide pull.

SETUP / FIRST RUN
------------------
1. pip3 install requests openpyxl
2. Run once with DISCOVER_FIELDS_ONLY = True to print every field
   name available on the layer. Use that list to confirm the exact
   spelling/casing of LUC_MSG (and any other fields you want in the
   output), then set DISCOVER_FIELDS_ONLY = False and fill in
   LUC_MSG_VALUE below.
3. python3 bernco_countywide_search-by_LUC_MSG.py
"""

import sys
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ─── CONFIGURATION ──────────────────────────────────────────────────────────

# Set True to just print available field names and exit (already confirmed).
DISCOVER_FIELDS_ONLY = False

# Exact LUC_MSG value to match. Confirmed ALL CAPS text from a live sample
# (e.g. "VACANT RESIDENTIAL"). Edit this single line to pull any other
# land-use category — just change the text to the exact LUC_MSG wording.
LUC_MSG_VALUE = "MULTIFAMILY RES U4U"

# Leave both as None for a full countywide search.
# Otherwise supply EITHER:
#   - a lat/lon tuple (WGS84 decimal degrees), e.g. (35.2200, -106.7800)
#   - OR a plain address string, e.g. "10500 Eagle Rock Ave NE, Albuquerque, NM"
#     (the script will geocode it automatically via the free US Census
#     geocoder — no API key needed)
NW_CORNER = None   # (lat, lon) OR "address string"
SE_CORNER = None   # (lat, lon) OR "address string"

# Fields to pull and write to the spreadsheet. These are confirmed real
# field names from the live layer (discovery run). Add/remove as needed.
PARCEL_FIELDS = [
    "UPC",
    "SITUSNUM", "SITUSSTR", "SITUSSTRTY", "SITUSDIREC",
    "SITUSCITY", "SITUSSTATE", "SITUSZIP", "OWNER",
    "OWNHSENUM", "OWNSTR", "OWNSTRTYPE", "OWNUNIT", "OWNUNITNO",
    "OWNDIRECT", "OWNBOX", "OWNCITY", "OWNSTATE", "OWNZIPCODE", "OWNZIP4",
    "LEGALDESC", "PROPCLASS",
    "LUC", "LUC_MSG",
    "ACREAGE", "LANDVALUE", "IMPTVALUE", "TOTVALUE", "IMPTTXBLE", "LANDTXBLE", "TOTTXBLE",
    "TAXDIST", "TAXYR",
]

OUTPUT_FILE = "BernCo_Countywide_MultiFamilyResU4U.xlsx"

# ─── SERVICE CONFIG (do not normally need to edit) ─────────────────────────

PARCEL_SERVICE = (
    "https://assessormap.bernco.gov/server/rest/services/"
    "GIS/ASROnline_Public_Map/MapServer/22/query"
)
PAGE_SIZE = 2000   # this server's typical MaxRecordCount; auto-paginates

# ─────────────────────────────────────────────────────────────────────────


def discover_fields():
    """Query the layer for a single record and print all field names
    and a sample value, so you can confirm the LUC_MSG field name."""
    params = {
        "where": "1=1",
        "outFields": "*",
        "returnGeometry": "false",
        "f": "json",
        "resultRecordCount": 1,
    }
    print("Querying layer for field discovery...\n")
    resp = requests.get(PARCEL_SERVICE, params=params, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    if "error" in data:
        raise RuntimeError(f"Service error: {data['error']}")
    if not data.get("fields"):
        print("No 'fields' returned — raw response below:")
        print(data)
        return
    print(f"{'FIELD NAME':30s} {'TYPE':20s} SAMPLE VALUE")
    print("-" * 80)
    sample = data["features"][0]["attributes"] if data.get("features") else {}
    for f in data["fields"]:
        name = f.get("name", "")
        ftype = f.get("type", "")
        val = sample.get(name, "")
        print(f"{name:30s} {ftype:20s} {val}")
    print("\nLook for the field that holds land-use descriptions like")
    print("'MOBILE HOME PARK' (likely named LUC_MSG, LUC, or similar).")
    print("Update PARCEL_FIELDS and LUC_MSG_VALUE in the CONFIGURATION")
    print("section above, set DISCOVER_FIELDS_ONLY = False, and re-run.")


def geocode_address(address):
    """Resolve a plain address string to (lat, lon) using the free US
    Census Bureau geocoder (no API key required)."""
    print(f"  Geocoding '{address}'... ", end="", flush=True)
    resp = requests.get(
        "https://geocoding.geo.census.gov/geocoder/locations/onelineaddress",
        params={"address": address, "benchmark": "Public_AR_Current", "format": "json"},
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    matches = data.get("result", {}).get("addressMatches", [])
    if not matches:
        raise RuntimeError(
            f"Could not geocode address: '{address}'. Try adding city/state/zip, "
            f"or supply a (lat, lon) tuple directly instead."
        )
    coords = matches[0]["coordinates"]
    lat, lon = coords["y"], coords["x"]
    print(f"-> ({lat:.5f}, {lon:.5f})")
    return lat, lon


def resolve_corner(corner):
    """Accept either a (lat, lon) tuple or an address string and return
    a (lat, lon) tuple."""
    if corner is None:
        return None
    if isinstance(corner, (tuple, list)) and len(corner) == 2:
        return tuple(corner)
    if isinstance(corner, str):
        return geocode_address(corner)
    raise ValueError(f"Unrecognized corner format: {corner!r}")


def build_bbox_geometry():
    """Build an ArcGIS REST envelope geometry from NW/SE corners, which
    may be lat/lon tuples or address strings (auto-geocoded)."""
    nw = resolve_corner(NW_CORNER)
    se = resolve_corner(SE_CORNER)
    if not nw or not se:
        return None, None, None, None
    nw_lat, nw_lon = nw
    se_lat, se_lon = se
    envelope = {
        "xmin": min(nw_lon, se_lon),
        "ymin": min(nw_lat, se_lat),
        "xmax": max(nw_lon, se_lon),
        "ymax": max(nw_lat, se_lat),
        "spatialReference": {"wkid": 4326},
    }
    import json
    return json.dumps(envelope), 4326, nw, se


def fetch_all_pages():
    where_clause = f"LUC_MSG = '{LUC_MSG_VALUE}'"
    geometry, geom_sr, nw_resolved, se_resolved = build_bbox_geometry()

    all_rows = []
    offset = 0
    page = 1
    print(f"\nWHERE: {where_clause}")
    if geometry:
        print(f"BBOX: NW{nw_resolved}  SE{se_resolved}")
    else:
        print("BBOX: none (full countywide search)")
    print()

    while True:
        params = {
            "where": where_clause,
            "outFields": ",".join(PARCEL_FIELDS),
            "returnGeometry": "false",
            "f": "json",
            "resultRecordCount": PAGE_SIZE,
            "resultOffset": offset,
        }
        if geometry:
            params.update({
                "geometry": geometry,
                "geometryType": "esriGeometryEnvelope",
                "inSR": geom_sr,
                "spatialRel": "esriSpatialRelIntersects",
            })

        print(f"  Page {page} (offset {offset})... ", end="", flush=True)
        resp = requests.get(PARCEL_SERVICE, params=params, timeout=120)
        resp.raise_for_status()
        data = resp.json()
        if "error" in data:
            raise RuntimeError(f"Service error: {data['error']}")
        features = data.get("features", [])
        rows = [f["attributes"] for f in features]
        all_rows.extend(rows)
        print(f"{len(rows)} records")

        if not data.get("exceededTransferLimit") and len(rows) < PAGE_SIZE:
            break
        offset += PAGE_SIZE
        page += 1

    return all_rows


def build_xlsx(all_rows, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Parcels"

    header_fill = PatternFill("solid", fgColor="185FA5")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    for col_idx, header in enumerate(PARCEL_FIELDS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    data_font = Font(name="Arial", size=10)
    for row_idx, record in enumerate(all_rows, start=2):
        for col_idx, field in enumerate(PARCEL_FIELDS, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=record.get(field, ""))
            cell.font = data_font

    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 10), 40)

    ws.freeze_panes = "A2"
    wb.save(output_path)
    print(f"\nSaved: {output_path}  ({len(all_rows)} rows)")


def main():
    print("\nBernalillo County Assessor — Countywide Parcel Search by LUC_MSG")

    if DISCOVER_FIELDS_ONLY:
        discover_fields()
        sys.exit(0)

    rows = fetch_all_pages()
    if not rows:
        print("\nNo parcels found. Check LUC_MSG_VALUE matches the field's")
        print("exact text (case/spacing) — re-run with DISCOVER_FIELDS_ONLY")
        print("= True and inspect a few real records' LUC_MSG values if unsure.")
        return

    build_xlsx(rows, output_path=OUTPUT_FILE)


if __name__ == "__main__":
    main()
