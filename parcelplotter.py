import io
import json
import re
import requests
import pandas as pd
import geopandas as gpd
import contextily as cx
import matplotlib.pyplot as plt
import ipyleaflet
import ipywidgets as widgets
from IPython.display import display
# import os
# for var in ["http_proxy", "https_proxy", "HTTP_PROXY", "HTTPS_PROXY", "all_proxy", "ALL_PROXY"]:
#     os.environ.pop(var, None)


headers = {
    "User-Agent": "berncoexplorer/1.0 (alex.wm.stevenson@proton.me; contact for tile usage)"
}



# ─── Shared configuration ────────────────────────────────────────────────────
PARCEL_SERVICE = (
    "https://assessormap.bernco.gov/server/rest/services/"
    "GIS/ASROnline_Public_Map/MapServer/22/query"
)
PAGE_SIZE = 2000

# Fields the attribute query pulls (same list as the original script)
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

# ─── Address normalization helpers ──────────────────────────────────────────
_SUFFIX_MAP = {
    "ST": "STREET", "STR": "STREET", "AVE": "AVENUE", "AV": "AVENUE",
    "RD": "ROAD", "DR": "DRIVE", "LN": "LANE", "CT": "COURT",
    "PL": "PLACE", "BLVD": "BOULEVARD", "CIR": "CIRCLE",
    "WAY": "WAY", "TER": "TERRACE", "TRCE": "TERRACE",
    "PKWY": "PARKWAY", "PKY": "PARKWAY", "HWY": "HIGHWAY",
    "LOOP": "LOOP", "TRL": "TRAIL",
}
_DIR_MAP = {
    "N": "NORTH", "S": "SOUTH", "E": "EAST", "W": "WEST",
    "NE": "NORTHEAST", "NW": "NORTHWEST",
    "SE": "SOUTHEAST", "SW": "SOUTHWEST",
}

def _normalize_street(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).upper()
    s = re.sub(r"[.,#]", " ", s)
    out = []
    for tok in s.split():
        if tok in _DIR_MAP:
            out.append(_DIR_MAP[tok])
        elif tok in _SUFFIX_MAP:
            out.append(_SUFFIX_MAP[tok])
        else:
            out.append(tok)
    return " ".join(out)

def _normalize_number(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


# ─── Step 1: fetch attributes (returns DataFrame, optionally saves Parquet) ─
# def fetch_parcel_attributes(nw_corner, se_corner, luc_msg_values,
#                             output_parquet=None):
#     """
#     Query the Bernalillo County parcel layer for parcels inside the bbox
#     whose LUC_MSG is in `luc_msg_values`. Returns a pandas DataFrame of
#     attributes. If `output_parquet` is given, also saves it there.
#     """
#     nw_lat, nw_lon = nw_corner
#     se_lat, se_lon = se_corner
#     envelope = {
#         "xmin": min(nw_lon, se_lon),
#         "ymin": min(nw_lat, se_lat),
#         "xmax": max(nw_lon, se_lon),
#         "ymax": max(nw_lat, se_lat),
#         "spatialReference": {"wkid": 4326},
#     }

#     if luc_msg_values:
#         quoted = ", ".join(f"'{v}'" for v in luc_msg_values)
#         where_clause = f"LUC_MSG IN ({quoted})"
#     else:
#         where_clause = "1=1"

#     all_rows = []
#     offset = 0
#     while True:
#         params = {
#             "where": where_clause,
#             "geometry": json.dumps(envelope),
#             "geometryType": "esriGeometryEnvelope",
#             "inSR": 4326,
#             "spatialRel": "esriSpatialRelIntersects",
#             "outFields": ",".join(PARCEL_FIELDS),
#             "returnGeometry": "false",
#             "f": "json",
#             # "resultRecordCount": 2000,
#             # "resultOffset": 2000,
#             "resultRecordCount": PAGE_SIZE,
#             "resultOffset": offset,
#         }
#         # resp = requests.get(PARCEL_SERVICE, params=params, timeout=120)
#         print("URL:", PARCEL_SERVICE)
#         print("Data length:", len(str(params)))
#         resp = requests.post(PARCEL_SERVICE, data=params, timeout=120)
#         # print(resp.status_code)
#         # print(resp.text[:500])  # first 500 chars
#         resp.raise_for_status()
#         data = resp.json()
#         if "error" in data:
#             raise RuntimeError(f"Service error: {data['error']}")

#         features = data.get("features", [])
#         all_rows.extend(f["attributes"] for f in features)

#         # if not data.get("exceededTransferLimit") and len(features) < PAGE_SIZE:
#         if not data.get("exceededTransferLimit"):
#             break
#         offset += PAGE_SIZE

#     df = pd.DataFrame(all_rows)
#     if df.empty:
#         return df

#     # Coerce numerics so Parquet preserves types
#     for col in ["ACREAGE", "LANDVALUE", "IMPTVALUE", "TOTVALUE",
#                 "IMPTTXBLE", "LANDTXBLE", "TOTTXBLE", "TAXYR"]:
#         if col in df.columns:
#             df[col] = pd.to_numeric(df[col], errors="coerce")

#     if output_parquet:
#         df.to_parquet(output_parquet, index=False)

#     return df

def fetch_parcel_attributes(nw_corner, se_corner, luc_msg_values, output_parquet=None):
    ids, oid_field = fetch_all_parcel_ids(nw_corner, se_corner, luc_msg_values)
    if not ids:
        return pd.DataFrame()

    all_rows = []
    # Split into batches of 2000 IDs and request attributes
    for i in range(0, len(ids), PAGE_SIZE):
        batch = ids[i:i + PAGE_SIZE]
        params = {
            "objectIds": ",".join(str(x) for x in batch),
            "outFields": ",".join(PARCEL_FIELDS),
            "returnGeometry": "false",
            "f": "json",
        }
        resp = requests.post(PARCEL_SERVICE, data=params, timeout=120)
        resp.raise_for_status()
        data = resp.json()
        if "error" in data:
            raise RuntimeError(f"Service error: {data['error']}")
        features = data.get("features", [])
        all_rows.extend(f["attributes"] for f in features)
        print(f"Attributes batch {i//PAGE_SIZE + 1}: {len(features)} rows")

    df = pd.DataFrame(all_rows)
    if df.empty:
        return df
    for col in ["ACREAGE", "LANDVALUE", "IMPTVALUE", "TOTVALUE",
                "IMPTTXBLE", "LANDTXBLE", "TOTTXBLE", "TAXYR"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    if output_parquet:
        df.to_parquet(output_parquet, index=False)
    return df


# ─── Step 2: fetch geometry for a given set of UPCs ─────────────────────────
# def fetch_parcel_geometry(nw_corner, se_corner, luc_msg_values):
#     """
#     Query the same layer but request polygons (GeoJSON) for parcels in the
#     bbox matching the LUC_MSG values. Returns a GeoDataFrame with UPC and
#     geometry only (join attributes separately).
#     """
#     nw_lat, nw_lon = nw_corner
#     se_lat, se_lon = se_corner
#     envelope = {
#         "xmin": min(nw_lon, se_lon),
#         "ymin": min(nw_lat, se_lat),
#         "xmax": max(nw_lon, se_lon),
#         "ymax": max(nw_lat, se_lat),
#         "spatialReference": {"wkid": 4326},
#     }

#     if luc_msg_values:
#         quoted = ", ".join(f"'{v}'" for v in luc_msg_values)
#         where_clause = f"LUC_MSG IN ({quoted})"
#     else:
#         where_clause = "1=1"

#     params = {
#         "where": where_clause,
#         "geometry": json.dumps(envelope),
#         "geometryType": "esriGeometryEnvelope",
#         "inSR": 4326,
#         "spatialRel": "esriSpatialRelIntersects",
#         "outFields": "UPC",
#         "returnGeometry": "true",
#         "outSR": 4326,
#         "f": "geojson",
#     }
#     # resp = requests.get(PARCEL_SERVICE, params=params, timeout=120)
#     resp = requests.post(PARCEL_SERVICE, data=params, timeout=120)
#     resp.raise_for_status()

#     # gpd.read_file accepts a file-like object; the raw bytes are GeoJSON
#     gdf = gpd.read_file(io.BytesIO(resp.content))
#     return gdf[["UPC", "geometry"]]

def fetch_parcel_geometry(nw_corner, se_corner, luc_msg_values):
    ids, oid_field = fetch_all_parcel_ids(nw_corner, se_corner, luc_msg_values)
    if not ids:
        return gpd.GeoDataFrame()

    all_gdfs = []
    for i in range(0, len(ids), PAGE_SIZE):
        batch = ids[i:i + PAGE_SIZE]
        params = {
            "objectIds": ",".join(str(x) for x in batch),
            "outFields": "UPC",
            "returnGeometry": "true",
            "outSR": 4326,
            "f": "geojson",
        }
        resp = requests.post(PARCEL_SERVICE, data=params, timeout=120)
        resp.raise_for_status()
        gdf_batch = gpd.read_file(io.BytesIO(resp.content))
        all_gdfs.append(gdf_batch)
        print(f"Geometry batch {i//PAGE_SIZE + 1}: {len(gdf_batch)} features")

    if not all_gdfs:
        return gpd.GeoDataFrame()
    gdf = pd.concat(all_gdfs, ignore_index=True)
    return gpd.GeoDataFrame(gdf, geometry="geometry", crs="EPSG:4326")

# ─── Step 3: plot (the function you actually call) ──────────────────────────
def plot_parcels_by_luc_bbox(nw_corner, se_corner, luc_msg_values,
                             output_parquet=None, figsize=(14, 12), zoom=15):
    """
    Fetch parcel attributes + geometry for a bounding box and LUC_MSG list,
    classify homeowner vs landlord by address match, and plot over a basemap.

    If `output_parquet` is provided, the attribute DataFrame is also saved
    there for later reuse.
    """
    # Attributes
    df = fetch_parcel_attributes(nw_corner, se_corner, luc_msg_values,
                                 output_parquet=output_parquet)
    if df.empty:
        raise ValueError("No parcels matched the bounding box and LUC_MSG filter.")

    # Geometry
    gdf = fetch_parcel_geometry(nw_corner, se_corner, luc_msg_values)
    if gdf.empty:
        raise ValueError("No polygons returned for the bounding box.")

    # Join
    merged = gdf.merge(df, on="UPC", how="inner")
    if merged.empty:
        raise ValueError("Geometry and attribute results had no UPCs in common.")

    # Classify
    situs_num = merged["SITUSNUM"].map(_normalize_number)
    situs_str = merged["SITUSSTR"].map(_normalize_street)
    own_num = merged["OWNHSENUM"].map(_normalize_number)
    own_str = merged["OWNSTR"].map(_normalize_street)
    same_address = (situs_num == own_num) & (situs_str == own_str)
    merged["owner_type"] = "landlord"
    merged.loc[same_address, "owner_type"] = "homeowner"

    # Plot (reproject to Web Mercator for basemap tiles)
    merged_wm = merged.to_crs(epsg=3857)
    fig, ax = plt.subplots(figsize=figsize)

    merged_wm[merged_wm["owner_type"] == "landlord"].plot(
        ax=ax, color="red", alpha=0.6, edgecolor="darkred",
        linewidth=0.5, label="Landlord"
    )
    merged_wm[merged_wm["owner_type"] == "homeowner"].plot(
        ax=ax, color="blue", alpha=0.6, edgecolor="darkblue",
        linewidth=0.5, label="Homeowner"
    )
    # cx.add_basemap(ax, zoom=zoom, source=cx.providers.OpenStreetMap.Mapnik)
    cx.add_basemap(
        ax,
        zoom=zoom,
        source=cx.providers.OpenStreetMap.Mapnik,
        headers=headers
    )

    ax.set_axis_off()
    ax.legend(loc="upper right")
    plt.title(f"Parcels by Owner Type ({len(merged)} total)")
    plt.tight_layout()
    plt.show()

    counts = merged["owner_type"].value_counts()
    print(f"Homeowner parcels: {counts.get('homeowner', 0)}")
    print(f"Landlord parcels:  {counts.get('landlord', 0)}")
    return merged

# ─── Color palette for LUC_MSG categories ──────────────────────────────────
LUC_COLORS = [
    "#e6194b", "#3cb44b", "#4363d8", "#f58231", "#911eb4",
    "#46f0f0", "#f032e6", "#bcf60c", "#fabebe", "#008080",
    "#e6beff", "#9a6324", "#fffac8", "#800000", "#aaffc3",
]


def build_parcel_explorer(default_center=(35.0844, -106.6504), default_zoom=11):
    """
    Build an interactive ipyleaflet UI with:
      - Clickable map to set NW/SE corners
      - Toggle for which corner is being set
      - Multi-select for LUC_MSG values
      - Fetch button to query the API and render colored parcels
    """
    # ── State ──────────────────────────────────────────────────────────────
    # We'll store the clicked corners as plain dicts so the widget
    # callbacks can read them without relying on closure scoping tricks.
    corner_state = {"nw": None, "se": None}
    selected_corner = widgets.ToggleButtons(
        options=["NW", "SE"],
        value="NW",
        description="Set:",
        button_style="info",
    )

    # ── Coordinate display boxes ───────────────────────────────────────────
    nw_lat_box = widgets.Text(value="", description="NW lat:", disabled=True, layout=widgets.Layout(width="300px"))
    nw_lon_box = widgets.Text(value="", description="NW lon:", disabled=True, layout=widgets.Layout(width="300px"))
    se_lat_box = widgets.Text(value="", description="SE lat:", disabled=True, layout=widgets.Layout(width="300px"))
    se_lon_box = widgets.Text(value="", description="SE lon:", disabled=True, layout=widgets.Layout(width="300px"))

    # ── Map ────────────────────────────────────────────────────────────────
    m = ipyleaflet.Map(
        center=default_center,
        zoom=default_zoom,
        scroll_wheel_zoom=True,
        dragging=False,
        layout=widgets.Layout(width="100%", height="600px"),
    )
    # Inside build_parcel_explorer(...), after creating m:
    # After creating m, add a global click layer
    click_catcher = ipyleaflet.GeoJSON(
        data={"type": "FeatureCollection", "features": []},
        style={"opacity": 0, "fillOpacity": 0, "weight": 0},
        name="click_catcher",
    )
    m.add_layer(click_catcher)

    def on_map_click(event, feature, **kwargs):
        lat, lon = event  # on_click of GeoJSON layer passes (lat, lon) coordinates
        if selected_corner.value == "NW":
            corner_state["nw"] = (lat, lon)
            nw_lat_box.value = f"{lat:.6f}"
            nw_lon_box.value = f"{lon:.6f}"
            # ... update marker ...
        else:
            corner_state["se"] = (lat, lon)
            se_lat_box.value = f"{lat:.6f}"
            se_lon_box.value = f"{lon:.6f}"
            # ... update marker ...

    click_catcher.on_click(on_map_click)

    drag_lock = widgets.ToggleButton(
        value=False,
        description="Lock Map for Clicking",
        tooltip="Turn ON to disable panning so clicks register as corner selections",
        button_style="warning",
        layout=widgets.Layout(width="300px"),
    )

    def toggle_dragging(change):
        m.dragging = not change["new"]  # Lock ON -> dragging OFF

    drag_lock.observe(toggle_dragging, names="value")
    # Use a basemap that is less likely to be blocked than raw OSM tiles.
    # Esri WorldTopoMap is a common fallback; you can swap for CartoDB.Positron.
    m.basemap = ipyleaflet.basemaps.Esri.WorldTopoMap

    # Marker layers so the user can see where they clicked
    nw_marker = None
    se_marker = None

    # ── Click handler ──────────────────────────────────────────────────────
    def handle_click(**kwargs):
        nonlocal nw_marker, se_marker
        if kwargs.get("type") != "click":
            return
        lat, lon = kwargs.get("coordinates")  # (lat, lon) order in ipyleaflet

        if selected_corner.value == "NW":
            corner_state["nw"] = (lat, lon)
            nw_lat_box.value = f"{lat:.6f}"
            nw_lon_box.value = f"{lon:.6f}"
            if nw_marker is not None:
                m.remove_layer(nw_marker)
            nw_marker = ipyleaflet.Marker(location=(lat, lon), draggable=False, title="NW")
            m.add_layer(nw_marker)
        else:
            corner_state["se"] = (lat, lon)
            se_lat_box.value = f"{lat:.6f}"
            se_lon_box.value = f"{lon:.6f}"
            if se_marker is not None:
                m.remove_layer(se_marker)
            se_marker = ipyleaflet.Marker(location=(lat, lon), draggable=False, title="SE")
            m.add_layer(se_marker)

    m.on_interaction(handle_click)

    # ── LUC multi-select ───────────────────────────────────────────────────
    # Populate with a distinct query when the UI starts.
    # If the query is slow or fails, fall back to an empty list.
    try:
        luc_values = _fetch_distinct_luc_values()
    except Exception as e:
        print(f"Could not fetch LUC_MSG list: {e}")
        luc_values = []

    luc_selector = widgets.SelectMultiple(
        options=luc_values,
        value=[],
        description="LUC_MSG:",
        layout=widgets.Layout(width="300px", height="400px"),
        disabled=False,
    )

    # ── Fetch button and output area ───────────────────────────────────────
    fetch_button = widgets.Button(description="Fetch & Plot", button_style="success")
    output_area = widgets.Output()

    def on_fetch_clicked(_):
        with output_area:
            output_area.clear_output()
            nw = corner_state["nw"]
            se = corner_state["se"]
            lus = list(luc_selector.value)

            if not nw or not se:
                print("Please set both NW and SE corners by clicking the map.")
                return
            if not lus:
                print("Please select at least one LUC_MSG value.")
                return

            print(f"NW: {nw}  SE: {se}")
            print(f"LUC_MSG: {lus}")
            print("Fetching...")

            try:
                df = fetch_parcel_attributes(nw, se, lus)
                if df.empty:
                    print("No parcels matched.")
                    return
                gdf = fetch_parcel_geometry(nw, se, lus)
                if gdf.empty:
                    print("No polygons returned.")
                    return

                merged = gdf.merge(df, on="UPC", how="inner")
                print(f"Plotting {len(merged)} parcels...")

                # Remove any previous parcel layer
                for layer in list(m.layers):
                    if isinstance(layer, ipyleaflet.GeoData):
                        m.remove_layer(layer)

                # Render each LUC_MSG group as its own GeoData layer
                for i, luc in enumerate(lus):
                    subset = merged[merged["LUC_MSG"] == luc]
                    if subset.empty:
                        continue
                    color = LUC_COLORS[i % len(LUC_COLORS)]
                    geo_layer = ipyleaflet.GeoData(
                        geo_dataframe=subset[["UPC", "geometry"]],
                        style={
                            "color": color,
                            "fillColor": color,
                            "opacity": 1.0,
                            "fillOpacity": 0.6,
                            "weight": 1.5,
                        },
                        name=luc,
                    )
                    m.add_layer(geo_layer)

                # Add a legend-ish layer control so the user can toggle groups
                if not any(isinstance(c, ipyleaflet.LayersControl) for c in m.controls):
                    m.add_control(ipyleaflet.LayersControl(position="topright"))

                print("Done.")
            except Exception as e:
                print(f"Error: {e}")

    fetch_button.on_click(on_fetch_clicked)

    # ── Layout assembly ────────────────────────────────────────────────────
    control_panel = widgets.VBox([
        widgets.HTML("<h3>Corner Selection</h3>"),
        selected_corner,
        drag_lock,
        nw_lat_box, nw_lon_box,
        se_lat_box, se_lon_box,
        widgets.HTML("<hr>"),
        widgets.HTML("<h3>LUC_MSG Filter</h3>"),
        luc_selector,
        fetch_button,
        output_area,
    ], layout=widgets.Layout(width="350px", padding="10px"))

    ui = widgets.HBox([control_panel, m])
    display(ui)


def _fetch_distinct_luc_values():
    """Query the ArcGIS service for unique LUC_MSG values."""
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
    values = sorted({
        f["attributes"]["LUC_MSG"]
        for f in data.get("features", [])
        if f.get("attributes", {}).get("LUC_MSG")
    })
    return values

def fetch_all_parcel_ids(nw_corner, se_corner, luc_msg_values):
    """Get all matching parcel IDs at once using returnIdsOnly."""
    nw_lat, nw_lon = nw_corner
    se_lat, se_lon = se_corner
    envelope = {
        "xmin": min(nw_lon, se_lon),
        "ymin": min(nw_lat, se_lat),
        "xmax": max(nw_lon, se_lon),
        "ymax": max(nw_lat, se_lat),
        "spatialReference": {"wkid": 4326},
    }

    if luc_msg_values:
        quoted = ", ".join(f"'{v}'" for v in luc_msg_values)
        where_clause = f"LUC_MSG IN ({quoted})"
    else:
        where_clause = "1=1"

    params = {
        "where": where_clause,
        "geometry": json.dumps(envelope),
        "geometryType": "esriGeometryEnvelope",
        "inSR": 4326,
        "spatialRel": "esriSpatialRelIntersects",
        "returnIdsOnly": "true",       # Key: only return IDs
        "returnGeometry": "false",
        "f": "json",
    }
    resp = requests.post(PARCEL_SERVICE, data=params, timeout=120)
    resp.raise_for_status()
    data = resp.json()
    if "error" in data:
        raise RuntimeError(f"Service error: {data['error']}")
    # The ID field name returned by the service might be OBJECTID or similar; dynamically read it from the response
    oid_field = data.get("objectIdFieldName", "OBJECTID")
    return data.get("objectIds", []), oid_field