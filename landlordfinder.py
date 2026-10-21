import re
import pandas as pd

# Common street suffix/type abbreviations → canonical form.
# Extend as you encounter new variants in your data.
_SUFFIX_MAP = {
    "ST": "STREET", "STR": "STREET",
    "AVE": "AVENUE", "AV": "AVENUE",
    "RD": "ROAD",
    "DR": "DRIVE",
    "LN": "LANE",
    "CT": "COURT",
    "PL": "PLACE",
    "BLVD": "BOULEVARD",
    "CIR": "CIRCLE",
    "WAY": "WAY",
    "TER": "TERRACE", "TRCE": "TERRACE",
    "PKWY": "PARKWAY", "PKY": "PARKWAY",
    "HWY": "HIGHWAY",
    "LOOP": "LOOP",
    "TRL": "TRAIL",
}

# Directional abbreviations → canonical form
_DIR_MAP = {
    "N": "NORTH", "S": "SOUTH", "E": "EAST", "W": "WEST",
    "NE": "NORTHEAST", "NW": "NORTHWEST",
    "SE": "SOUTHEAST", "SW": "SOUTHWEST",
}

def _normalize_street(value):
    """Uppercase, strip punctuation, expand common abbreviations,
    and collapse whitespace so '123 N. Main St.' == '123 NORTH MAIN STREET'."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).upper()
    s = re.sub(r"[.,#]", " ", s)          # drop punctuation that varies by source
    tokens = s.split()
    out = []
    for tok in tokens:
        if tok in _DIR_MAP:
            out.append(_DIR_MAP[tok])
        elif tok in _SUFFIX_MAP:
            out.append(_SUFFIX_MAP[tok])
        else:
            out.append(tok)
    return " ".join(out)

def _normalize_number(value):
    """House numbers come through as int, float, or string. Coerce to a
    clean string of digits, dropping trailing '.0' from float coercion."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip()
    # Handle '123.0' style floats
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s

def find_landlord_owned(parquet_path):
    """Return a DataFrame of parcels where the property (situs) address
    does not match the owner's mailing address, plus summary stats."""
    df = pd.read_parquet(parquet_path)

    required = ["SITUSNUM", "SITUSSTR", "OWNHSENUM", "OWNSTR"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"Missing expected columns: {missing}")

    situs_num = df["SITUSNUM"].map(_normalize_number)
    situs_str = df["SITUSSTR"].map(_normalize_street)
    own_num   = df["OWNHSENUM"].map(_normalize_number)
    own_str   = df["OWNSTR"].map(_normalize_street)

    # A parcel is "landlord-owned" (in this heuristic) when either the
    # house number OR the street name differs. Requiring both to match
    # avoids flagging parcels whose owner address is just incomplete.
    number_mismatch = situs_num != own_num
    street_mismatch = situs_str != own_str
    is_mismatch = number_mismatch | street_mismatch

    result = df[is_mismatch].copy()
    result["_SITUS_NORMALIZED"] = situs_num + " " + situs_str
    result["_OWNER_NORMALIZED"] = own_num + " " + own_str

    total = len(df)
    flagged = len(result)
    print(f"Total parcels:            {total}")
    print(f"Address mismatch (landlord heuristic): {flagged} "
          f"({flagged / total:.1%})" if total else "No parcels.")
    print(f"  -- house number differs: {number_mismatch.sum()}")
    print(f"  -- street name differs:  {street_mismatch.sum()}")

    return result

# Example usage
if __name__ == "__main__":
    landlords = find_landlord_owned("myblock.parquet")
    print(landlords[["UPC", "SITUSNUM", "SITUSSTR",
                     "OWNHSENUM", "OWNSTR", "OWNER"]].head(20))