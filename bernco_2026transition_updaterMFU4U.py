#!/usr/bin/env python3
"""
BernCo 2026 Transition — Property Profile Updater
---------------------------------------------------
Reads UPC codes from BernCo_Countywide_VacantResidential_2026_Transition.xlsx
(column A, data starting at row 3) and for each one navigates to the
Bernalillo County Assessor property profile page to update sixteen fields
with 2026 data:

    Column I  (9)  — OWNER      (current owner name)
    Columns J–T (10–20) — owner MAILING ADDRESS, parsed into components:
        J  OWNHSENUM   K  OWNSTR      L  OWNSTRTYPE  M  OWNUNIT
        N  OWNUNITNO   O  OWNDIRECT   P  OWNBOX      Q  OWNCITY
        R  OWNSTATE    S  OWNZIPCODE  T  OWNZIP4
    Column X  (24) — LUC_MSG    (land use code description)
    Column AH (34) — SQ FT      (primary building square footage)
    Column AI (35) — YR BUILT   (year built)
    Column AJ (36) — STYLE      (building style)

Any existing 2025 data in those columns is intentionally overwritten
with the 2026 values. Rows already processed by this script are skipped
on re-run (STYLE column used as the completion sentinel — "(none)"
is written when the assessor record genuinely has no style listed).

SETUP (one time):
    pip3 install --upgrade selenium beautifulsoup4 openpyxl

    You need Google Chrome installed.

USAGE:
    python3 bernco_2026transition_updater.py

NOTES:
    - Saves the spreadsheet every SAVE_XLSX_EVERY completed rows and
      on Ctrl+C. Safe to stop and resume.
    - Chrome is restarted every RESTART_DRIVER_EVERY rows to avoid
      the ~48-50 request per session rate limit.
    - A checkpoint CSV is written after every row so scraped data is
      never lost even between spreadsheet saves.
"""

import csv
import os
import re
import time
import tempfile
from datetime import datetime
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, WebDriverException
)
from bs4 import BeautifulSoup

# ─── CONFIGURATION ──────────────────────────────────────────────────────────

XLSX_PATH  = "BernCo_Countywide_MultiFamilyResU4U.xlsx"
SHEET_NAME = "Parcels"

HEADER_ROW = 1       # row containing column labels
START_ROW  = 2       # first data row  (CHECK: must match your sheet)
END_ROW    = None    # None = process through last row in sheet
TEST_LIMIT = None    # None = no limit; set to e.g. 5 to test before full run

HEADLESS               = False
PAUSE_BETWEEN_PARCELS  = 6.0   # seconds — do not go below 5
RATE_LIMIT_WAIT_SECONDS = 120  # wait after hitting OverLimit (fresh Chrome restarts reset the session)
MAX_RATE_LIMIT_RETRIES  = 3
RESTART_DRIVER_EVERY    = 40   # restart Chrome every N rows (server session limit ≈ 48-50)
SAVE_XLSX_EVERY         = 20   # full spreadsheet save every N completed rows
MAX_CONSECUTIVE_FAILURES = 5   # stop run if this many rows in a row all fail

ERROR_LOG_PATH  = "update_errors_2026transition.csv"
# NOTE: new filename — the checkpoint now has extra columns, so it must not
# be appended to an older 7-column checkpoint file.
CHECKPOINT_CSV  = "checkpoint_2026transition_v2.csv"

# Set True ONCE to go back over rows that a previous run already completed
# (STYLE filled) but whose owner-mailing columns (J–T) are still empty.
# Leave False for normal runs.
BACKFILL_MAILING = False

# ─── COLUMN ASSIGNMENTS (1-indexed) ─────────────────────────────────────────

UPC_COL        = 1    # A
SITUSNUM_COL   = 2    # B  — for address-search fallback
SITUSSTR_COL   = 3    # C  — for address-search fallback
OWNER_COL      = 9    # I
OWNHSENUM_COL  = 10   # J
OWNSTR_COL     = 11   # K
OWNSTRTYPE_COL = 12   # L
OWNUNIT_COL    = 13   # M
OWNUNITNO_COL  = 14   # N
OWNDIRECT_COL  = 15   # O
OWNBOX_COL     = 16   # P
OWNCITY_COL    = 17   # Q
OWNSTATE_COL   = 18   # R
OWNZIPCODE_COL = 19   # S
OWNZIP4_COL    = 20   # T
LUC_MSG_COL    = 24   # X
SQFT_COL       = 34   # AH
YEARBUILT_COL  = 35   # AI
STYLE_COL      = 36   # AJ  — also used as the "row done" sentinel

# key in the fields dict -> spreadsheet column
MAIL_COL_MAP = {
    "ownhsenum":  OWNHSENUM_COL,
    "ownstr":     OWNSTR_COL,
    "ownstrtype": OWNSTRTYPE_COL,
    "ownunit":    OWNUNIT_COL,
    "ownunitno":  OWNUNITNO_COL,
    "owndirect":  OWNDIRECT_COL,
    "ownbox":     OWNBOX_COL,
    "owncity":    OWNCITY_COL,
    "ownstate":   OWNSTATE_COL,
    "ownzipcode": OWNZIPCODE_COL,
    "ownzip4":    OWNZIP4_COL,
}

STYLE_SENTINEL = "(none)"   # written when profile exists but has no Style listed

# ─── SERVICE ────────────────────────────────────────────────────────────────

SEARCH_URL = (
    "https://assessor.bernco.gov/public.access/search/"
    "commonsearch.aspx?mode=realprop"
)

# ────────────────────────────────────────────────────────────────────────────


class OverLimitError(RuntimeError):
    """Raised when OverLimit.aspx is encountered at ANY navigation step."""
    pass


# ─── HELPERS ────────────────────────────────────────────────────────────────

def clean_int(text):
    """Strip formatting and return int, or None if blank/unparseable."""
    if not text:
        return None
    text = text.strip().replace(",", "").replace("$", "")
    if text in ("", "&nbsp;"):
        return None
    try:
        return int(float(text))
    except ValueError:
        return None


def clean_str(text):
    """Strip whitespace; return None if blank."""
    if not text:
        return None
    text = text.strip()
    return text if text else None


def write_cell(ws, row, col, value):
    """Write a cell value explicitly (works correctly for None too)."""
    ws.cell(row=row, column=col).value = value


def atomic_save(wb, path):
    """Save to a temp file, verify it opens, then atomically replace."""
    dir_name = os.path.dirname(os.path.abspath(path))
    fd, tmp = tempfile.mkstemp(suffix=".tmp.xlsx", dir=dir_name)
    os.close(fd)
    try:
        wb.save(tmp)
        from openpyxl import load_workbook as _lw
        _lw(tmp, read_only=True).close()
        os.replace(tmp, path)
    except Exception:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


# ─── MAILING-ADDRESS PARSING ────────────────────────────────────────────────

STREET_SUFFIXES = {
    "ST", "AVE", "RD", "DR", "BLVD", "LN", "CT", "PL", "WAY", "CIR", "TRL",
    "LOOP", "PKWY", "HWY", "TER", "PASS", "PATH", "RUN", "ROW", "SQ", "ALY",
    "XING", "CV", "CIRCLE", "COURT", "DRIVE", "LANE", "PLACE", "ROAD",
    "STREET", "AVENUE", "BOULEVARD", "TRAIL", "PARKWAY", "HIGHWAY", "TRAIL",
    "CT.", "ST.", "AVE.", "RD.", "DR.",
}
DIRECTIONALS = {"N", "S", "E", "W", "NE", "NW", "SE", "SW"}

BOX_RE  = re.compile(r"^(?:P\.?\s?O\.?\s*BOX|POST OFFICE BOX|BOX)\s*#?\s*(\S+)")
UNIT_RE = re.compile(
    r"(?:[\s,]+(APT|APARTMENT|UNIT|STE|SUITE|BLDG|SPC|SPACE|LOT|RM|ROOM|FL|FLOOR)\b\.?\s*#?\s*"
    r"|[\s,]*#\s*)([A-Z0-9-]+)$"
)
HOUSE_RE = re.compile(r"^(\d+[A-Z]?(?:\s+1/2)?)\s+(.+)$")
CSZ_RE = re.compile(
    r"^(?:(?P<pre>.*),\s*)?(?P<city>[A-Z .'-]+?),?\s+(?P<state>[A-Z]{2})\s+"
    r"(?P<zip>\d{5})(?:-?(?P<zip4>\d{4}))?$"
)

_mailing_debug_written = False


def _cell_lines(cell):
    txt = cell.get_text(separator="\n").replace("\xa0", " ")
    return [re.sub(r"\s+", " ", l).strip() for l in txt.split("\n") if l.strip()]


LABEL_WORDS = {
    "owner", "owner name", "owner(s)", "current owner", "owner information",
    "mailing address", "mailing", "address", "owner address",
    "city state zip", "city, state zip", "mailing city state zip",
    "city/state/zip",
}
OWNER_ANCHORS = {"owner", "owner name", "owner(s)", "current owner",
                 "owner information"}


def _norm(l):
    return re.sub(r"\s+", " ", l.replace("\xa0", " ")).strip()


def page_text_lines(soup):
    return [x for x in (_norm(l) for l in soup.get_text("\n").split("\n")) if x]


def extract_mailing_lines(soup):
    """Return the owner's mailing address as a list of text lines, or [].

    Strategy 1: a table row labelled 'Mailing Address' (plus continuation rows).
    Strategy 2 (layout-independent): find the 'Owner' label in the page text,
    then look at the next few text lines for one shaped like
    'CITY ST 87108[-1234]'.  The lines in between are the owner name(s) and
    street line(s); label words are dropped."""
    # Strategy 1
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        for i, row in enumerate(rows):
            cells = row.find_all("td")
            if len(cells) != 2:
                continue
            lbl = cells[0].get_text(strip=True).rstrip(":").lower()
            if lbl.startswith("mailing address") or lbl == "mailing":
                lines = _cell_lines(cells[1])
                for nxt in rows[i + 1:]:
                    c = nxt.find_all("td")
                    if len(c) != 2:
                        break
                    l = c[0].get_text(strip=True).rstrip(":").lower()
                    if l == "" or "city" in l or "zip" in l or "state" in l:
                        lines.extend(_cell_lines(c[1]))
                    else:
                        break
                lines = [x for x in lines if x]
                if any(CSZ_RE.match(x.upper()) for x in lines):
                    return lines

    # Strategy 2
    text = page_text_lines(soup)
    for i, line in enumerate(text):
        if line.rstrip(":").lower() not in OWNER_ANCHORS:
            continue
        window = text[i + 1:i + 10]
        for j, w in enumerate(window):
            if CSZ_RE.match(w.upper()):
                block = [x for x in window[:j + 1]
                         if x.rstrip(":").lower() not in LABEL_WORDS]
                return block
    return []


def _parse_street(street, out):
    """Split one street line into house number, street, type, directional,
    unit and PO box (written into the dict `out`)."""
    street = re.sub(r"\s+", " ", street).strip().upper()
    m = BOX_RE.match(street)
    if m:
        out["ownbox"] = m.group(1)
        return

    m = UNIT_RE.search(street)
    if m:
        out["ownunit"]   = (m.group(1) or "#")
        out["ownunitno"] = m.group(2)
        street = street[:m.start()].strip(" ,")

    m = HOUSE_RE.match(street)
    if m:
        out["ownhsenum"] = m.group(1)
        rest = m.group(2)
    else:
        rest = street

    tokens = rest.split()
    if len(tokens) > 1 and tokens[-1] in DIRECTIONALS:
        out["owndirect"] = tokens.pop()
    if len(tokens) > 1 and tokens[-1].rstrip(".") in STREET_SUFFIXES:
        out["ownstrtype"] = tokens.pop().rstrip(".")
    if not out["owndirect"] and len(tokens) > 1 and tokens[0] in DIRECTIONALS:
        out["owndirect"] = tokens.pop(0)
    out["ownstr"] = " ".join(tokens) or None


UNIT_FIELD_RE = re.compile(r"^(?:([A-Z]+)\.?\s*)?#?\s*([A-Z0-9-]+)$")


def _owner_table_pairs(soup):
    """label -> value for the 'Current Owner' table (falls back to the
    'Ownership for Tax Year Selected' table).  First occurrence wins."""
    for tid in ("Current Owner", "Ownership for Tax Year Selected"):
        table = soup.find("table", id=tid)
        if not table:
            continue
        pairs = {}
        for tr in table.find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) == 2:
                lbl = _norm(tds[0].get_text(" ", strip=True)).rstrip(":").lower()
                val = _norm(tds[1].get_text(" ", strip=True))
                if lbl and lbl not in pairs:
                    pairs[lbl] = val
        if (pairs.get("owner mailing address") or pairs.get("city")
                or pairs.get("other mailing address")):
            return pairs
    return None


def extract_owner_mailing(soup):
    """Returns (parsed_dict, raw_text).  Reads the labelled rows
    Owner Mailing Address / Unit / City / State / Zip Code from the
    'Current Owner' table.  Falls back to the layout-independent text
    search if that table isn't present.  raw_text is '' if nothing found."""
    pairs = _owner_table_pairs(soup)
    if pairs is None:
        lines = extract_mailing_lines(soup)
        return parse_mailing_address(lines), " | ".join(lines)

    out = {k: None for k in MAIL_COL_MAP}
    street = (pairs.get("owner mailing address") or "").upper()
    other  = (pairs.get("other mailing address") or "").upper()
    unit   = (pairs.get("unit") or "").upper()
    city   = (pairs.get("city") or "").upper()
    state  = (pairs.get("state") or "").upper()
    zipraw = (pairs.get("zip code") or "").upper()

    if street:
        _parse_street(street, out)
    # A PO box may be given on the 'Other Mailing Address' line
    if not out["ownbox"] and other:
        m = BOX_RE.match(other)
        if m:
            out["ownbox"] = m.group(1)

    if unit:
        m = UNIT_FIELD_RE.match(unit)
        if m and m.group(2):
            out["ownunit"]   = m.group(1)      # e.g. SUITE, APT (may be None)
            out["ownunitno"] = m.group(2)      # e.g. 200
        else:
            out["ownunit"] = unit

    out["owncity"]  = city or None
    out["ownstate"] = state or None
    digits = re.sub(r"[^0-9]", "", zipraw)
    if len(digits) >= 5:
        out["ownzipcode"] = digits[:5]
        out["ownzip4"] = digits[5:9] if len(digits) >= 9 else None
    elif zipraw:
        out["ownzipcode"] = zipraw

    raw = " | ".join(x for x in (street, other, unit, city, state, zipraw) if x)
    return out, raw


def parse_mailing_address(lines):
    """Split mailing-address lines into the 11 OWN* components.
    Every value is a string or None.  Nothing is dropped: if the street
    line can't be broken down, the whole line goes into ownstr."""
    out = {k: None for k in MAIL_COL_MAP}
    lines = [re.sub(r"\s+", " ", l.replace("\xa0", " ")).strip().upper()
             for l in (lines or [])]
    lines = [l for l in lines if l]
    if not lines:
        return out

    # City / State / Zip — search from the bottom up
    for idx in range(len(lines) - 1, -1, -1):
        m = CSZ_RE.match(lines[idx])
        if m:
            out["owncity"]    = m.group("city").strip(" ,")
            out["ownstate"]   = m.group("state")
            out["ownzipcode"] = m.group("zip")
            out["ownzip4"]    = m.group("zip4")
            street_lines = lines[:idx]
            if m.group("pre"):
                street_lines.append(m.group("pre").strip())
            break
    else:
        street_lines = lines

    # Pick the street line: PO box or one starting with a digit, else last
    street = None
    for l in reversed(street_lines):
        if BOX_RE.match(l) or re.match(r"^\d", l):
            street = l
            break
    if street is None and street_lines:
        street = street_lines[-1]
    if not street:
        return out

    _parse_street(street, out)
    return out


def mail_cell_value(key, value):
    """Value to write to the sheet.  The existing OWN* columns hold text
    (house numbers and ZIPs included), so everything is written as text."""
    if value in (None, ""):
        return None
    return str(value)


# ─── BROWSER HELPERS ────────────────────────────────────────────────────────

def create_driver():
    options = webdriver.ChromeOptions()
    if HEADLESS:
        options.add_argument("--headless=new")
        options.add_argument(
            "user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
        )
    return webdriver.Chrome(options=options)


def check_for_overlimit(driver, context=""):
    if "overlimit" in driver.current_url.lower():
        raise OverLimitError(
            f"Hit OverLimit.aspx{(' during ' + context) if context else ''} "
            f"(url={driver.current_url})"
        )


def accept_disclaimer_if_present(driver):
    try:
        agree = WebDriverWait(driver, 4).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//*[self::a or self::button or self::input]"
                "[translate(normalize-space(text()),"
                "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz') = 'agree' "
                "or translate(normalize-space(@value),"
                "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz') = 'agree']"
            ))
        )
        agree.click()
    except TimeoutException:
        pass


def no_records_found(driver):
    """Return True if the search results page says no records were found."""
    return "did not find any records" in driver.page_source


def search_upc(driver, upc):
    driver.get(SEARCH_URL)
    accept_disclaimer_if_present(driver)
    wait = WebDriverWait(driver, 15)
    # Explicitly set Tax Year = 2026 BEFORE searching so the record
    # returned is the 2026 assessment, not a prior year's.
    try:
        yr_sel = wait.until(EC.presence_of_element_located((By.ID, "inpTaxyr")))
        Select(yr_sel).select_by_value("2026")
    except Exception:
        pass   # if dropdown not present, site default (2026) applies
    field = wait.until(EC.presence_of_element_located((By.ID, "inpSuf")))
    field.clear()
    field.send_keys(upc)
    driver.find_element(By.ID, "btSearch").click()
    # If the 2026 record doesn't exist, raise clearly rather than
    # proceeding to a page with no profile link.
    if no_records_found(driver):
        raise RuntimeError(
            f"No 2026 record found for UPC {upc} on the Assessor website. "
            f"This UPC may not yet have a 2026 assessment."
        )


def search_address(driver, house_num, street_name):
    driver.get(SEARCH_URL)
    accept_disclaimer_if_present(driver)
    wait = WebDriverWait(driver, 15)
    try:
        yr_sel = wait.until(EC.presence_of_element_located((By.ID, "inpTaxyr")))
        Select(yr_sel).select_by_value("2026")
    except Exception:
        pass
    no_field = wait.until(EC.presence_of_element_located((By.ID, "inpNo")))
    no_field.clear()
    no_field.send_keys(str(house_num))
    street_field = driver.find_element(By.ID, "inpStreet")
    street_field.clear()
    street_field.send_keys(str(street_name))
    driver.find_element(By.ID, "btSearch").click()


def do_search_with_backoff(driver, search_fn, label):
    """Run a search; restart Chrome and retry on OverLimit."""
    for attempt in range(MAX_RATE_LIMIT_RETRIES + 1):
        search_fn()
        check_for_overlimit(driver, f"search for {label}")
        return   # success
    # Should not reach here normally — OverLimitError propagates up


def ensure_on_profile_page(driver):
    """After searching, confirm we are on the property profile page.
    Clicks into the first result row if we land on a list page instead."""
    check_for_overlimit(driver, "search result")
    wait = WebDriverWait(driver, 15)
    try:
        wait.until(EC.presence_of_element_located((By.ID, "Real Property Attributes")))
        return
    except TimeoutException:
        pass
    # Fallback: results list page
    try:
        check_for_overlimit(driver, "search results list")
        row_el = driver.find_element(By.CSS_SELECTOR, "tr.SearchResults")
        row_el.click()
        time.sleep(0.8)
        check_for_overlimit(driver, "SearchResults row click")
        wait.until(EC.presence_of_element_located((By.ID, "Real Property Attributes")))
    except OverLimitError:
        raise
    except (NoSuchElementException, TimeoutException):
        try:
            with open("debug_last_failure.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            driver.save_screenshot("debug_last_failure.png")
        except Exception:
            pass
        raise RuntimeError(
            f"Could not reach property profile page "
            f"(url={driver.current_url}, title={driver.title!r})"
        )


def ensure_tax_year_2026(driver):
    """After landing on the property profile page, verify the Record
    Navigator Tax Year dropdown is set to 2026 and select it if not.
    This ensures owner name, class, and all attributes reflect 2026
    specifically, not a prior year carried over from a previous session."""
    wait = WebDriverWait(driver, 10)
    try:
        dropdown_el = wait.until(
            EC.presence_of_element_located((By.ID, "DTLNavigator_ddTaxYear"))
        )
        sel = Select(dropdown_el)
        if sel.first_selected_option.get_attribute("value") == "2026":
            return   # already on 2026 — nothing to do
        sel.select_by_value("2026")
        # Wait for the page to reload with 2026 data
        wait.until(EC.staleness_of(dropdown_el))
        wait.until(EC.presence_of_element_located((By.ID, "DTLNavigator_ddTaxYear")))
        time.sleep(0.5)   # brief settle after reload
    except TimeoutException:
        pass   # some profile pages don't have a year selector — continue


def extract_parid(driver):
    """Read the PARID from the profile page header (used for UPC correction)."""
    soup = BeautifulSoup(driver.page_source, "html.parser")
    for td in soup.find_all("td", class_="DataletHeaderTop"):
        text = td.get_text(strip=True)
        if text.upper().startswith("PARID"):
            parts = text.split(":", 1)
            if len(parts) == 2:
                return parts[1].strip()
    return None


def extract_all_fields(driver):
    """Extract all five target fields from the property profile page.

    Strategy:
      1. Scrape every label → value pair from every 2-column table row on
         the page. This catches OWNER wherever the "Current Owner" section
         renders, and picks up LUC_MSG however the site labels it.
      2. Additionally parse the "Real Property Attributes" table explicitly
         for SQ FT, YR BUILT, and STYLE (and as a second source for LUC_MSG).

    Returns a dict with keys: owner, luc_msg, sqft, year_built, style, the 11
    mailing-address keys in MAIL_COL_MAP, and mailing_raw (original text).
    Any field not found is None.
    """
    soup = BeautifulSoup(driver.page_source, "html.parser")

    # --- Pass 1: harvest ALL label→value pairs across the whole page ---
    all_pairs = {}   # label (stripped) → value (stripped)
    for table in soup.find_all("table"):
        for row in table.find_all("tr"):
            cells = row.find_all("td")
            if len(cells) == 2:
                lbl = cells[0].get_text(strip=True)
                val = cells[1].get_text(strip=True)
                if lbl and lbl not in all_pairs:   # first occurrence wins
                    all_pairs[lbl] = val

    # --- Pass 2: parse Real Property Attributes table specifically ---
    attrs = {}
    rpa_table = soup.find("table", id="Real Property Attributes")
    if rpa_table:
        for row in rpa_table.find_all("tr"):
            cells = row.find_all("td")
            if len(cells) == 2:
                attrs[cells[0].get_text(strip=True)] = cells[1].get_text(strip=True)

    # --- OWNER ---
    # The "Current Owner" section labels the owner row simply "Owner"
    owner = None
    for lbl in ("Owner", "Owner Name", "OWNER"):
        if lbl in all_pairs and all_pairs[lbl]:
            owner = clean_str(all_pairs[lbl])
            break

    # --- LUC_MSG ---
    # On the Bernalillo County profile page the land-use description appears
    # under several possible labels depending on the record type.
    luc_msg = None
    for lbl in ("Use Code", "Land Use", "Code", "Use", "Class Description",
                 "Land Use Code", "Primary Land Use", "LUC", "LUC_MSG"):
        val = attrs.get(lbl) or all_pairs.get(lbl)
        if val:
            luc_msg = clean_str(val)
            break
    # Last-resort: "Class" on the profile header (may be PROPCLASS desc rather
    # than LUC_MSG, but is still useful if nothing better is found)
    if not luc_msg:
        luc_msg = clean_str(all_pairs.get("Class"))

    # --- SQ FT, YR BUILT, STYLE ---
    sqft       = clean_int(attrs.get("Primary Building SQ FT"))
    year_built = clean_int(attrs.get("Year Built"))
    style      = clean_str(attrs.get("Style"))

    # --- OWNER MAILING ADDRESS ---
    global _mailing_debug_written
    mail_parsed, mail_raw = extract_owner_mailing(soup)
    if not mail_raw and not _mailing_debug_written:
        # Label not found — save the page once so the label can be checked
        try:
            with open("debug_no_mailing_address.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            with open("debug_no_mailing_address.txt", "w", encoding="utf-8") as f:
                f.write("\n".join(page_text_lines(soup)))
            print("  (No mailing address found — saved debug_no_mailing_address"
                  ".html and .txt; send these for diagnosis)")
        except Exception:
            pass
        _mailing_debug_written = True

    result = {
        "owner":       owner,
        "luc_msg":     luc_msg,
        "sqft":        sqft,
        "year_built":  year_built,
        "style":       style,
        "mailing_raw": mail_raw,
    }
    result.update(mail_parsed)
    return result


# ─── PARCEL PROCESSING ──────────────────────────────────────────────────────

def process_parcel(driver, upc, situsnum, situsstr):
    """Navigate to the property profile for upc, extract all target fields.

    Returns (fields_dict, corrected_upc). corrected_upc is None unless an
    address fallback was used and produced a different UPC.
    """
    corrected_upc = None
    try:
        do_search_with_backoff(driver, lambda: search_upc(driver, upc), f"UPC {upc}")
        ensure_on_profile_page(driver)
        ensure_tax_year_2026(driver)
    except OverLimitError:
        raise
    except RuntimeError as e:
        # UPC didn't lead to a valid record — try address fallback
        if not situsnum or not situsstr:
            raise RuntimeError(
                f"UPC search failed and no address available for fallback: {e}"
            )
        print(f"  UPC {upc} didn't resolve — trying address search "
              f"({situsnum} {situsstr})...")
        do_search_with_backoff(
            driver,
            lambda: search_address(driver, situsnum, situsstr),
            f"address {situsnum} {situsstr}"
        )
        ensure_on_profile_page(driver)
        ensure_tax_year_2026(driver)
        corrected_upc = extract_parid(driver)
        if not corrected_upc:
            raise RuntimeError(
                "Address fallback reached a profile page but could not read PARID."
            )
        print(f"  Address search found correct UPC: {corrected_upc} (was {upc}).")

    fields = extract_all_fields(driver)
    return fields, corrected_upc


# ─── LOGGING ────────────────────────────────────────────────────────────────

def log_error(row, upc, message):
    exists = os.path.exists(ERROR_LOG_PATH)
    with open(ERROR_LOG_PATH, "a", newline="") as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(["timestamp", "row", "upc", "error"])
        w.writerow([datetime.now().isoformat(timespec="seconds"), row, upc, message])


CHECKPOINT_FIELDS = (["row", "upc", "owner", "luc_msg", "sqft", "year_built", "style"]
                     + list(MAIL_COL_MAP) + ["mailing_raw"])


def append_checkpoint(row, upc, fields):
    exists = os.path.exists(CHECKPOINT_CSV)
    with open(CHECKPOINT_CSV, "a", newline="") as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(CHECKPOINT_FIELDS)
        w.writerow([
            row, upc,
            fields.get("owner") or "",
            fields.get("luc_msg") or "",
            fields.get("sqft") if fields.get("sqft") is not None else "",
            fields.get("year_built") if fields.get("year_built") is not None else "",
            fields.get("style") or "",
        ] + [fields.get(k) or "" for k in MAIL_COL_MAP]
          + [fields.get("mailing_raw") or ""])


def replay_checkpoint(ws):
    """On startup, re-apply anything in the checkpoint that might not have
    made it into the last spreadsheet save.  Returns count of rows applied."""
    if not os.path.exists(CHECKPOINT_CSV):
        return 0
    applied = 0
    with open(CHECKPOINT_CSV, "r", newline="") as f:
        for entry in csv.DictReader(f):
            row = int(entry["row"])
            if str(ws.cell(row=row, column=UPC_COL).value or "").strip() != entry["upc"].strip():
                continue   # UPC mismatch — skip to avoid writing to wrong row
            write_cell(ws, row, OWNER_COL,     entry["owner"]      or None)
            write_cell(ws, row, LUC_MSG_COL,   entry["luc_msg"]    or None)
            write_cell(ws, row, SQFT_COL,
                       int(entry["sqft"]) if entry["sqft"] else None)
            write_cell(ws, row, YEARBUILT_COL,
                       int(entry["year_built"]) if entry["year_built"] else None)
            write_cell(ws, row, STYLE_COL,     entry["style"]      or None)
            if (entry.get("mailing_raw") or "").strip():
                for key, col in MAIL_COL_MAP.items():
                    write_cell(ws, row, col, mail_cell_value(key, entry.get(key)))
            applied += 1
    return applied


# ─── MAIN ───────────────────────────────────────────────────────────────────

def main():
    print("\nBernCo 2026 Transition — Property Profile Updater")
    print(f"  File: {XLSX_PATH}\n")

    wb = load_workbook(XLSX_PATH)
    ws = wb[SHEET_NAME]

    recovered = replay_checkpoint(ws)
    if recovered:
        print(f"Recovered {recovered} row(s) from checkpoint into spreadsheet.")
        atomic_save(wb, XLSX_PATH)

    driver = create_driver()
    end_row = END_ROW if END_ROW is not None else ws.max_row
    processed = 0
    last_completed_row = None
    consecutive_failures = 0

    try:
        for row in range(START_ROW, end_row + 1):
            if TEST_LIMIT is not None and processed >= TEST_LIMIT:
                print(f"TEST_LIMIT of {TEST_LIMIT} reached — stopping.")
                break

            upc = ws.cell(row=row, column=UPC_COL).value
            if not upc:
                continue

            # Skip if STYLE column already has a value (row completed by
            # a prior run of this script)
            existing_style = ws.cell(row=row, column=STYLE_COL).value
            if existing_style not in (None, ""):
                if not BACKFILL_MAILING:
                    continue
                # Backfill mode: only revisit rows with no mailing data yet
                if any(ws.cell(row=row, column=c).value not in (None, "")
                       for c in MAIL_COL_MAP.values()):
                    continue

            situsnum = ws.cell(row=row, column=SITUSNUM_COL).value
            situsstr = ws.cell(row=row, column=SITUSSTR_COL).value

            fields = None
            corrected_upc = None
            crash_retries = 0
            overlimit_retries = 0

            while True:
                try:
                    fields, corrected_upc = process_parcel(
                        driver, str(upc), situsnum, situsstr
                    )
                    break

                except OverLimitError as e:
                    overlimit_retries += 1
                    if overlimit_retries > MAX_RATE_LIMIT_RETRIES:
                        print(f"  Rate limit persisted after {MAX_RATE_LIMIT_RETRIES} "
                              f"Chrome restarts — skipping UPC {upc}.")
                        log_error(row, upc, f"OverLimit gave up: {e}")
                        fields = None
                        break
                    print(f"  Rate limit hit — restarting Chrome "
                          f"(attempt {overlimit_retries}/{MAX_RATE_LIMIT_RETRIES})...")
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    driver = create_driver()
                    wait_secs = RATE_LIMIT_WAIT_SECONDS * overlimit_retries
                    print(f"  Waiting {wait_secs}s before retrying UPC {upc}...")
                    time.sleep(wait_secs)
                    continue

                except WebDriverException as e:
                    crash_retries += 1
                    print(f"  Browser crashed on row {row} UPC {upc} — restarting Chrome...")
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    driver = create_driver()
                    if crash_retries > 2:
                        log_error(row, upc, f"WebDriverException after {crash_retries} restarts: {e}")
                        fields = None
                        break
                    continue

                except Exception as e:
                    print(f"  ERROR row {row} UPC {upc}: {e}")
                    log_error(row, upc, str(e))
                    fields = None
                    break

            if fields is None:
                consecutive_failures += 1
                if consecutive_failures >= MAX_CONSECUTIVE_FAILURES:
                    print(f"\n{consecutive_failures} consecutive failures — stopping run.")
                    print(f"Last completed row: {last_completed_row}")
                    atomic_save(wb, XLSX_PATH)
                    break
                continue
            consecutive_failures = 0

            # Correct UPC in col A if the address fallback found a better one
            if corrected_upc and corrected_upc != str(upc):
                write_cell(ws, row, UPC_COL, corrected_upc)
                print(f"  Row {row}: UPC corrected from {upc} → {corrected_upc}")
                upc = corrected_upc

            # Write all sixteen fields; use sentinel for text fields with no data
            write_cell(ws, row, OWNER_COL,    fields["owner"]      or None)
            write_cell(ws, row, LUC_MSG_COL,  fields["luc_msg"]    or None)
            write_cell(ws, row, SQFT_COL,
                       fields["sqft"]      if fields["sqft"]      is not None else 0)
            write_cell(ws, row, YEARBUILT_COL,
                       fields["year_built"] if fields["year_built"] is not None else 0)
            # STYLE is the "done" sentinel — always write something
            write_cell(ws, row, STYLE_COL,
                       fields["style"] if fields["style"] else STYLE_SENTINEL)

            if fields.get("mailing_raw"):
                for key, col in MAIL_COL_MAP.items():
                    write_cell(ws, row, col, mail_cell_value(key, fields.get(key)))
            else:
                # Do NOT blank out the existing address — leave J–T as they were
                print("  WARNING: no mailing address found on page — "
                      "columns J–T left unchanged.")
                log_error(row, upc, "Mailing address not found — J-T unchanged")

            append_checkpoint(row, str(upc), fields)

            processed += 1
            last_completed_row = row
            print(
                f"Row {row}/{end_row}: UPC {upc} done "
                f"[Owner={fields['owner'] or '—'}  "
                f"LUC={fields['luc_msg'] or '—'}  "
                f"SqFt={fields['sqft']}  "
                f"YrBuilt={fields['year_built']}  "
                f"Style={fields['style'] or '—'}  "
                f"Mail={fields['mailing_raw'] or '—'}] "
                f"({processed} this run)"
            )

            if processed % SAVE_XLSX_EVERY == 0:
                atomic_save(wb, XLSX_PATH)
                print(f"  (Spreadsheet saved through row {row})")

            if RESTART_DRIVER_EVERY and processed % RESTART_DRIVER_EVERY == 0:
                print(f"  (Proactively restarting Chrome after {processed} parcels...)")
                try:
                    driver.quit()
                except Exception:
                    pass
                driver = create_driver()

            time.sleep(PAUSE_BETWEEN_PARCELS)

    except KeyboardInterrupt:
        print("\nStopped by user (Ctrl+C).")
        print(f"Last fully completed row: {last_completed_row}")
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        atomic_save(wb, XLSX_PATH)
        print("Spreadsheet saved. Safe to close.")

    print(f"\nFinished. {processed} rows updated this run.")


if __name__ == "__main__":
    main()
