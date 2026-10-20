#!/usr/bin/env python3
"""
BernCo Tax History Scraper
---------------------------
Reads UPC codes from a spreadsheet (Parcels sheet, column A) and, for each
one, looks up the Bernalillo County Assessor "Notice of Value" data for tax
years 2017-2026, writing Value Land / Value Structures / Taxable (1/3 Full
Value) into the pre-built columns (AH:BK) of that same spreadsheet.

SETUP (one time):
    pip3 install --upgrade selenium beautifulsoup4 openpyxl

    You need Google Chrome installed. Selenium 4.6+ will automatically
    download a matching chromedriver the first time you run this.

USAGE:
    python3 bernco_tax_history_scraper.py

    Edit the CONFIG section below first (file path, test limit, headless).

NOTES:
    - The script saves the spreadsheet after every parcel it completes, so
      it is safe to stop (Ctrl+C) and re-run later -- it will skip any row
      whose 2017 Taxable value is already filled in.
    - Start with TEST_LIMIT = 3 to confirm everything works before turning
      it loose on the full list.
"""

import csv
import os
import time
import tempfile
from datetime import datetime
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from bs4 import BeautifulSoup


class OverLimitError(Exception):
    """Raised whenever OverLimit.aspx is encountered at ANY navigation step,
    not just the initial search.  Caught separately in main() so the parcel
    is retried with a wait rather than logged as a permanent failure."""
    pass


def check_for_overlimit(driver, context=""):
    """Raise OverLimitError if the browser is currently on the rate-limit page."""
    if "overlimit" in driver.current_url.lower():
        raise OverLimitError(
            f"Hit OverLimit.aspx{(' during ' + context) if context else ''} "
            f"(url={driver.current_url})"
        )

# ----------------------------- CONFIG ---------------------------------
XLSX_PATH = "BernCo_Countywide_MultiFamilyResU4U.xlsx"   # path to your spreadsheet
SHEET_NAME = "Parcels"
TEST_LIMIT = None         # set to None to run all rows; start small to verify
HEADLESS = False        # set True once you've confirmed it works, for speed
PAUSE_BETWEEN_PARCELS = 6.0   # seconds — county server rate-limits aggressively; don't go below 5
ERROR_LOG_PATH = "scrape_errors.csv"
RATE_LIMIT_WAIT_SECONDS = 120   # cooldown after OverLimit — short because Chrome restarts give a fresh session
MAX_RATE_LIMIT_RETRIES = 3      # retries before giving up on a parcel
RESTART_DRIVER_EVERY = 40       # restart Chrome every N parcels — server session limit is ~48-50
# ------------------------------------------------------------------------

SEARCH_URL = "https://assessor.bernco.gov/public.access/search/commonsearch.aspx?mode=realprop"
YEARS = [str(y) for y in range(2026, 2016, -1)]  # 2026 down to 2017
UPC_COL = 1  # column A

# Starting column (1-indexed) of the 3-column block (Land, Structures, Taxable) per year
YEAR_COLS = {
    "2026": 39, "2025": 42, "2024": 45, "2023": 48, "2022": 51,
    "2021": 54, "2020": 57, "2019": 60, "2018": 63, "2017": 66,
}


def clean_money(text):
    if not text:
        return None
    text = text.strip().replace("$", "").replace(",", "")
    if text in ("", "&nbsp;"):
        return None
    try:
        return float(text)
    except ValueError:
        return None


def accept_disclaimer_if_present(driver):
    """Best-effort: click an Agree button/link if a disclaimer page appears."""
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
    """Return True if the current page says no records were found."""
    return "did not find any records" in driver.page_source


def set_search_tax_year(driver, year):
    """Change the Tax Year filter on the search/results page and click Search.
    The tax year select (id='inpTaxyr') and Search button remain visible on
    the 'no records' page, so we can change and retry without navigating away."""
    wait = WebDriverWait(driver, 10)
    try:
        el = wait.until(EC.presence_of_element_located((By.ID, "inpTaxyr")))
        Select(el).select_by_value(str(year))
    except Exception as e:
        raise RuntimeError(
            f"Could not locate Tax Year select (id='inpTaxyr') on page "
            f"(url={driver.current_url}): {e}"
        )
    driver.find_element(By.ID, "btSearch").click()


def search_upc(driver, upc):
    driver.get(SEARCH_URL)
    accept_disclaimer_if_present(driver)
    wait = WebDriverWait(driver, 15)
    field = wait.until(EC.presence_of_element_located((By.ID, "inpSuf")))
    field.clear()
    field.send_keys(upc)
    driver.find_element(By.ID, "btSearch").click()


def search_upc_with_backoff(driver, upc):
    """Search for a UPC, handling rate limits and the case where a parcel has
    no 2026 record yet.  Returns the effective tax year string used to find the
    parcel ("2026" or "2025"), so the caller knows whether 2026 data is missing.
    Raises RuntimeError if the parcel cannot be found in either year."""
    for attempt in range(MAX_RATE_LIMIT_RETRIES + 1):
        search_upc(driver, upc)
        if "overlimit" in driver.current_url.lower():
            if attempt < MAX_RATE_LIMIT_RETRIES:
                print(
                    f"  Hit site rate limit (OverLimit.aspx). "
                    f"Waiting {RATE_LIMIT_WAIT_SECONDS}s before retrying UPC {upc} "
                    f"(attempt {attempt + 1}/{MAX_RATE_LIMIT_RETRIES})..."
                )
                time.sleep(RATE_LIMIT_WAIT_SECONDS)
                continue
            else:
                raise OverLimitError(
                    "Site rate limit (OverLimit.aspx) persisted after "
                    f"{MAX_RATE_LIMIT_RETRIES} cooldown(s). Increase PAUSE_BETWEEN_PARCELS."
                )

        if no_records_found(driver):
            # No 2026 record — try 2025 before giving up
            print(f"  No 2026 record for UPC {upc} — trying Tax Year 2025...")
            set_search_tax_year(driver, "2025")
            if no_records_found(driver):
                raise RuntimeError(
                    f"No records found for UPC {upc} in Tax Year 2026 or 2025."
                )
            print(f"  Found UPC {upc} under Tax Year 2025 — 2026 columns will be blank.")
            return "2025"   # caller will leave 2026 cells empty

        return "2026"   # found on first try


def goto_notice_of_value(driver):
    # Guard: if the search itself landed us on OverLimit, raise immediately
    # so the main loop can wait and retry rather than treating it as a
    # permanent failure.
    check_for_overlimit(driver, "search result")

    wait = WebDriverWait(driver, 15)
    try:
        link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Notice of Value")))
        link.click()
        time.sleep(0.8)   # brief wait for URL to settle after click
        check_for_overlimit(driver, "Notice of Value click")
        return
    except OverLimitError:
        raise   # propagate — do not fall through to the debug dump below
    except TimeoutException:
        pass

    # Fallback: maybe we're on a results list (site uses <tr class="SearchResults">
    # onclick rows rather than plain links when a parcel has duplicate records).
    try:
        check_for_overlimit(driver, "search results list")
        row_el = driver.find_element(By.CSS_SELECTOR, "tr.SearchResults")
        row_el.click()
        time.sleep(0.8)
        check_for_overlimit(driver, "SearchResults row click")
        link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Notice of Value")))
        link.click()
        time.sleep(0.8)
        check_for_overlimit(driver, "Notice of Value click (fallback)")
    except OverLimitError:
        raise   # propagate
    except NoSuchElementException:
        # Save debug info so we can see what page we actually landed on
        try:
            with open("debug_last_failure.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            driver.save_screenshot("debug_last_failure.png")
        except Exception:
            pass
        raise RuntimeError(
            f"Could not find 'Notice of Value' link on this page "
            f"(url={driver.current_url}, title={driver.title!r})"
        )


def extract_values(driver):
    soup = BeautifulSoup(driver.page_source, "html.parser")
    table = soup.find("table", id="Values")
    data = {}
    if not table:
        return data
    for row in table.find_all("tr"):
        cells = row.find_all("td")
        if len(cells) == 2:
            label = cells[0].get_text(strip=True)
            value = cells[1].get_text(strip=True)
            data[label] = value
    return data


def select_year(driver, year):
    wait = WebDriverWait(driver, 15)
    dropdown_el = wait.until(EC.presence_of_element_located((By.ID, "DTLNavigator_ddTaxYear")))
    select = Select(dropdown_el)
    if select.first_selected_option.get_attribute("value") == year:
        return
    select.select_by_value(year)
    wait.until(EC.staleness_of(dropdown_el))
    wait.until(EC.presence_of_element_located((By.ID, "DTLNavigator_ddTaxYear")))
    time.sleep(0.4)


def process_parcel(driver, upc):
    results = {}
    effective_year = search_upc_with_backoff(driver, upc)
    goto_notice_of_value(driver)
    for year in YEARS:
        # If this parcel was only found under 2025, leave 2026 cells blank
        # rather than trying to select a year that doesn't exist for it.
        if year == "2026" and effective_year == "2025":
            results[year] = {"land": None, "structures": None, "taxable": None}
            continue
        select_year(driver, year)
        vals = extract_values(driver)
        results[year] = {
            "land": clean_money(vals.get("Land")),
            "structures": clean_money(vals.get("Structures")),
            "taxable": clean_money(vals.get("Taxable (1/3 Full Value)")),
        }
    return results


def atomic_save(wb, path):
    """Save workbook to a temp file first, verify it opens, then atomically
    replace the real file. A crash mid-save leaves the original intact."""
    dir_name = os.path.dirname(os.path.abspath(path))
    fd, tmp_path = tempfile.mkstemp(suffix=".tmp.xlsx", dir=dir_name)
    os.close(fd)
    try:
        wb.save(tmp_path)
        # Quick sanity-check: make sure the file we just wrote actually opens
        from openpyxl import load_workbook as _lw
        _lw(tmp_path, read_only=True).close()
        # Atomic replace — on all major OSes this is a single filesystem op
        os.replace(tmp_path, path)
    except Exception:
        # Clean up the temp file and re-raise so the caller knows
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise


def log_error(row, upc, message):
    file_exists = False
    try:
        with open(ERROR_LOG_PATH, "r"):
            file_exists = True
    except FileNotFoundError:
        pass
    with open(ERROR_LOG_PATH, "a", newline="") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["timestamp", "row", "upc", "error"])
        writer.writerow([datetime.now().isoformat(timespec="seconds"), row, upc, message])


def create_driver():
    options = webdriver.ChromeOptions()
    if HEADLESS:
        options.add_argument("--headless=new")
        options.add_argument(
            "user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
        )
    return webdriver.Chrome(options=options)


def main():
    wb = load_workbook(XLSX_PATH)
    ws = wb[SHEET_NAME]

    driver = create_driver()

    start_row = 3
    end_row = ws.max_row
    processed = 0
    last_completed_row = None
    consecutive_failures = 0
    MAX_CONSECUTIVE_FAILURES = 60  # stop the whole run if this many rows in a row fail

    try:
        for row in range(start_row, end_row + 1):
            if TEST_LIMIT is not None and processed >= TEST_LIMIT:
                break

            upc = ws.cell(row=row, column=UPC_COL).value
            if not upc:
                continue

            taxable_2017_col = YEAR_COLS["2017"] + 2
            if ws.cell(row=row, column=taxable_2017_col).value not in (None, ""):
                continue  # already done

            data = None
            crash_retries = 0
            max_crash_retries = 2
            overlimit_retries = 0
            max_overlimit_retries = 3
            while True:
                try:
                    data = process_parcel(driver, str(upc))
                    break
                except OverLimitError as e:
                    overlimit_retries += 1
                    if overlimit_retries > max_overlimit_retries:
                        print(f"  Rate limit persisted after {max_overlimit_retries} Chrome "
                              f"restarts — skipping UPC {upc} and moving on.")
                        log_error(row, upc, f"OverLimit gave up after restarts: {e}")
                        data = None
                        break
                    # The server enforces a per-SESSION limit (~48-50 requests).
                    # Waiting in the same session never works — restart Chrome to
                    # get a fresh session, then wait briefly for any IP cooldown.
                    print(f"  Rate limit hit (session exhausted). Restarting Chrome for "
                          f"a fresh session (attempt {overlimit_retries}/{max_overlimit_retries})...")
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    driver = create_driver()
                    wait_secs = RATE_LIMIT_WAIT_SECONDS * overlimit_retries
                    print(f"  Waiting {wait_secs}s before retrying UPC {upc}...")
                    time.sleep(wait_secs)
                    continue   # retry the same parcel with the fresh driver
                except WebDriverException as e:
                    crash_retries += 1
                    print(f"ERROR on row {row}, UPC {upc}: browser/driver crashed ({e}). Restarting Chrome...")
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    driver = create_driver()
                    if crash_retries > max_crash_retries:
                        log_error(row, upc, f"WebDriverException, gave up after {max_crash_retries} restarts: {e}")
                        data = None
                        break
                    continue  # retry this same row with the fresh driver
                except Exception as e:
                    print(f"ERROR on row {row}, UPC {upc}: {e}")
                    log_error(row, upc, str(e))
                    data = None
                    break

            if data is None:
                consecutive_failures += 1
                if consecutive_failures >= MAX_CONSECUTIVE_FAILURES:
                    print(
                        f"\n{consecutive_failures} rows in a row failed even after retries -- "
                        "this looks like a sustained site block (likely IP/session-based, "
                        "not something a few more minutes of waiting will fix)."
                    )
                    print(f"Stopping the run here. Last fully completed row: {last_completed_row}")
                    print("Spreadsheet is saved up through that row. Try again later (a few hours, or tomorrow).")
                    break
                continue
            consecutive_failures = 0

            for year, vals in data.items():
                col = YEAR_COLS[year]
                ws.cell(row=row, column=col, value=vals["land"])
                ws.cell(row=row, column=col + 1, value=vals["structures"])
                ws.cell(row=row, column=col + 2, value=vals["taxable"])

            atomic_save(wb, XLSX_PATH)
            processed += 1
            last_completed_row = row
            print(f"Row {row}/{end_row}: UPC {upc} done ({processed} this run)")

            if RESTART_DRIVER_EVERY and processed % RESTART_DRIVER_EVERY == 0:
                print(f"  (Proactively restarting Chrome after {processed} parcels this run...)")
                try:
                    driver.quit()
                except Exception:
                    pass
                driver = create_driver()
            time.sleep(PAUSE_BETWEEN_PARCELS)

    except KeyboardInterrupt:
        print("\nStopped by user (Ctrl+C).")
        print(f"Last fully completed row: {last_completed_row}")
        print("Spreadsheet is already saved up through that row. Safe to close.")
    finally:
        driver.quit()

    print("Finished.")


if __name__ == "__main__":
    main()
