"""
scraper.py
==========
HBYS automation — reads TC kimlik numbers from an Excel file, finds each
patient in the Poliklinik module via "Hasta Geçmişi", checks for a Yatış
(inpatient admission) record in the current calendar month across all
hospitals, and writes the ward name (Birim) + admission date (Sevk Tarihi)
back to the Excel file.

Workflow per patient:
  1. Enter TC in Kimlik No: field → click Sorgula
  2. Click the patient row in the results grid
  3. Click Hasta Geçmişi button
  4. Click Tüm Hastaneler (red square → all-hospital view)
  5. Scan popup rows for current-month row where Sevk Tipi == "Yatış"
  6. Extract Birim + Sevk Tarihi from that row
  7. Close popup → repeat

Usage:
    python scraper.py

Prerequisites:
    1. pip install -r requirements.txt
    2. Create .env:  HBYS_USERNAME=...  HBYS_PASSWORD=...
    3. Place input.xlsx with TC numbers in column A from row 2
    4. Verify/fill TODO selectors in selectors.py
"""

import sys
import time
import logging
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
)
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl

import config
import locators as SEL

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Driver
# ---------------------------------------------------------------------------

def make_driver() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    # Uncomment for headless (no browser window) production runs:
    # options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.implicitly_wait(0)
    return driver

# ---------------------------------------------------------------------------
# Wait helpers
# ---------------------------------------------------------------------------

def wait_visible(driver, by, selector, timeout=None):
    t = timeout or config.WAIT_TIMEOUT
    return WebDriverWait(driver, t).until(
        EC.visibility_of_element_located((by, selector))
    )


def wait_clickable(driver, by, selector, timeout=None):
    t = timeout or config.WAIT_TIMEOUT
    return WebDriverWait(driver, t).until(
        EC.element_to_be_clickable((by, selector))
    )


def safe_find(context, by, selector):
    try:
        return context.find_element(by, selector)
    except NoSuchElementException:
        return None


def switch_to_iframe(driver, css_selector):
    if css_selector is None:
        return
    frame = wait_visible(driver, By.CSS_SELECTOR, css_selector)
    driver.switch_to.frame(frame)

# ---------------------------------------------------------------------------
# Date parsing
# ---------------------------------------------------------------------------

def is_current_month(date_string: str) -> bool:
    """Return True if date_string falls in the current calendar month."""
    now = datetime.now()
    for fmt in ("%d.%m.%Y %H:%M", "%d.%m.%Y", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(date_string.strip(), fmt)
            return dt.year == now.year and dt.month == now.month
        except ValueError:
            continue
    log.warning("Could not parse date: '%s'", date_string)
    return False

# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------

def load_workbook(filepath: str):
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    tc_col    = openpyxl.utils.column_index_from_string(config.TC_COLUMN)
    tarih_col = openpyxl.utils.column_index_from_string(config.TARIH_COLUMN)
    min_col   = min(tc_col, tarih_col)
    max_col   = max(tc_col, tarih_col)
    records = []
    for row in ws.iter_rows(min_row=config.DATA_START_ROW, min_col=min_col, max_col=max_col):
        tc_cell    = row[tc_col - min_col]
        tarih_cell = row[tarih_col - min_col]
        if tc_cell.value is not None:
            tc    = str(tc_cell.value).strip().zfill(11)
            tarih = tarih_cell.value  # datetime or string from Excel
            if isinstance(tarih, str):
                for fmt in ("%Y-%m-%d %H:%M:%S", "%d.%m.%Y %H:%M", "%Y-%m-%d"):
                    try:
                        tarih = datetime.strptime(tarih.strip(), fmt)
                        break
                    except ValueError:
                        continue
            records.append((tc_cell.row, tc, tarih))
    log.info("Loaded %d records from %s", len(records), filepath)
    return records, wb, ws


def ensure_output_columns(ws) -> dict:
    header_row = 1
    headers = {
        ws.cell(row=header_row, column=c).value: c
        for c in range(1, ws.max_column + 1)
    }
    col_map = {}
    for col_name in (config.COL_YATIS_VAR, config.COL_BIRIM, config.COL_SEVK_TARIHI):
        if col_name not in headers:
            next_col = ws.max_column + 1
            ws.cell(row=header_row, column=next_col, value=col_name)
            headers[col_name] = next_col
        col_map[col_name] = headers[col_name]
    return col_map


def write_result(ws, row_idx: int, col_map: dict,
                 yatis_var: str, birim: str, tarih: str):
    ws.cell(row=row_idx, column=col_map[config.COL_YATIS_VAR],  value=yatis_var)
    ws.cell(row=row_idx, column=col_map[config.COL_BIRIM],      value=birim)
    ws.cell(row=row_idx, column=col_map[config.COL_SEVK_TARIHI], value=tarih)

# ---------------------------------------------------------------------------
# Login
# ---------------------------------------------------------------------------

def dismiss_kritik_stok_popup(driver):
    """Dismiss the 'Kritik Stok' dialog that appears after every login."""
    try:
        btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, SEL.KRITIK_STOK_HAYIR_BUTTON))
        )
        btn.click()
        log.info("Dismissed Kritik Stok popup.")
    except TimeoutException:
        pass  # Popup didn't appear — continue normally


def login(driver):
    log.info("Navigating to %s", config.BASE_URL)
    driver.get(config.BASE_URL)
    switch_to_iframe(driver, SEL.MAIN_APP_IFRAME)

    wait_clickable(driver, By.CSS_SELECTOR, SEL.LOGIN_USERNAME_INPUT).send_keys(config.USERNAME)

    # Click password field — this blurs username and triggers org store load
    password_el = wait_clickable(driver, By.CSS_SELECTOR, SEL.LOGIN_PASSWORD_INPUT)
    password_el.click()

    # Wait for org store to filter to exactly 1 record (user's hospital only)
    # Store starts with 7 items on page load, drops to 1 after username blur
    WebDriverWait(driver, config.WAIT_TIMEOUT).until(
        lambda d: d.execute_script(
            "try { return Ext.ComponentQuery.query('combobox[name=\"kullaniciOrganizasyon\"]')[0]"
            "?.getStore()?.getCount() === 1; } catch(e) { return false; }"
        )
    )

    # Pass the full record to setValue — more reliable than passing value field alone
    org_name = driver.execute_script("""
        var c = Ext.ComponentQuery.query('combobox[name="kullaniciOrganizasyon"]')[0];
        var rec = c.getStore().getAt(0);
        c.setValue(rec);
        c.fireEvent('select', c, [rec]);
        return rec.get(c.displayField);
    """)
    log.info("Selected org: '%s'", org_name)
    time.sleep(0.5)

    # Re-find password field in case element reference went stale
    password_el = wait_clickable(driver, By.CSS_SELECTOR, SEL.LOGIN_PASSWORD_INPUT)
    password_el.clear()
    password_el.send_keys(config.PASSWORD)
    time.sleep(0.3)

    submit = wait_clickable(driver, By.CSS_SELECTOR, SEL.LOGIN_SUBMIT_BUTTON)
    driver.execute_script("arguments[0].click();", submit)

    # Wait for Poliklinik tile on desktop — confirms login completed
    WebDriverWait(driver, config.WAIT_TIMEOUT).until(
        EC.presence_of_element_located((By.XPATH, SEL.POLIKLINIK_MENU_ITEM))
    )
    log.info("Login successful.")
    dismiss_kritik_stok_popup(driver)

# ---------------------------------------------------------------------------
# Open Poliklinik module
# ---------------------------------------------------------------------------

def open_poliklinik_module(driver):
    log.info("Opening Poliklinik module...")
    driver.switch_to.default_content()
    switch_to_iframe(driver, SEL.MAIN_APP_IFRAME)
    tile = wait_clickable(driver, By.XPATH, SEL.POLIKLINIK_MENU_ITEM)
    ActionChains(driver).move_to_element(tile).click().perform()
    time.sleep(3)

    # Log what opened so we can debug
    tab_count = driver.execute_script("return document.querySelectorAll('a.x-tab').length")
    iframe_count = driver.execute_script("return document.querySelectorAll('iframe').length")
    body_classes = driver.execute_script("return document.body.className")
    log.info("After tile click: tabs=%s iframes=%s body=%s", tab_count, iframe_count, body_classes[:80])

    switch_to_iframe(driver, SEL.POLIKLINIK_IFRAME)
    wait_visible(driver, By.XPATH, SEL.POLIKLINIK_LOADED_LANDMARK, timeout=30)
    log.info("Poliklinik module loaded.")

# ---------------------------------------------------------------------------
# Search patient by TC
# ---------------------------------------------------------------------------

def search_patient(driver, tc: str) -> bool:
    """Enter TC and click Sorgula. Returns True if at least one row appears."""
    log.info("Searching TC: %s", tc)

    tc_input = wait_clickable(driver, By.XPATH, SEL.TC_SEARCH_INPUT)
    tc_input.clear()
    tc_input.send_keys(tc)

    wait_clickable(driver, By.XPATH, SEL.TC_SEARCH_BUTTON).click()

    # Wait for the grid to settle: either rows appear or it stays empty.
    # x-grid-empty is always present (initial state + no-results), so we
    # can't use it as a signal — we just wait for rows or timeout.
    try:
        WebDriverWait(driver, config.WAIT_TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH, SEL.PATIENT_ROW_XPATH))
        )
    except TimeoutException:
        log.info("No patient rows found for TC: %s", tc)
        return False

    rows = driver.find_elements(By.XPATH, SEL.PATIENT_ROW_XPATH)
    log.info("Found %d row(s) for TC: %s", len(rows), tc)
    return bool(rows)

# ---------------------------------------------------------------------------
# Select patient row closest to transfer date
# ---------------------------------------------------------------------------

def select_patient_row_by_date(driver, transfer_dt: datetime):
    """
    Click the patient grid row whose visit date is closest to transfer_dt
    and within 24 hours of it.  Falls back to the first row if the date
    column cannot be read or no row falls within 24 hours.
    """
    container = wait_visible(driver, By.XPATH, SEL.PATIENT_LIST_CONTAINER)
    rows = container.find_elements(By.XPATH, SEL.PATIENT_ROW_XPATH)
    if not rows:
        raise RuntimeError("No patient rows to select.")

    best_row  = None
    best_delta = None

    for row in rows:
        date_cell = safe_find(row, By.XPATH, SEL.PATIENT_ROW_DATE_XPATH)
        if not date_cell:
            continue
        date_text = date_cell.text.strip()
        row_dt = None
        for fmt in ("%d.%m.%Y %H:%M", "%d.%m.%Y", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
            try:
                row_dt = datetime.strptime(date_text, fmt)
                break
            except ValueError:
                continue
        if row_dt is None:
            continue
        delta = abs((row_dt - transfer_dt).total_seconds())
        if delta <= 86400 and (best_delta is None or delta < best_delta):
            best_row   = row
            best_delta = delta

    if best_row is None:
        log.warning("No row within 24h of transfer date %s — using first row.", transfer_dt)
        best_row = rows[0]
    else:
        log.info("Selected row with delta %.0f s from transfer date.", best_delta)

    best_row.click()
    time.sleep(config.NAV_SETTLE_WAIT)

# ---------------------------------------------------------------------------
# Open Hasta Geçmişi popup
# ---------------------------------------------------------------------------

def open_hasta_gecmisi(driver):
    """Click the Hasta Geçmişi button and wait for the popup."""
    log.info("Opening Hasta Geçmişi...")
    wait_clickable(driver, By.XPATH, SEL.HASTA_GECMISI_BUTTON).click()
    wait_visible(driver, By.XPATH, SEL.GECMIS_POPUP)
    log.info("Hasta Geçmişi popup opened.")

# ---------------------------------------------------------------------------
# Click Tüm Hastaneler
# ---------------------------------------------------------------------------

def click_tum_hastaneler(driver):
    """
    Click the red 'Tüm Hastaneler' toggle inside the popup to load
    results from all connected hospitals (including the tertiary centre).
    After clicking, the toggle turns green and the grid reloads.
    TUM_HASTANELER_BUTTON is a relative XPath — scoped to the popup element.
    """
    log.info("Clicking Tüm Hastaneler...")
    popup = wait_visible(driver, By.XPATH, SEL.GECMIS_POPUP)
    btn = WebDriverWait(popup, config.WAIT_TIMEOUT).until(
        lambda ctx: ctx.find_element(By.XPATH, SEL.TUM_HASTANELER_BUTTON)
    )
    btn.click()
    # Wait a moment for the grid to reload with all-hospital data
    time.sleep(config.NAV_SETTLE_WAIT)
    log.info("Tüm Hastaneler loaded.")

# ---------------------------------------------------------------------------
# Scan popup rows for Yatış
# ---------------------------------------------------------------------------

def find_yatis_in_popup(driver) -> dict:
    """
    Scan all rows in the Hasta Geçmişi popup.
    Return the first row in the current month where Sevk Tipi == 'Yatış'.
    Returns {"found": bool, "birim": str, "tarih": str}.
    """
    result = {"found": False, "birim": "", "tarih": ""}

    popup = wait_visible(driver, By.XPATH, SEL.GECMIS_POPUP)
    rows = popup.find_elements(By.XPATH, SEL.GECMIS_ROW_XPATH)
    log.info("Popup has %d rows.", len(rows))

    for row in rows:
        try:
            tarih_cell    = safe_find(row, By.XPATH, SEL.GECMIS_ROW_TARIH_XPATH)
            sevk_tipi_cell = safe_find(row, By.XPATH, SEL.GECMIS_ROW_SEVK_TIPI_XPATH)
            birim_cell    = safe_find(row, By.XPATH, SEL.GECMIS_ROW_BIRIM_XPATH)

            if not (tarih_cell and sevk_tipi_cell and birim_cell):
                continue

            tarih_text    = tarih_cell.text.strip()
            sevk_tipi_text = sevk_tipi_cell.text.strip()
            birim_text    = birim_cell.text.strip()

            if sevk_tipi_text != "Yatış":
                continue                          # skip "Diğer Yatış" and others

            if not is_current_month(tarih_text):
                continue

            log.info("  Yatış found: Birim=%s | Tarih=%s", birim_text, tarih_text)
            result = {"found": True, "birim": birim_text, "tarih": tarih_text}
            break                                 # take the first match

        except StaleElementReferenceException:
            log.warning("Stale row element — skipping.")
            continue

    return result

# ---------------------------------------------------------------------------
# Close popup
# ---------------------------------------------------------------------------

def close_gecmis_popup(driver):
    try:
        wait_clickable(driver, By.XPATH, SEL.GECMIS_POPUP_CLOSE, timeout=5).click()
        # Wait for popup to disappear
        WebDriverWait(driver, 5).until(
            EC.invisibility_of_element_located((By.XPATH, SEL.GECMIS_POPUP))
        )
    except TimeoutException:
        log.warning("Could not close Hasta Geçmişi popup — continuing anyway.")

# ---------------------------------------------------------------------------
# Reset search
# ---------------------------------------------------------------------------

def reset_search(driver):
    """Clear Kriter Paneli so the next TC can be entered."""
    try:
        wait_clickable(driver, By.XPATH, SEL.CLEAR_SEARCH_BUTTON, timeout=5).click()
    except TimeoutException:
        tc_input = safe_find(driver, By.XPATH, SEL.TC_SEARCH_INPUT)
        if tc_input:
            tc_input.clear()

# ---------------------------------------------------------------------------
# Per-patient pipeline
# ---------------------------------------------------------------------------

def process_patient(driver, tc: str, transfer_dt: datetime) -> dict:
    not_found = {"yatis_var": "Hasta Bulunamadı", "birim": "", "tarih": ""}
    no_yatis  = {"yatis_var": "Hayır",            "birim": "", "tarih": ""}

    if not search_patient(driver, tc):
        return not_found

    try:
        select_patient_row_by_date(driver, transfer_dt)
    except RuntimeError as exc:
        log.warning("Could not select patient row: %s", exc)
        reset_search(driver)
        return not_found

    open_hasta_gecmisi(driver)
    click_tum_hastaneler(driver)

    info = find_yatis_in_popup(driver)
    close_gecmis_popup(driver)
    reset_search(driver)

    if info["found"]:
        return {"yatis_var": "Evet", "birim": info["birim"], "tarih": info["tarih"]}
    return no_yatis

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if not config.USERNAME or not config.PASSWORD:
        log.error(
            "Credentials not set. Create .env with:\n"
            "  HBYS_USERNAME=...\n  HBYS_PASSWORD=..."
        )
        sys.exit(1)

    records, wb, ws = load_workbook(config.INPUT_FILE)
    col_map = ensure_output_columns(ws)

    driver = make_driver()
    try:
        login(driver)
        open_poliklinik_module(driver)

        for idx, (row_idx, tc, transfer_dt) in enumerate(records, start=1):
            log.info("=== %d / %d | TC: %s ===", idx, len(records), tc)
            try:
                result = process_patient(driver, tc, transfer_dt)
                write_result(ws, row_idx, col_map,
                             yatis_var=result["yatis_var"],
                             birim=result["birim"],
                             tarih=result["tarih"])
                wb.save(config.INPUT_FILE)
                log.info("Written: %s", result)
            except Exception:
                log.exception("Unhandled error for TC %s (row %d)", tc, row_idx)
                write_result(ws, row_idx, col_map,
                             yatis_var="HATA", birim="", tarih="")
                wb.save(config.INPUT_FILE)

    finally:
        driver.quit()
        log.info("Done.")


if __name__ == "__main__":
    main()
