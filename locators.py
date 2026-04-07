"""
selectors.py
============
Selectors for the Fonet HBYS web application (ExtJS / Sencha framework).

IMPORTANT — NEVER use ext-comp-XXXX or ext-gen-XXXX ids.
ExtJS regenerates all numeric ids on every page load.  Always use:
  - name= attributes on <input> elements
  - stable CSS class names (e.g. goHistory16, logingirisbuton)
  - text-based XPath (normalize-space()='...')
  - structural XPath relative to stable class anchors

RIGHT-CLICK IS BLOCKED by ExtJS.  Use the DevTools Console instead:
  copy(element.outerHTML)        — puts HTML in clipboard
  copy(element.innerText)        — puts visible text in clipboard

=====================================================================
CONFIRMED SELECTORS (do not change these)
=====================================================================
  Login page:   no top-level iframe
  Desktop body: body.x-desktop-ct  (post-login landmark)
  Username:     input[name='kullaniciAdi']
  Password:     input[name='kullaniciSifre']
  Login button: a.logingirisbuton
  Hasta Geçmişi button: goHistory16 icon class (confirmed)
  Poliklinik iframe: None (module renders in main document)
=====================================================================
"""

# ---------------------------------------------------------------------------
# IFRAMES
# ---------------------------------------------------------------------------

# Confirmed: no top-level iframe. Login and desktop render in main document.
MAIN_APP_IFRAME = None

# Confirmed: the only iframe in the Poliklinik window is id="data_export_iframe"
# (a hidden export helper). Module content renders directly in the document.
POLIKLINIK_IFRAME = None

# ---------------------------------------------------------------------------
# LOGIN PAGE
# ---------------------------------------------------------------------------

# Confirmed from page HTML.
LOGIN_USERNAME_INPUT = "input[name='kullaniciAdi']"
LOGIN_PASSWORD_INPUT = "input[name='kullaniciSifre']"

# Confirmed: ExtJS combobox on the login form.
# Click the trigger button (dropdown arrow) to open the list, then
# click the first item.  The trigger is a sibling element inside the
# same trigger-wrap as the input.
LOGIN_ORG_INPUT      = "input[name='kullaniciOrganizasyon']"
LOGIN_ORG_TRIGGER    = "//input[@name='kullaniciOrganizasyon']/..//*[contains(@class,'x-form-trigger')]"
LOGIN_ORG_FIRST_ITEM = ".x-boundlist-item"

# Confirmed: ExtJS <a> button, not a <button> element.
LOGIN_SUBMIT_BUTTON = "a.logingirisbuton"

# Confirmed from DevTools breadcrumb: body carries class "x-desktop-ct" only
# after login — not present on the login page.
POST_LOGIN_LANDMARK = "body.x-desktop-ct"

# ---------------------------------------------------------------------------
# KRITIK STOK POPUP — dismissal
# ---------------------------------------------------------------------------

# A "Kritik Stok" dialog appears after every login asking to open the stock
# screen.  We click "Hayır" to dismiss it without opening anything.
# Uses standard ExtJS <a class="x-btn"> pattern confirmed throughout this app.
KRITIK_STOK_HAYIR_BUTTON = "//a[.//span[normalize-space()='Hayır']]"

# ---------------------------------------------------------------------------
# DESKTOP NAVIGATION — opening the Poliklinik module
# ---------------------------------------------------------------------------

# Confirmed: after opening the module, click the Poliklinik tab in the tab bar
# to load the patient search panel (module renders empty without this click).
POLIKLINIK_TAB = "//a[contains(@class,'x-btn-toolbar') and not(contains(@class,'x-pressed')) and .//span[normalize-space()='Poliklinik']]"

# Confirmed from Query P: tile text is in <span class="ux-desktop-shortcut-text">.
# The click handler is on the parent element, so we navigate up with /..
POLIKLINIK_MENU_ITEM = "//span[contains(@class,'ux-desktop-shortcut-text') and normalize-space()='Poliklinik']/.."

# TODO: A stable element that appears only after the Poliklinik module has
# fully loaded.  The "Sorgula" button or "Kriter Paneli" heading is a good
# choice.  Run this in Console while the module is open:
#   copy(document.querySelector('a span.x-btn-inner')?.closest('a')?.outerHTML)
# then pick any always-visible button label unique to the module.
# Confirmed: "Sorgula" button is always present in the Kriter Paneli once loaded.
POLIKLINIK_LOADED_LANDMARK = "//a[.//span[normalize-space()='Sorgula']]"

# ---------------------------------------------------------------------------
# PATIENT SEARCH — Kriter Paneli (bottom-left filter panel)
# ---------------------------------------------------------------------------

# Confirmed from full input scan: name="tcKimlikNo".
# There are two fields with this name — one in the Kriter Paneli search form,
# one in the patient detail display area.  The Kriter Paneli one renders first
# in the DOM, so [1] reliably picks the search input.
TC_SEARCH_INPUT = "(//input[@name='tcKimlikNo'])[1]"

# Confirmed: button label is "Sorgula".
TC_SEARCH_BUTTON = "//a[.//span[normalize-space()='Sorgula']]"

# "Temizle" button resets the Kriter Paneli search fields.
CLEAR_SEARCH_BUTTON = "//a[.//span[normalize-space()='Temizle']]"   # TODO: verify

# Confirmed: the grid shows this div when there are no result rows.
# Also present before any search (initial state message).
# search_patient() does NOT rely on this — it simply checks for zero rows.
PATIENT_NOT_FOUND_INDICATOR = "div.x-grid-empty"

# ---------------------------------------------------------------------------
# PATIENT ROW — left grid list
# ---------------------------------------------------------------------------

# Confirmed from console: container class includes x-grid-view.
PATIENT_LIST_CONTAINER = "//div[contains(@class,'x-grid-view')]"

# Confirmed from console: row class includes x-grid-data-row.
PATIENT_ROW_XPATH = ".//tr[contains(@class,'x-grid-data-row')]"

# Confirmed from Query O: TD 9 holds visit date+time as "dd.mm.yyyy HH:MM" (no seconds).
# TD 8 is identical but includes seconds — TD 9 format matches the parser directly.
PATIENT_ROW_DATE_XPATH = ".//td[9]"

# ---------------------------------------------------------------------------
# HASTA GEÇMİŞİ BUTTON
# ---------------------------------------------------------------------------

# Confirmed from console result 1:
# <a ...><span ...><span class="x-btn-icon-el goHistory16 " ...></span></span></a>
# The "goHistory16" icon class is stable across sessions.
HASTA_GECMISI_BUTTON = "//span[contains(@class,'goHistory16')]/ancestor::a[1]"

# ---------------------------------------------------------------------------
# HASTA GEÇMİŞİ POPUP
# ---------------------------------------------------------------------------

# The popup window.  ExtJS windows have class "x-window".  We narrow to the
# one whose title contains "Hasta Geçmi" (handles "Hasta Geçmiş" with any
# trailing text).
# TODO: Verify by running in Console after the popup opens:
#   copy(Array.from(document.querySelectorAll('.x-window')).find(w =>
#     w.innerText.includes('Hasta Geçmi') && w.style.display !== 'none')?.className)
GECMIS_POPUP = "//div[contains(@class,'x-window') and not(contains(@style,'display: none')) and .//span[contains(text(),'Hasta Geçmi')]]"

# "Tüm Hastaneler" toggle — label text is "Tüm Hastaneler:" (with colon), inside a
# toolbar docked at the bottom of the popup grid. Clicked via JavaScript using the
# label's `for` attribute to find the associated input — XPath scoping was unreliable.
# This selector is kept for reference only; the actual click uses JS in scraper.py.
TUM_HASTANELER_BUTTON = ".//label[contains(normalize-space(),'Tüm Hastaneler')]/following::input[@type='button' and contains(@class,'x-form-checkbox')][1]"

# Individual row in the popup history grid.
# Relative to the GECMIS_POPUP element.
GECMIS_ROW_XPATH = ".//tr[contains(@class,'x-grid-row')]"

# Column indices (1-based <td> positions) — confirmed from row HTML inspection.
# Rendered TD layout (hidden columns produce no TD):
#   1: checkbox  2: row#  3: status icon
#   4: Organizasyon  5: İşlem No  6: Arşiv No  7: Vaka Türü
#   8: Sevk Tarihi   9: Sevk Tipi  10: Muayene Tipi  11: Kabul Şekli
#   12: Birim  (gridcolumn-6399..6402 = Branş Grup Id/Birim Id/Birim Kodu/Birim Grubu are hidden)
# Cell text lives inside <div class="x-grid-cell-inner"> — Selenium .text reads it correctly.
GECMIS_ROW_TARIH_XPATH    = ".//td[8]"   # Sevk Tarihi
GECMIS_ROW_SEVK_TIPI_XPATH = ".//td[9]"  # Sevk Tipi
GECMIS_ROW_BIRIM_XPATH    = ".//td[12]"  # Birim (ward)

# Close (X) button for the Hasta Geçmişi popup.
# ExtJS close tools have class "x-tool-close".
GECMIS_POPUP_CLOSE = "//div[contains(@class,'x-window') and not(contains(@style,'display: none')) and .//span[contains(text(),'Hasta Geçmi')]]//*[contains(@class,'x-tool-close')]"
