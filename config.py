import os
from dotenv import load_dotenv

load_dotenv()

# --- Credentials (set in .env file) ---
USERNAME = os.getenv("HBYS_USERNAME", "")
PASSWORD = os.getenv("HBYS_PASSWORD", "")

# --- URLs ---
BASE_URL = "http://fonetsbys.usakism.gov.tr/hbys-web/desktop/desktop.html"

# --- Excel ---
INPUT_FILE = "input.xlsx"
TC_COLUMN    = "B"    # Column holding TC kimlik numbers (TCKIMLIKNO = column B)
TARIH_COLUMN = "G"    # Column holding transfer date/time (TARIH = column G)
DATA_START_ROW = 2    # First row of data (row 1 = header)

# Output column headers (appended after the last existing column if not present)
COL_YATIS_VAR   = "Yatış Var mı"   # "Evet" / "Hayır" / "Hasta Bulunamadı" / "HATA"
COL_BIRIM       = "Birim"           # Ward name from Hasta Geçmişi popup
COL_SEVK_TARIHI = "Sevk Tarihi"     # Admission date from Hasta Geçmişi popup

# --- Timing ---
# Maximum seconds to wait for any dynamic element to appear
WAIT_TIMEOUT = 20
# Brief settle wait (seconds) after opening a visit detail — GWT fires async RPCs
# after navigation; only used in one place, all other waits are WebDriverWait
NAV_SETTLE_WAIT = 1
