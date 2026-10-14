import os
import io
import re
import csv
import time
import random
import base64
import shutil
import logging
import tempfile
import threading
import multiprocessing as mp
from queue import Empty

import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains


# ================= CONFIG =================

WORKERS = 1               # Single worker avoids IP concurrency rate limiting on IndiGo
WAIT = 30
WAIT_SHORT = 15
PDF_MIN_SIZE = 5000

MAX_RETRIES = 2           # retries per PNR before giving up (increased to handle transient errors)
TASK_TIMEOUT = 120        # hard ceiling (seconds) for one PNR before we force-kill the driver
RESTART_AFTER = 20        # recycle the browser after this many PNRs to stop memory/slowdown creep
RESUME = True             # if True, skip PNRs already present in output.csv / pnr_errors.csv

# ---- Human-like delays between requests ----
# The "sorry, unable to retrieve" error is almost always Akamai/rate-limiting.
# Adding realistic inter-request delays is the most effective fix.
DELAY_BETWEEN_PNRS_MIN = 4    # minimum seconds to wait between PNRs (same worker)
DELAY_BETWEEN_PNRS_MAX = 9    # maximum seconds to wait between PNRs (same worker)
DELAY_AFTER_ERROR_MIN = 15    # wait longer after any error — the site needs cooldown
DELAY_AFTER_ERROR_MAX = 30
TYPING_DELAY_MIN = 0.05       # seconds between each keystroke (human-like typing)
TYPING_DELAY_MAX = 0.18

BLOCK_IMAGES = False      # False: IndiGo/Akamai requires standard assets to load
HEADLESS = False
STAGGER_START = 8

KEEP_WINDOWS_OUT_OF_THE_WAY = False  # Avoid window occlusion JS throttling in Chrome
OFFSCREEN_POSITION = (-32000, -32000)
WINDOW_SIZE = "1366,768"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INVOICE_DIR = os.path.join(BASE_DIR, "invoices")
os.makedirs(INVOICE_DIR, exist_ok=True)

INPUT_FILE = os.path.join(BASE_DIR, "pnr.csv")
OUTPUT_FILE = os.path.join(BASE_DIR, "output.csv")
ERROR_FILE = os.path.join(BASE_DIR, "pnr_errors.csv")

# ---- Corrected XPaths (website was redesigned) ----
# OLD (broken): placeholder="PNR", placeholder="Email", button text="Retrieve"
# NEW (working): placeholder="PNR / Booking Reference", placeholder="GST Email ID",
#                button text="Retrieve GST Invoice"
XPATH_PNR_INPUT    = '//input[contains(@placeholder,"PNR")]'
XPATH_EMAIL_INPUT  = '//input[contains(@placeholder,"GST Email")]'
XPATH_RETRIEVE_BTN = '//button[contains(.,"Retrieve GST Invoice")]'

# Error messages the site shows when rate-limited or when details are wrong.
# These are checked after every form submit so we fail fast instead of hanging.
ERROR_TEXT_PATTERNS = [
    "sorry",
    "unable to retrieve",
    "unable to process",
    "try again later",
    "something went wrong",
    "no invoice",
    "not found",
    "invalid",
    "incorrect",
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("indigo")


# ================= INPUT & OUTPUT (CSV / EXCEL) =================

def find_input_file():
    """
    Auto-detects whether an Excel file (.xlsx, .xls) or CSV file is present
    in the directory, so users can drop in either without renaming.
    """
    candidates = [
        "pnr.xlsx",
        "pnr.csv",
        "pnr.xls",
        "input.xlsx",
        "input.csv",
    ]
    for c in candidates:
        full = os.path.join(BASE_DIR, c)
        if os.path.exists(full):
            return full
    # Default fallback
    return INPUT_FILE


def load_input_rows(filepath):
    """
    Loads raw rows from either:
      1. Excel file (.xlsx / .xlsm)
      2. Excel file saved with .csv extension (starts with PK zip magic bytes)
      3. Old Excel file (.xls)
      4. Standard CSV (handles utf-8, utf-8-sig, cp1252, latin1, etc.)
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Input file not found: {filepath}")

    with open(filepath, "rb") as f:
        magic = f.read(8)

    # XLSX / XLSM (ZIP format with PK\x03\x04 header)
    is_xlsx = magic.startswith(b"PK\x03\x04")
    # XLS (OLE2 format)
    is_xls = magic.startswith(b"\xd0\xcf\x11\xe0")

    rows = []
    if is_xlsx or filepath.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        with open(filepath, "rb") as f:
            data = io.BytesIO(f.read())
        wb = openpyxl.load_workbook(data, data_only=True)
        sheet = wb.active
        for r in sheet.iter_rows(values_only=True):
            if any(cell is not None and str(cell).strip() != "" for cell in r):
                rows.append([str(cell).strip() if cell is not None else "" for cell in r])

    elif is_xls or filepath.lower().endswith(".xls"):
        import pandas as pd
        df = pd.read_excel(filepath)
        rows = [list(df.columns)] + [
            [str(val).strip() if pd.notna(val) else "" for val in row] for row in df.values
        ]

    else:
        # Standard CSV with robust multi-encoding detection
        content = None
        for enc in ["utf-8-sig", "utf-8", "cp1252", "latin1"]:
            try:
                with open(filepath, "r", encoding=enc) as f:
                    content = f.read()
                    break
            except Exception:
                continue

        if content is None:
            with open(filepath, "r", encoding="latin1", errors="replace") as f:
                content = f.read()

        sample = content[:2048]
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
            delim = dialect.delimiter
        except Exception:
            delim = ","

        reader = csv.reader(io.StringIO(content), delimiter=delim)
        for r in reader:
            if any(cell.strip() for cell in r):
                rows.append([cell.strip() for cell in r])

    return rows


def parse_pnr_list(rows):
    """
    Extracts (pnr, email) tuples from loaded rows.
    Intelligently identifies columns by header name (e.g., 'TICKET NO', 'PNR', 'EMAIL', etc.)
    or defaults to column 0 for PNR and column 1 for Email.
    """
    if not rows:
        return []

    header = rows[0]
    pnr_col = 0
    email_col = 1
    has_header = False

    for idx, col_name in enumerate(header):
        cleaned = str(col_name).lower().replace("_", " ").replace("-", " ")
        if any(k in cleaned for k in ["pnr", "ticket", "booking", "ref"]):
            pnr_col = idx
            has_header = True
        elif any(k in cleaned for k in ["email", "mail"]):
            email_col = idx
            has_header = True

    data_rows = rows[1:] if has_header else rows
    parsed = []
    for r in data_rows:
        if len(r) <= pnr_col:
            continue
        pnr = str(r[pnr_col]).strip()
        # Clean numeric float values like '123456.0' if numbers are entered in Excel
        if pnr.endswith(".0") and pnr[:-2].isdigit():
            pnr = pnr[:-2]
        if not pnr or pnr.lower() in ["pnr", "ticket no", "ticket", "booking reference"]:
            continue
        email = str(r[email_col]).strip() if len(r) > email_col else ""
        parsed.append((pnr, email))

    return parsed


def init_files():
    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow(["PNR", "INVOICES"])
    with open(ERROR_FILE, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow(["PNR", "ERROR", "TIME"])


def load_done_pnrs():
    done = set()
    for path in (OUTPUT_FILE, ERROR_FILE):
        if os.path.exists(path):
            with open(path, newline="", encoding="utf-8", errors="replace") as f:
                rows = list(csv.reader(f))[1:]
                for r in rows:
                    if r:
                        done.add(r[0].strip())
    return done


def write_output(csv_lock, pnr, invs):
    with csv_lock:
        with open(OUTPUT_FILE, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([pnr, ",".join(invs)])


def write_error(csv_lock, pnr, msg):
    with csv_lock:
        with open(ERROR_FILE, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([pnr, msg, time.strftime("%H:%M:%S")])


# ================= DRIVER =================

def _get_chrome_major_version():
    """
    Reads the installed Chrome version from the registry/filesystem so we never
    need to hard-code version_main. Returns None if detection fails (uc will
    then attempt its own auto-detection).
    """
    import subprocess
    try:
        result = subprocess.run(
            ["powershell", "-Command",
             "(Get-Item 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe')"
             ".VersionInfo.FileVersion"],
            capture_output=True, text=True, timeout=10
        )
        ver = result.stdout.strip().split(".")[0]
        if ver.isdigit():
            return int(ver)
    except Exception:
        pass
    # Fallback: try x86 path
    try:
        result = subprocess.run(
            ["powershell", "-Command",
             "(Get-Item 'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe')"
             ".VersionInfo.FileVersion"],
            capture_output=True, text=True, timeout=10
        )
        ver = result.stdout.strip().split(".")[0]
        if ver.isdigit():
            return int(ver)
    except Exception:
        pass
    return None   # let undetected-chromedriver auto-detect


def _build_chrome_options(wid, profile_dir):
    """
    Builds a fresh ChromeOptions object every time it's called.
    IMPORTANT: uc.Chrome() mutates the options object internally, so you MUST
    create a new one for every attempt — reusing the same object causes the
    'cannot reuse ChromeOptions' error on any retry.
    """
    opts = uc.ChromeOptions()
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument(f"--window-size={WINDOW_SIZE}")
    opts.page_load_strategy = "eager"
    opts.add_argument("--disable-extensions")
    opts.add_argument("--disable-background-networking")
    opts.add_argument("--disable-background-timer-throttling")
    opts.add_argument("--disable-backgrounding-occluded-windows")
    opts.add_argument("--disable-renderer-backgrounding")
    opts.add_argument("--disable-sync")
    opts.add_argument("--disable-translate")
    opts.add_argument("--mute-audio")
    opts.add_argument("--metrics-recording-only")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    if HEADLESS:
        opts.add_argument("--headless=new")
    if KEEP_WINDOWS_OUT_OF_THE_WAY and not HEADLESS:
        x, y = OFFSCREEN_POSITION
        opts.add_argument(f"--window-position={x},{y}")
    opts.add_argument(f"--user-data-dir={profile_dir}")
    return opts


def create_driver(wid):
    """
    Each call creates a fully independent Chrome browser with its own isolated
    profile directory. Each worker lives in its own OS process so
    undetected_chromedriver's internal state is never shared.
    """
    # Auto-detect Chrome version — avoids the version_main mismatch error
    # when Chrome updates without the script being changed.
    chrome_version = _get_chrome_major_version()
    if chrome_version:
        log.info(f"[WORKER {wid}] Detected Chrome version: {chrome_version}")
    else:
        log.warning(f"[WORKER {wid}] Could not detect Chrome version; letting uc auto-detect.")

    profile_dir = tempfile.mkdtemp(prefix=f"uc_profile_w{wid}_")

    last_err = None
    for attempt in range(1, 4):
        # ⚠ Build a FRESH ChromeOptions for every attempt — uc.Chrome() mutates
        # the object internally, and reusing it on retries raises:
        # "you cannot reuse the ChromeOptions object"
        options = _build_chrome_options(wid, profile_dir)
        try:
            driver = uc.Chrome(
                options=options,
                version_main=chrome_version,  # auto-detected; None = uc auto-detects
                use_subprocess=False
            )
            driver.set_page_load_timeout(TASK_TIMEOUT)  # must be >> WAIT; IndiGo's summary
            driver.set_script_timeout(TASK_TIMEOUT)      # page redirect can take 60–90s
            driver._uc_profile_dir = profile_dir

            # Patch navigator.webdriver to False — the #1 Akamai/bot-detection signal.
            driver.execute_script(
                "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            )

            if KEEP_WINDOWS_OUT_OF_THE_WAY and not HEADLESS:
                try:
                    x, y = OFFSCREEN_POSITION
                    driver.set_window_position(x, y)
                    driver.minimize_window()
                except Exception:
                    pass

            try:
                driver.execute_cdp_cmd("Network.enable", {})
            except Exception:
                pass

            return driver
        except Exception as e:
            last_err = e
            log.warning(f"[WORKER {wid}] Driver creation attempt {attempt}/3 failed: {e}")
            time.sleep(3)

    raise RuntimeError(f"[WORKER {wid}] Could not create Chrome driver after 3 attempts: {last_err}")


def quit_driver(driver):
    profile_dir = getattr(driver, "_uc_profile_dir", None)
    try:
        driver.quit()
    except Exception:
        pass
    if profile_dir:
        shutil.rmtree(profile_dir, ignore_errors=True)


def wait_ready(driver, timeout=WAIT_SHORT):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
    except Exception:
        pass


def set_image_blocking(driver, block: bool):
    """
    Toggles image loading at the network level.
    Block images on the fast PNR search/lookup pages for speed;
    always re-enable before the invoice page that gets printed to PDF.
    """
    try:
        patterns = ["*.png", "*.jpg", "*.jpeg", "*.gif", "*.webp", "*.svg", "*.ico"] if block else []
        driver.execute_cdp_cmd("Network.setBlockedURLs", {"urls": patterns})
    except Exception:
        pass


def wait_images_loaded(driver, timeout=WAIT_SHORT):
    """Confirms every <img> on the page has actually finished downloading."""
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script(
                "return Array.from(document.images).every(img => img.complete && img.naturalWidth > 0)"
            )
        )
    except Exception:
        pass


def human_type(element, text):
    """
    Types text into an element with randomised per-keystroke delays to
    mimic human typing speed and fool keystroke-cadence bot detection.
    """
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(TYPING_DELAY_MIN, TYPING_DELAY_MAX))


def check_for_site_error(driver, wid="?"):
    """
    Returns an error message string if the site is showing a rate-limit /
    'sorry' error or 'details incorrect' toast, or None if the page looks normal.
    """
    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text.lower()
        for pattern in ERROR_TEXT_PATTERNS:
            if pattern in body_text:
                return f"Site error detected ('{pattern}' found in page)"

        # Also inspect visible alert/modal/toast elements
        alerts = driver.find_elements(
            By.CSS_SELECTOR,
            ".modal, [role='dialog'], .toast, .alert, .skyplus-alert, .error, .skyplus-toast"
        )
        for a in alerts:
            if a.is_displayed():
                txt = a.text.strip()
                lower = txt.lower()
                if "privacy" in lower or "cookie" in lower or not txt:
                    continue
                for pattern in ERROR_TEXT_PATTERNS:
                    if pattern in lower:
                        return f"Site error: {txt}"
    except Exception:
        pass
    return None


# ================= FETCH =================

def _dismiss_cookie_consent(driver):
    """Dismisses cookie consent banner if present so it doesn't block interactions."""
    try:
        btns = driver.find_elements(
            By.XPATH,
            '//button[contains(.,"Accept All") or contains(.,"Accept Essential Only")]'
        )
        for b in btns:
            if b.is_displayed():
                b.click()
                time.sleep(0.5)
                break
    except Exception:
        pass


def _fill_and_submit_pnr_form(driver, pnr, email, wlog, wid):
    wait = WebDriverWait(driver, WAIT_SHORT)
    pnr_in = wait.until(EC.presence_of_element_located((By.XPATH, XPATH_PNR_INPUT)))
    email_in = wait.until(EC.presence_of_element_located((By.XPATH, XPATH_EMAIL_INPUT)))

    driver.execute_script("arguments[0].scrollIntoView({behavior:'smooth',block:'center'});", pnr_in)
    time.sleep(0.3)

    wlog.info(f"[WORKER {wid}] {pnr}: typing PNR")
    pnr_in.clear()
    pnr_in.send_keys(pnr)
    time.sleep(0.3)

    wlog.info(f"[WORKER {wid}] {pnr}: typing email")
    email_in.clear()
    email_in.send_keys(email)
    time.sleep(0.4)

    # Dispatch events so Angular registers the input and enables the submit button
    driver.execute_script("""
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
    """, pnr_in)
    driver.execute_script("""
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
    """, email_in)
    time.sleep(0.5)

    wlog.info(f"[WORKER {wid}] {pnr}: clicking 'Retrieve GST Invoice'")
    retrieve_btn = wait.until(EC.presence_of_element_located((By.XPATH, XPATH_RETRIEVE_BTN)))
    driver.execute_script("arguments[0].scrollIntoView({behavior:'smooth',block:'center'});", retrieve_btn)
    time.sleep(0.3)
    try:
        retrieve_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", retrieve_btn)


def fetch_invoices(driver, pnr, email, wid="?"):
    wlog = logging.getLogger(f"indigo.w{wid}")

    wlog.info(f"[WORKER {wid}] {pnr}: opening PNR lookup page")
    set_image_blocking(driver, BLOCK_IMAGES)

    url = "https://www.goindigo.in/view-gst-invoice.html"
    driver.get(url)
    wait_ready(driver)
    time.sleep(random.uniform(1.5, 2.5))
    _dismiss_cookie_consent(driver)

    # First submit attempt
    wlog.info(f"[WORKER {wid}] {pnr}: submitting PNR lookup form...")
    _fill_and_submit_pnr_form(driver, pnr, email, wlog, wid)

    # Wait briefly to observe if it proceeds or needs the refresh mechanism
    time.sleep(2.5)

    # Check if summary page reached immediately or if page refresh is needed
    current_url = driver.current_url
    if "summary" not in current_url:
        wlog.info(f"[WORKER {wid}] {pnr}: executing page refresh mechanism...")
        driver.refresh()
        wait_ready(driver)
        time.sleep(random.uniform(2.5, 3.5))
        _dismiss_cookie_consent(driver)

        wlog.info(f"[WORKER {wid}] {pnr}: resubmitting form after refresh...")
        _fill_and_submit_pnr_form(driver, pnr, email, wlog, wid)

    # Poll for summary page
    deadline = time.time() + WAIT
    navigated = False
    while time.time() < deadline:
        try:
            current_url = driver.current_url
            if "summary" in current_url:
                navigated = True
                wlog.info(f"[WORKER {wid}] {pnr}: navigated to summary page: {current_url}")
                break
        except Exception:
            pass

        # Check for visible error messages (like 'details incorrect')
        site_err = check_for_site_error(driver, wid)
        if site_err and "sorry" not in site_err.lower():
            raise Exception(site_err)

        time.sleep(1.0)

    if not navigated:
        # Check for error text
        site_err = check_for_site_error(driver, wid)
        if site_err:
            raise Exception(site_err)
        raise Exception(f"Summary page not reached after submit and refresh mechanism for {pnr}")

    # Wait for invoice content to render on Angular page
    wlog.info(f"[WORKER {wid}] {pnr}: waiting for invoice details to render...")
    try:
        WebDriverWait(driver, WAIT).until(
            lambda d: "invoice number" in d.find_element(By.TAG_NAME, "body").text.lower()
        )
    except Exception:
        pass

    # Extract invoice numbers from text
    body_text = driver.find_element(By.TAG_NAME, "body").text
    invoices = re.findall(r"Invoice Number:\s*([A-Za-z0-9]+)", body_text)

    # Fallback to DOM elements if regex found none
    if not invoices:
        elements = driver.find_elements(By.XPATH, '//*[contains(text(),"Invoice Number:")]')
        for el in elements:
            txt = el.text.strip()
            if "Invoice Number:" in txt:
                inv_num = txt.split("Invoice Number:")[-1].strip().split()[0]
                if inv_num and inv_num not in invoices:
                    invoices.append(inv_num)

    wlog.info(f"[WORKER {wid}] {pnr}: found {len(invoices)} invoice(s): {invoices}")
    return invoices


# ================= PDF =================

def download_pdf(driver, inv, wid="?"):
    path = os.path.join(INVOICE_DIR, f"{inv}.pdf")

    if os.path.exists(path) and os.path.getsize(path) > PDF_MIN_SIZE:
        return

    wlog = logging.getLogger(f"indigo.w{wid}")
    wlog.info(f"[WORKER {wid}] downloading invoice {inv}")

    # Always load images for the invoice page (needed for IndiGo logo in PDF).
    set_image_blocking(driver, False)
    driver.get(f"https://www.goindigo.in/view-gst-invoice/print.html?invoiceNumber={inv}")
    wait_ready(driver)
    wait_images_loaded(driver)

    # Check for errors on the print page too
    site_err = check_for_site_error(driver, wid)
    if site_err:
        raise Exception(f"Error on invoice print page: {site_err}")

    pdf = driver.execute_cdp_cmd("Page.printToPDF", {"printBackground": True})
    data = base64.b64decode(pdf["data"])

    if len(data) < PDF_MIN_SIZE:
        raise Exception("Blank PDF — invoice page may not have loaded correctly")

    with open(path, "wb") as f:
        f.write(data)


# ================= TASK (watched for hangs) =================

def process_pnr(driver, pnr, email, wid="?"):
    invoices = fetch_invoices(driver, pnr, email, wid)
    if not invoices:
        raise Exception("No invoices found")
    for inv in invoices:
        download_pdf(driver, inv, wid)
    return invoices


def run_with_timeout(driver, pnr, email, timeout, wid="?"):
    """
    Runs process_pnr in a helper daemon thread and waits up to `timeout` seconds.
    If it hangs, we reclaim control, recycle the driver, and move on — so one
    bad PNR can never freeze a worker permanently.
    """
    result = {}

    def runner():
        try:
            result["invoices"] = process_pnr(driver, pnr, email, wid)
        except Exception as e:
            result["error"] = e

    t = threading.Thread(target=runner, daemon=True)
    t.start()
    t.join(timeout)

    if t.is_alive():
        raise TimeoutError(f"Task exceeded {timeout}s (site likely stuck)")

    if "error" in result:
        raise result["error"]

    return result.get("invoices", [])


# ================= WORKER PROCESS =================

def worker(wid, task_queue, csv_lock):
    """
    Runs in its own OS process with its own independent Chrome browser.
    Pulls (pnr, email) tuples from the shared multiprocessing queue until
    it sees the sentinel None.
    """
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )
    wlog = logging.getLogger(f"indigo.w{wid}")
    wlog.info(f"[WORKER {wid}] starting own browser...")

    driver = create_driver(wid)
    wlog.info(f"[WORKER {wid}] browser ready.")
    processed_since_restart = 0

    while True:
        try:
            item = task_queue.get(timeout=5)
        except Empty:
            continue

        if item is None:
            break

        pnr, email = item
        attempt = 0
        succeeded = False
        last_error = None

        while attempt <= MAX_RETRIES and not succeeded:
            attempt += 1
            try:
                wlog.info(f"[WORKER {wid}] Processing {pnr} (attempt {attempt}/{MAX_RETRIES + 1})")
                invoices = run_with_timeout(driver, pnr, email, TASK_TIMEOUT, wid)
                write_output(csv_lock, pnr, invoices)
                succeeded = True
                processed_since_restart += 1

                # Human-like pause between successful PNRs.
                # This is the most important rate-limit mitigation.
                delay = random.uniform(DELAY_BETWEEN_PNRS_MIN, DELAY_BETWEEN_PNRS_MAX)
                wlog.info(f"[WORKER {wid}] Pausing {delay:.1f}s before next PNR...")
                time.sleep(delay)

            except Exception as e:
                last_error = str(e)
                wlog.warning(f"[WORKER {wid}] {pnr} failed (attempt {attempt}): {last_error}")

                # If the error is definitely invalid details from the site, don't waste time retrying
                if any(k in last_error.lower() for k in ["incorrect", "invalid", "not found"]):
                    wlog.info(f"[WORKER {wid}] Booking details invalid for {pnr}, skipping retries.")
                    break

                # After any failure, wait longer before retrying — give the site time
                # to "forget" about us and reset any rate-limit counters.
                cooldown = random.uniform(DELAY_AFTER_ERROR_MIN, DELAY_AFTER_ERROR_MAX)
                wlog.info(f"[WORKER {wid}] Cooling down {cooldown:.1f}s after error...")
                time.sleep(cooldown)

                # Recycle the driver to get a fresh session/cookies
                quit_driver(driver)
                driver = create_driver(wid)
                processed_since_restart = 0

        if not succeeded:
            write_error(csv_lock, pnr, last_error or "Unknown error")

        if processed_since_restart >= RESTART_AFTER:
            wlog.info(f"[WORKER {wid}] Recycling browser after {RESTART_AFTER} PNRs")
            quit_driver(driver)
            driver = create_driver(wid)
            processed_since_restart = 0

    quit_driver(driver)
    wlog.info(f"[WORKER {wid}] stopped.")


# ================= MAIN =================

def main():
    log.info("========== INDIGO INVOICE DOWNLOADER START (multi-browser) ==========")

    if RESUME and os.path.exists(OUTPUT_FILE) and os.path.exists(ERROR_FILE):
        done = load_done_pnrs()
        log.info(f"Resume mode: {len(done)} PNR(s) already processed, will be skipped.")
    else:
        init_files()
        done = set()

    input_path = find_input_file()
    log.info(f"Using input file: {input_path}")
    raw_rows = load_input_rows(input_path)
    all_pnrs = parse_pnr_list(raw_rows)
    log.info(f"Loaded {len(all_pnrs)} record(s) from input file.")

    pending = []
    for pnr, email in all_pnrs:
        if pnr in done:
            continue
        pending.append((pnr, email))

    log.info(f"Queued {len(pending)} PNR(s) for processing across {WORKERS} independent browser(s).")

    if not pending:
        log.info("Nothing to do.")
        return

    # PRE-WARM: patch the undetected_chromedriver executable once here,
    # in the main process, before any worker processes exist. This avoids the
    # one-time driver-patch race condition when workers launch in parallel.
    log.info("Pre-warming Chrome driver (one-time patch check)...")
    try:
        warm = create_driver("prewarm")
        quit_driver(warm)
        log.info("Driver pre-warm successful.")
    except Exception as e:
        log.error(f"Driver pre-warm failed: {e}")
        log.error(r"Try deleting %APPDATA%\undetected_chromedriver\ and rerunning.")
        return

    ctx = mp.get_context("spawn")   # 'spawn' is required/safest on Windows
    task_queue = ctx.Queue()
    csv_lock = ctx.Lock()

    for item in pending:
        task_queue.put(item)
    for _ in range(WORKERS):
        task_queue.put(None)  # one sentinel per worker to signal shutdown

    processes = []
    for wid in range(WORKERS):
        p = ctx.Process(target=worker, args=(wid, task_queue, csv_lock))
        p.start()
        processes.append(p)
        if wid < WORKERS - 1:
            log.info(f"Staggering next worker start by {STAGGER_START}s...")
            time.sleep(STAGGER_START)

    for p in processes:
        p.join()

    log.info("========== INDIGO INVOICE DOWNLOADER END ==========")


if __name__ == "__main__":
    mp.freeze_support()   # harmless on non-frozen scripts, required if ever built into an .exe
    main()
