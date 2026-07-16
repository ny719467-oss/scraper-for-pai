"""
PAI Portal Gram Panchayat Scraper - Web App
=============================================
A small Flask app that wraps the Selenium scraping logic in a web UI.

Run:
    pip install -r requirements.txt
    python app.py
Then open http://127.0.0.1:5000 in your browser.

How it works:
    - The frontend lets you pick a State from a dropdown (list pulled directly
      from the live portal page).
    - Clicking "Start Scraping" kicks off a background thread that drives a
      real Chrome browser through the State -> District -> Block cascade,
      exactly like the original script, but resilient to stale elements and
      slow AJAX postbacks.
    - Progress (current district/block, rows collected, live log lines) is
      exposed via /api/status and polled by the page every 2 seconds.
    - When done, a "Download Excel" button appears and hits /api/download.
"""

import os
import re
import time
import threading
import traceback
from datetime import datetime

import pandas as pd
from bs4 import BeautifulSoup
from flask import Flask, jsonify, request, render_template, send_file

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
)

try:
    # Optional, but makes life much easier - auto-downloads matching chromedriver
    from webdriver_manager.chrome import ChromeDriverManager
    HAVE_WDM = True
except ImportError:
    HAVE_WDM = False

BASE_URL = "https://pai.gov.in/PS/Public/TW-GP-New.aspx"
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Full state list from the live portal (visible option text, exact match required
# for Selenium's select_by_visible_text). Keep this in sync if the portal changes.
STATE_LIST = [
    "Andaman And Nicobar Islands [35, T-3]",
    "Andhra Pradesh [28, T-3]",
    "Arunachal Pradesh [12, T-2]",
    "Assam [18, T-3]",
    "Bihar [10, T-3]",
    "Chhattisgarh [22, T-3]",
    "Goa [30, T-2]",
    "Gujarat [24, T-3]",
    "Haryana [6, T-3]",
    "Himachal Pradesh [2, T-3]",
    "Jammu And Kashmir [1, T-3]",
    "Jharkhand [20, T-3]",
    "Karnataka [29, T-3]",
    "Kerala [32, T-3]",
    "Ladakh [37, T-3]",
    "Lakshadweep [31, T-2]",
    "Madhya Pradesh [23, T-3]",
    "Maharashtra [27, T-3]",
    "Manipur [14, T-2]",
    "Meghalaya [17, T-1]",
    "Mizoram [15, T-1]",
    "Nagaland [13, T-1]",
    "Odisha [21, T-3]",
    "Puducherry [34, T-2]",
    "Punjab [3, T-3]",
    "Rajasthan [8, T-3]",
    "Sikkim [11, T-2]",
    "Tamil Nadu [33, T-3]",
    "Telangana [36, T-3]",
    "The Dadra And Nagar Haveli And Daman And Diu [38, T-2]",
    "Tripura [16, T-3]",
    "Uttarakhand [5, T-3]",
    "Uttar Pradesh [9, T-3]",
    "West Bengal [19, T-3]",
]

app = Flask(__name__)

# ---------------------------------------------------------------------------
# Global job state (single-job-at-a-time, kept simple on purpose)
# ---------------------------------------------------------------------------
job_lock = threading.Lock()
job_state = {
    "running": False,
    "stop_requested": False,
    "state_selected": None,
    "current_district": None,
    "current_block": None,
    "districts_total": 0,
    "districts_done": 0,
    "rows_collected": 0,
    "logs": [],           # last N log lines shown in the UI
    "done": False,
    "error": None,
    "output_file": None,
    "started_at": None,
    "finished_at": None,
    "preview_columns": None,
    "preview_rows": None,
    "preview_truncated": False,
}

PREVIEW_ROW_LIMIT = 100


def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    job_state["logs"].append(line)
    # keep log list bounded
    if len(job_state["logs"]) > 500:
        job_state["logs"] = job_state["logs"][-500:]


def is_placeholder_option(text):
    """True for dropdown placeholder entries (e.g. 'Select', '- Select -', or
    the Hindi equivalent when the portal renders in that language). Real
    options on this portal always carry a numeric code in brackets
    (state/district/block ids), so 'no digit in the text' is a reliable,
    language-independent signal - unlike matching literal placeholder words,
    which breaks the moment the portal switches to Hindi."""
    return not re.search(r"\d", text or "")


def clean_state_name(raw_text):
    """'Uttarakhand [5, T-3]' -> 'Uttarakhand'"""
    if not raw_text:
        return raw_text
    return raw_text.split("[")[0].strip()


def normalize_option_key(raw_text):
    """Loose key for matching an option's name regardless of bracket suffix
    drift, '&' vs 'And', case, or extra whitespace."""
    text = clean_state_name(raw_text).lower().replace("&", "and")
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def extract_option_code(raw_text):
    """Pull the leading numeric id out of '... [35, T-3]' / '... [35, टी-3]' /
    '... [603]'. This id is language-independent, unlike the option's name -
    this portal renders State/District/Block dropdowns in Hindi or English
    seemingly per-request (observed flipping between an initial load and a
    later postback on the very same dropdown), but the numeric code stays put
    either way."""
    match = re.search(r"\[\s*(\d+)", raw_text or "")
    return match.group(1) if match else None


def select_dropdown_option(select_obj, target_text):
    """Select the option matching target_text in a State/District/Block
    dropdown. Tries, in order:
    1. Exact visible text match (fast path).
    2. Normalized name match (handles wording/whitespace/case drift).
    3. Numeric code match (handles the dropdown's language flipping between
       when its options were read and when we come back to select one - the
       bracketed id, unlike the name, doesn't change across languages).
    Returns the actual option text that was selected."""
    try:
        select_obj.select_by_visible_text(target_text)
        return target_text
    except NoSuchElementException:
        target_key = normalize_option_key(target_text)
        for option in select_obj.options:
            if normalize_option_key(option.text) == target_key:
                select_obj.select_by_visible_text(option.text)
                return option.text

        target_code = extract_option_code(target_text)
        if target_code:
            for option in select_obj.options:
                if extract_option_code(option.text) == target_code:
                    select_obj.select_by_visible_text(option.text)
                    return option.text
        raise


def switch_to_english(driver, wait):
    """The portal auto-translates itself via the government's Bhashini widget,
    which defaults to Hindi (`initial_preferred_language="hi"` in its script
    tag) and can flip individual elements back to English inconsistently
    between requests - this is what caused dropdown/table text to randomly
    switch languages mid-scrape. Force the widget to English once, right
    after page load, so the whole run stays in one consistent language.
    Best-effort: select_dropdown_option()'s code-based fallback still covers
    any residual drift if this widget interaction fails or changes shape."""
    try:
        toggle = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".bhashini-dropdown-btn")))
        toggle.click()
        english_option = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, '#bhashiniLanguageDropdown .language-option[data-value="en"]')
        ))
        english_option.click()
        time.sleep(1.5)  # let the widget's translation pass finish
    except (TimeoutException, NoSuchElementException):
        pass


def get_selected_option_text(driver, element_id):
    """Return the visible text of the currently selected <option> for a <select>."""
    try:
        sel = Select(driver.find_element(By.ID, element_id))
        return sel.first_selected_option.text.strip()
    except Exception:
        return None


def sniff_state_label_on_page(driver, fallback_text):
    """
    Best-effort: some result views render a confirmation label/header with the
    state name once results load (e.g. <span id='lblState'>Uttarakhand</span>).
    We don't know the exact id on this portal, so we try a few common patterns
    and fall back to the dropdown's own selected text if nothing is found.
    """
    candidate_ids = ["lblState", "lbl_State", "lblStateName", "spanState", "ctl00_lblState"]
    for cid in candidate_ids:
        try:
            el = driver.find_element(By.ID, cid)
            text = el.text.strip()
            if text:
                return text
        except NoSuchElementException:
            continue
    return fallback_text


def wait_for_options_loaded(wait, element_id, min_real_options=1, timeout=15, stable_checks=2):
    """Wait until a dependent dropdown (state/district/block) has finished its
    AJAX postback. Requires at least `min_real_options` real (non-placeholder)
    options AND requires the option list to read identically across
    `stable_checks` consecutive polls - a naive 'len(options) >= N' check can
    pass on a transient in-between state (leftover options from before the
    postback, or a briefly-inserted loading entry) before the real list has
    actually settled."""
    end_time = time.time() + timeout
    last_err = None
    previous_texts = None
    stable_count = 0
    while time.time() < end_time:
        try:
            dd = Select(wait._driver.find_element(By.ID, element_id))
            texts = tuple(o.text.strip() for o in dd.options)
            real_count = sum(1 for t in texts if not is_placeholder_option(t))
            if real_count >= min_real_options:
                stable_count = stable_count + 1 if texts == previous_texts else 1
                previous_texts = texts
                if stable_count >= stable_checks:
                    return dd
            else:
                previous_texts, stable_count = None, 0
        except (StaleElementReferenceException, NoSuchElementException) as e:
            last_err = e
            previous_texts, stable_count = None, 0
        time.sleep(0.4)
    if last_err:
        raise last_err
    raise TimeoutException(f"{element_id} did not populate in time")


def build_driver(headless=True):
    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1400,1000")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    if HAVE_WDM:
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=options)
    # Falls back to chromedriver already on PATH
    return webdriver.Chrome(options=options)


def parse_results_table(html):
    """Parse the GVdata results grid ourselves via BeautifulSoup instead of
    pd.read_html(). The grid uses DataTables Responsive, which marks
    overflow columns 'dtr-hidden' with inline display:none while keeping
    their real text content in the DOM (visually hidden behind a '+' expand
    toggle) - pd.read_html silently drops those columns instead of ignoring
    the styling, which was truncating every row down to ~13 of the full ~23
    columns. BeautifulSoup reads the raw markup and isn't affected by CSS."""
    soup = BeautifulSoup(html, "html5lib")
    headers = [th.get_text(strip=True) for th in soup.select("thead th")]
    rows = [[td.get_text(strip=True) for td in tr.find_all("td")] for tr in soup.select("tbody tr")]
    raw_df = pd.DataFrame(rows, columns=headers)
    columns = {}
    for col in raw_df.columns:
        converted = pd.to_numeric(raw_df[col], errors="coerce")
        columns[col] = converted if converted.notna().all() else raw_df[col]
    return pd.DataFrame(columns)


def run_scrape_job(state_name, headless=True):
    all_rows = []
    driver = None
    try:
        job_state["running"] = True
        job_state["stop_requested"] = False
        job_state["done"] = False
        job_state["error"] = None
        job_state["rows_collected"] = 0
        job_state["districts_done"] = 0
        job_state["output_file"] = None
        job_state["started_at"] = datetime.now().isoformat()
        job_state["logs"] = []
        job_state["preview_columns"] = None
        job_state["preview_rows"] = None
        job_state["preview_truncated"] = False

        log("Launching browser...")
        driver = build_driver(headless=headless)
        wait = WebDriverWait(driver, 20)

        log(f"Opening portal: {BASE_URL}")
        driver.get(BASE_URL)

        switch_to_english(driver, wait)

        state_dropdown = wait_for_options_loaded(wait, "ddl_State")
        try:
            selected_text = select_dropdown_option(state_dropdown, state_name)
        except NoSuchElementException:
            live_options = [o.text.strip() for o in state_dropdown.options]
            log(f"State selection failed. Live ddl_State options were: {live_options}")
            raise
        if selected_text != state_name:
            log(f"Note: portal's option text has drifted from the app's list "
                f"('{state_name}' -> matched '{selected_text}')")

        # Read the state name straight back off the page (the ask: "state name
        # also from website text", not just what we clicked). Kept as-is even
        # if the portal happens to be rendering in Hindi right now - it's the
        # raw/confirmation columns, not the primary "State" column.
        confirmed_state_text = get_selected_option_text(driver, "ddl_State") or selected_text
        # The primary "State" column always uses the name we requested (from
        # STATE_LIST, always English) so output stays consistent even though
        # the portal's own dropdown language can vary between runs.
        clean_state = clean_state_name(state_name)
        job_state["state_selected"] = clean_state
        log(f"Selected state (from page): {confirmed_state_text}")

        district_dropdown = wait_for_options_loaded(wait, "ddl_District")
        districts = [o.text.strip() for o in district_dropdown.options if not is_placeholder_option(o.text)]
        job_state["districts_total"] = len(districts)
        log(f"Found {len(districts)} districts for {clean_state}")

        for district in districts:
            if job_state["stop_requested"]:
                log("Stop requested - halting.")
                break

            job_state["current_district"] = district
            try:
                # Re-verify (not just re-locate) the dropdown before selecting:
                # right after a previous block's search postback, the element
                # can still be settling and a bare presence check hands back a
                # reference that goes stale moments later.
                district_dropdown = wait_for_options_loaded(wait, "ddl_District")
                select_dropdown_option(district_dropdown, district)
                log(f"District selected: {district}")

                block_dropdown = wait_for_options_loaded(wait, "ddl_Block")
                blocks = [o.text.strip() for o in block_dropdown.options if not is_placeholder_option(o.text)]

                for block in blocks:
                    if job_state["stop_requested"]:
                        break
                    job_state["current_block"] = block
                    try:
                        block_dropdown = wait_for_options_loaded(wait, "ddl_Block")
                        select_dropdown_option(block_dropdown, block)

                        search_btn = wait.until(EC.element_to_be_clickable((By.ID, "btnSubmit")))
                        driver.execute_script("arguments[0].scrollIntoView(true);", search_btn)
                        driver.execute_script("arguments[0].click();", search_btn)

                        # id="GVdata" is the actual results grid. A generic
                        # "//table" xpath was matching a hidden, always-empty
                        # "TableExport" helper table that renders earlier in
                        # the DOM, which is why every block used to come back
                        # with 0 rows.
                        table = wait.until(EC.presence_of_element_located((By.ID, "GVdata")))
                        html = table.get_attribute("outerHTML")
                        df = parse_results_table(html)

                        page_confirmed_state = sniff_state_label_on_page(driver, confirmed_state_text)

                        df.insert(0, "State", clean_state)
                        df.insert(1, "State_RawText", confirmed_state_text)
                        df.insert(2, "State_PageConfirmed", clean_state_name(page_confirmed_state))
                        df.insert(3, "District", district)
                        df.insert(4, "Block", block)

                        all_rows.append(df)
                        job_state["rows_collected"] += len(df)
                        log(f"  + {district} / {block}: {len(df)} rows")

                    except (TimeoutException, NoSuchElementException, StaleElementReferenceException):
                        log(f"  [skip] {district} / {block} (no data / load issue)")
                        continue

            except (TimeoutException, NoSuchElementException, StaleElementReferenceException):
                log(f"[skip] District: {district} (dropdown issue)")
            finally:
                job_state["districts_done"] += 1

        # Save results
        if all_rows:
            final_df = pd.concat(all_rows, ignore_index=True)
        else:
            final_df = pd.DataFrame()

        safe_state = clean_state.replace(" ", "_")
        filename = f"pai_{safe_state}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        filepath = os.path.join(OUTPUT_DIR, filename)
        final_df.to_excel(filepath, index=False)

        job_state["output_file"] = filepath
        log(f"Saved {len(final_df)} total rows to {filename}")

        if not final_df.empty:
            preview_df = final_df.head(PREVIEW_ROW_LIMIT).fillna("")
            job_state["preview_columns"] = [str(c) for c in preview_df.columns]
            job_state["preview_rows"] = preview_df.values.tolist()
            job_state["preview_truncated"] = len(final_df) > PREVIEW_ROW_LIMIT

    except Exception as e:
        job_state["error"] = str(e)
        log(f"ERROR: {e}")
        log(traceback.format_exc())
    finally:
        if driver is not None:
            driver.quit()
        job_state["running"] = False
        job_state["done"] = True
        job_state["finished_at"] = datetime.now().isoformat()
        job_state["current_district"] = None
        job_state["current_block"] = None


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html", states=STATE_LIST)


@app.route("/api/start", methods=["POST"])
def api_start():
    data = request.get_json(force=True)
    state_name = data.get("state")
    headless = data.get("headless", True)

    if not state_name or state_name not in STATE_LIST:
        return jsonify({"ok": False, "error": "Invalid or missing state"}), 400

    with job_lock:
        if job_state["running"]:
            return jsonify({"ok": False, "error": "A scrape job is already running"}), 409
        thread = threading.Thread(target=run_scrape_job, args=(state_name, headless), daemon=True)
        thread.start()

    return jsonify({"ok": True})


@app.route("/api/stop", methods=["POST"])
def api_stop():
    job_state["stop_requested"] = True
    return jsonify({"ok": True})


@app.route("/api/status")
def api_status():
    progress_pct = 0
    if job_state["districts_total"]:
        progress_pct = round(100 * job_state["districts_done"] / job_state["districts_total"], 1)

    return jsonify({
        "running": job_state["running"],
        "done": job_state["done"],
        "error": job_state["error"],
        "state_selected": job_state["state_selected"],
        "current_district": job_state["current_district"],
        "current_block": job_state["current_block"],
        "districts_total": job_state["districts_total"],
        "districts_done": job_state["districts_done"],
        "progress_pct": progress_pct,
        "rows_collected": job_state["rows_collected"],
        "logs": job_state["logs"][-200:],
        "output_ready": bool(job_state["output_file"] and os.path.exists(job_state["output_file"] or "")),
        "output_filename": os.path.basename(job_state["output_file"]) if job_state["output_file"] else None,
        "preview_columns": job_state["preview_columns"],
        "preview_rows": job_state["preview_rows"],
        "preview_truncated": job_state["preview_truncated"],
    })


@app.route("/api/download")
def api_download():
    if not job_state["output_file"] or not os.path.exists(job_state["output_file"]):
        return jsonify({"ok": False, "error": "No file ready yet"}), 404
    return send_file(
        job_state["output_file"],
        as_attachment=True,
        download_name=os.path.basename(job_state["output_file"]),
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    app.run(host="0.0.0.0", port=port, debug=debug)
