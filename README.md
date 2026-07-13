# PAI Portal Gram Panchayat Scraper — Web App

A local web app version of the Selenium scraper for
`https://pai.gov.in/PS/Public/TW-GP-New.aspx`.

## What it does

1. Open the app in your browser, pick a **State/UT** from the dropdown
   (list pulled from the live portal).
2. Click **Start Scraping**. In the background it launches Chrome and:
   - Selects the state
   - Loops through every **District**
   - For each district, loops through every **Block**
   - Clicks Search and reads the results table
   - Reads the state name **back off the page itself** (not just what was
     clicked) and stamps every row with:
     - `State` – cleaned name (e.g. `Uttarakhand`)
     - `State_RawText` – exact dropdown text (e.g. `Uttarakhand [5, T-3]`)
     - `State_PageConfirmed` – value read from an on-page confirmation label
       if the portal renders one, otherwise falls back to the dropdown text
     - `District`, `Block`
3. Progress (current district/block, rows collected, live log) updates
   every 2 seconds on the page.
4. When finished, a **Download Excel** button appears with all rows saved
   to `output/pai_<State>_<timestamp>.xlsx`.

## Setup

```bash
cd pai_scraper_app
python3 -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

You need Google Chrome installed. `webdriver-manager` (included in
requirements.txt) will automatically download the matching chromedriver —
you don't need to install chromedriver separately.

## Run

```bash
python app.py
```

Then open **http://127.0.0.1:5000** in your browser.

## Notes / things you may want to tune

- **Headless toggle**: uncheck "Run browser headless" on the page if you
  want to watch Chrome click through the dropdowns live — useful the first
  time, to confirm the portal's behavior matches expectations.
- **`sniff_state_label_on_page()`** in `app.py` tries a few common element
  IDs (`lblState`, `lbl_State`, etc.) to find an on-page confirmation label
  for the state name. The exact ID isn't documented, so if you know the
  real one, add it to the `candidate_ids` list for a more precise
  `State_PageConfirmed` value. Until then it safely falls back to the
  dropdown's own text.
- **One job at a time**: the app runs a single scrape job at a time by
  design (a government portal + many rapid concurrent Selenium sessions is
  a bad combination). Wait for one state to finish (or hit Stop) before
  starting the next.
- **Stop button**: sets a flag checked between blocks/districts — it won't
  kill mid-request, but it'll stop cleanly at the next safe point and still
  save whatever was collected so far.
- **Output location**: files land in `pai_scraper_app/output/`.
