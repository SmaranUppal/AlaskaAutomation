"""
Alaska Court Records Company Name Search Automation
======================================================
Site: records.courts.alaska.gov/eaccess

SETUP (one-time):
    pip install playwright
    playwright install chromium

RUN:
    python alaska_court_search.py

Results are printed to the console and saved to alaska_results.csv.
"""

import asyncio
import openpyxl
from datetime import date, timedelta
from pathlib import Path

from playwright.async_api import async_playwright, TimeoutError as PWTimeout


# ── Configuration ──────────────────────────────────────────────────────────────

NAMES: list[str] = [
    "J G W",
    "J.G. W",
    "J. G. W",
    "J.G.W",
    "JG W",
    "JGW",
    "J. G.W",
    "J GW",
    "DRB Cap",
    "Stone Street",
    "AA Ron I",
    "Abactor",
    "Abidole",
    "Adenna Med",
    "Adventura",
    "AGPI",
    "Aikman Structured Finance",
    "Annuity Transfers Ltd",
    "Apis Management",
    "Atlas Legal Funding III LP",
    "AXE Finance",
    "B.A.W.21",
    "B.R. Wright",
    "BHG Structured",
    "Bifco",
    "Blue Grape",
    "Catalina Structured Funding",
    "Concordis Group Limited",
    "Conrad Factoring",
    "Cornerstone Funding",
    "Fast Annuity S",
    "FL Assignments Corp",
    "G.D.T.R.F.B.",
    "G7 Crescenta",
    "Genex Capital Corp",
    "GJ 123",
    "Greenwood Funding",
    "Grier I",
    "Hakstol Group",
    "Hiddenview Ent, LLC",
    "JLC Capital Funding",
    "KN Direct Capital",
    "Lane Nimitz",
    "Lasko LLC",
    "Lasko, LLC",
    "Leaf 002 LLC",
    "Legere LLC",
    "Legere, LLC",
    "Lottery Funding",
    "M McDougall LLC",
    "M McDougall, LLC",
    "Majestic Funding",
    "Mic-Bry8",
    "Olive Branch Funding",
    "Palermo Group",
    "Palm Green Closing",
    "Palm Harbor",
    "Passira Mal",
    "Patriot Settlement",
    "QLS Funding",
    "Reliance Funding",
    "Rocorp Corporation",
    "RSL Funding",
    "Savannah Settlements",
    "Sempra Finance",
    "Seneca Originations",
    "SeneOne LLC",
    "Settlement Capital Corp",
    "Settlement Status",
    "Somerton LLC",
    "Somerton, LLC",
    "Stratcap Investments",
    "Stratton Asset",
    "Structured Asset",
    "TKD LLC",
    "TKD, LLC",
    "TRM V LLC",
    "TRM V, LLC",
    "Tybenz LLC",
    "Tybenz, LLC",
    "Uber Funding",
    "Vintage Equity Group",
    "Wepaymore Funding",
    "Zakho Way",
    "GREAT PLAINS MANAGEMENT CORPORATION",
    "RD FITZ LLC",
    "RD FITZ, LLC",
    "GA OFF LLC",
    "GA OFF, LLC",
    "Assured Management Corporation",
    "BENTZEN F",
]

DAYS_BACK: int = 60          # file date begin = today minus this many days
OUTPUT_XLSX: str = "alaska_results.xlsx"
HEADLESS: bool = False        # set True to run without a visible browser window

# Timeouts & pacing (milliseconds)
PAGE_LOAD_TIMEOUT = 15_000
ELEMENT_TIMEOUT   = 10_000
AJAX_SETTLE_DELAY = 1_500    # wait after typing Begin date for Wicket Ajax to settle
ACTION_DELAY      = 800
SEARCH_DELAY      = 2_000    # pause between consecutive name searches


# ── Date helper ────────────────────────────────────────────────────────────────

def file_date_begin(days_back: int) -> str:
    """Return MM/DD/YYYY for today minus days_back."""
    d = date.today() - timedelta(days=days_back)
    return d.strftime("%m/%d/%Y")


# ── Selectors ──────────────────────────────────────────────────────────────────
# Element IDs are session-dynamic on this Wicket site, so we use stable
# name/data/value attributes throughout.

SEL_SEARCH_CASES_BTN = "a.anchorButton"
SEL_COMPANY_FIELD    = "input[name='companyName']"
SEL_DATE_BEGIN       = "input[data='dateInputBegin4']"   # File Date Begin
SEL_SUBMIT           = "input[value='Search']"           # the Search button


# ── Logging ────────────────────────────────────────────────────────────────────

def log(msg: str) -> None:
    from datetime import datetime
    print(f"[{datetime.now():%H:%M:%S}] {msg}")


# ── Core helpers ───────────────────────────────────────────────────────────────

async def type_into_field(page, selector: str, value: str) -> None:
    """
    Click a field, clear it, then type character-by-character.
    Avoids page.fill() which doesn't reliably trigger Wicket's onchange Ajax.
    """
    el = page.locator(selector)
    await el.wait_for(state="visible", timeout=ELEMENT_TIMEOUT)
    await el.click()
    await page.keyboard.press("Control+A")
    await page.keyboard.press("Delete")
    await page.keyboard.type(value, delay=50)


async def click_name_tab(page) -> None:
    tab = page.get_by_role("link", name="Name", exact=True)
    await tab.wait_for(state="visible", timeout=ELEMENT_TIMEOUT)
    await tab.click()
    await page.wait_for_timeout(ACTION_DELAY)


async def search_one(page, company_name: str, begin: str) -> list[dict]:
    log(f"\n🔍  Searching: '{company_name}'  |  file date from {begin}")

    # Company Name
    await type_into_field(page, SEL_COMPANY_FIELD, company_name)
    await page.wait_for_timeout(ACTION_DELAY)

    # File Date Begin — Tab out to fire Wicket onchange, then wait for Ajax
    await type_into_field(page, SEL_DATE_BEGIN, begin)
    await page.keyboard.press("Tab")
    await page.wait_for_load_state("networkidle", timeout=PAGE_LOAD_TIMEOUT)
    await page.wait_for_timeout(AJAX_SETTLE_DELAY)

    # Confirm value stuck
    begin_val = await page.locator(SEL_DATE_BEGIN).input_value()
    log(f"  ✏  Company='{company_name}'  Begin='{begin_val}'")

    # Click Search
    log("  🖱  Clicking Search…")
    await page.locator(SEL_SUBMIT).click()
    await page.wait_for_load_state("networkidle", timeout=PAGE_LOAD_TIMEOUT)
    await page.wait_for_timeout(ACTION_DELAY)

    return await parse_results(page, company_name)


async def expect_either(page, locator_a, locator_b, timeout: int) -> None:
    """Wait until either locator_a or locator_b becomes visible."""
    await page.wait_for_function(
        """([selA, selB]) =>
            document.querySelector(selA)?.offsetParent !== null ||
            document.querySelector(selB)?.offsetParent !== null""",
        arg=["div.feedback", "table.tableResults"],
        timeout=timeout,
    )


async def parse_results(page, company_name: str) -> list[dict]:
    """
    Parse table.tableResults.  Column layout (0-indexed <td>):
      0,1 = sort icons  2=Case Number  3=Case Type  4=File Date
      5=Party/Company   6=Party Type   7=DOB        8=Case Status  9=Affiliation
    """
    # Fast path: site renders "No Records Found" in div.feedback immediately.
    # Check for it before waiting the full timeout for the results table.
    no_records = page.locator("div.feedback", has_text="No Records Found")
    results_table = page.locator("table.tableResults")
    try:
        await expect_either(page, no_records, results_table, timeout=PAGE_LOAD_TIMEOUT)
    except PWTimeout:
        log(f"  ⚠  Neither results nor 'No Records Found' appeared for '{company_name}'")
        return []

    if await no_records.is_visible():
        log(f"  📭  No records found for '{company_name}'")
        return []

    rows = await page.locator("table.tableResults tbody tr").all()
    results = []
    for row in rows:
        cells = await row.locator("td").all()
        if len(cells) < 9:
            continue
        results.append({
            "case_number":   (await cells[2].inner_text()).strip(),
            "party":         (await cells[5].inner_text()).strip(),
            "file_date":     (await cells[4].inner_text()).strip(),
        })

    log(f"  📋  {len(results)} result(s) found")
    for r in results:
        log(f"      {r['case_number']} | {r['file_date']} | {r['party']}")
    return results


async def save_xlsx(all_rows: list[dict], path: str) -> None:
    if not all_rows:
        log("\n⚠  No results to save.")
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Results"
    ws.append(["Case Number", "Party", "File Date"])
    for r in all_rows:
        ws.append([r["case_number"], r["party"], r["file_date"]])
    wb.save(path)
    log(f"\n💾  Saved {len(all_rows)} row(s) → {Path(path).resolve()}")


# ── Main ───────────────────────────────────────────────────────────────────────

async def run() -> None:
    begin = file_date_begin(DAYS_BACK)
    print("╔══════════════════════════════════════════════════╗")
    print("║  Alaska Court Records – Automation Starting      ║")
    print("╚══════════════════════════════════════════════════╝")
    log(f"File date begin : {begin}  ({DAYS_BACK} days ago)")
    log(f"Names           : {', '.join(NAMES)}")
    log(f"Output          : {OUTPUT_XLSX}\n")

    all_results: list[dict] = []

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=HEADLESS)
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            )
        )
        page = await context.new_page()

        # Home page
        log("🌐  Loading home page…")
        await page.goto(
            "https://records.courts.alaska.gov/eaccess/home.page.2",
            wait_until="networkidle",
            timeout=PAGE_LOAD_TIMEOUT,
        )
        await page.wait_for_timeout(5_000)

        # Click Search Cases
        log("🖱  Clicking 'Search Cases'…")
        await page.locator(SEL_SEARCH_CASES_BTN).first.click()
        await page.wait_for_load_state("networkidle", timeout=PAGE_LOAD_TIMEOUT)
        await page.wait_for_timeout(ACTION_DELAY)

        # Capture search page URL — used to return here between searches
        search_page_url = page.url
        log(f"   Search page URL: {search_page_url}")

        # Loop through names
        name_counts: dict[str, int] = {}
        for i, name in enumerate(NAMES):
            if i > 0:
                log(f"\n↩  Returning to search form…")
                await page.goto(search_page_url, wait_until="networkidle",
                                timeout=PAGE_LOAD_TIMEOUT)
                await page.wait_for_timeout(SEARCH_DELAY)

            await click_name_tab(page)

            try:
                rows = await search_one(page, name, begin)
                name_counts[name] = len(rows)
                all_results.extend(rows)
            except Exception as exc:
                name_counts[name] = 0
                log(f"  ❌  Error searching '{name}': {exc}")

        await browser.close()

    # Summary
    print("\n╔══════════════════════════════════════════════════╗")
    print("║  SUMMARY                                         ║")
    print("╚══════════════════════════════════════════════════╝")
    for name in NAMES:
        log(f"  '{name}': {name_counts.get(name, 0)} result(s)")

    await save_xlsx(all_results, OUTPUT_XLSX)
    log("\n✅  Done.")


if __name__ == "__main__":
    asyncio.run(run())