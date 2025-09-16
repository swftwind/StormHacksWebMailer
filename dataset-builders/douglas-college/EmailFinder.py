import re
import time
import random
import pandas as pd
from pathlib import Path
import asyncio
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

FACULTY_DIR = "https://www.douglascollege.ca/faculty-directory"
INPUT_CSV = "input.csv"
OUTPUT_CSV = "output.csv"
HEADLESS = False  # set True once it's stable

def pause(a=0.6, b=1.2):
    time.sleep(random.uniform(a, b))

def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()

def split_name(full: str):
    full = norm(full)
    parts = full.split()
    if len(parts) >= 2:
        return parts[0], " ".join(parts[1:])
    if parts:
        return "", parts[0]
    return "", ""

def read_names(csv_path: str):
    df = pd.read_csv(csv_path)
    cols = {c.lower(): c for c in df.columns}
    names = []
    if "name" in cols:
        for n in df[cols["name"]].astype(str).tolist():
            f, l = split_name(n)
            names.append({"first": f, "last": l, "full": norm(n)})
    else:
        first_col = cols.get("first")
        last_col  = cols.get("last")
        if not last_col:
            raise ValueError("CSV must have either a 'Name' column or 'First' + 'Last'.")
        for _, row in df.iterrows():
            f = norm(str(row[first_col])) if first_col else ""
            l = norm(str(row[last_col]))
            names.append({"first": f, "last": l, "full": norm(f"{" " if f else ""}{f} {l}")})
    return names

async def fill_and_submit_on_faculty_dir(page, first, last):
    """On the faculty directory page, type the name and submit search.
    The site opens **results in a new tab**; we return that page."""
    await page.goto(FACULTY_DIR, wait_until="domcontentloaded", timeout=30000)
    pause()

    # Try several robust ways to find the search box and button
    # 1) Search input by role/placeholder/label
    filled = False
    search_text = (first + " " + last).strip() or last or first

    # Common patterns: a single search box for names, or separate First/Last inputs.
    # Try single search box first:
    for selector in [
        'input[placeholder*="Search" i]',
        'input[aria-label*="Search" i]',
        'input[name*="search" i]',
        'input[type="search"]',
        'input[type="text"]',
        '[role="search"] input',
    ]:
        try:
            loc = page.locator(selector).first
            if await loc.count() > 0:
                await loc.fill(search_text)
                filled = True
                break
        except Exception:
            pass

    # If not found, try two-input layout (First/Last)
    if not filled:
        try:
            # these are heuristics; adjust if needed after you watch the run
            first_box = page.get_by_label(re.compile(r"First\s*Name", re.I))
            last_box  = page.get_by_label(re.compile(r"Last\s*Name", re.I))
            if await first_box.count() > 0 and first:
                await first_box.fill(first)
                filled = True
            if await last_box.count() > 0 and last:
                await last_box.fill(last)
                filled = True
        except Exception:
            pass

    if not filled:
        # last resort: fill the first input on the page
        try:
            any_input = page.locator("input").first
            if await any_input.count() > 0:
                await any_input.fill(search_text)
                filled = True
        except Exception:
            pass

    # Click the Search button and wait for popup tab
    # Try common button texts
    button = None
    for text in ["Search", "Find", "Submit", "Go"]:
        try:
            candidate = page.get_by_role("button", name=re.compile(text, re.I))
            if await candidate.count() > 0:
                button = candidate.first
                break
        except Exception:
            pass
    if not button:
        # generic submit on forms: press Enter
        async with page.context.expect_page(timeout=7000) as new_page_info:
            await page.keyboard.press("Enter")
        new_page = await new_page_info.value
        await new_page.wait_for_load_state("domcontentloaded")
        return new_page

    async with page.context.expect_page(timeout=10000) as new_page_info:
        await button.click()
    new_page = await new_page_info.value
    await new_page.wait_for_load_state("domcontentloaded")
    return new_page

def clean_mailto(href: str) -> str:
    if not href:
        return ""
    href = href.strip()
    if href.lower().startswith("mailto:"):
        href = href[7:]
    return href.split("?")[0].strip()

async def extract_table_rows(page):
    rows = []
    # wait briefly for a table of results
    try:
        await page.wait_for_selector("table tr", timeout=8000)
    except Exception:
        return rows

    table = page.locator("table").first
    if await table.count() == 0:
        return rows

    trs = table.locator("tr")
    n = await trs.count()
    for i in range(1, n):  # skip header row
        tds = trs.nth(i).locator("td")
        if await tds.count() < 7:
            continue
        # collect text
        last  = norm(await tds.nth(0).inner_text())
        first = norm(await tds.nth(1).inner_text())
        email_href = ""
        try:
            a = tds.nth(6).locator("a").first
            if await a.count() > 0:
                email_href = (await a.get_attribute("href")) or ""
        except Exception:
            pass

        rows.append({
            "first": first,
            "last": last,
            "email_href": email_href
        })
    return rows

def choose_best(rows, first, last):
    tf, tl = norm(first).lower(), norm(last).lower()

    # 1) exact first+last
    for r in rows:
        if r["first"].lower() == tf and r["last"].lower() == tl:
            return r
    # 2) unique last match
    last_matches = [r for r in rows if r["last"].lower() == tl]
    if len(last_matches) == 1:
        return last_matches[0]
    # 3) heuristic on email username
    if last_matches and tf:
        initial = tf[:1]
        for r in last_matches:
            em = r.get("email_href","").lower()
            if em.startswith("mailto:"):
                user = em[7:].split("@")[0]
                if user.startswith(initial + tl) or user.startswith(tf + "." + tl):
                    return r
    # 4) only one total row
    if len(rows) == 1:
        return rows[0]
    return None

async def run():
    names = read_names(INPUT_CSV)
    out = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(channel="msedge", headless=HEADLESS)
        context = await browser.new_context()
        page = await context.new_page()

        for entry in names:
            first, last, full = entry["first"], entry["last"], entry["full"]
            pause()

            try:
                result_page = await fill_and_submit_on_faculty_dir(page, first, last)
            except Exception:
                out.append({"name": full, "email": ""})
                continue

            rows = await extract_table_rows(result_page)
            best = choose_best(rows, first, last)
            email = clean_mailto(best["email_href"]) if best else ""
            out.append({"name": full, "email": email})

            # close result tab to keep things tidy
            try:
                await result_page.close()
            except Exception:
                pass

        await browser.close()

    pd.DataFrame(out).to_csv(OUTPUT_CSV, index=False)
    print(f"Saved {len(out)} rows to {OUTPUT_CSV}")

if __name__ == "__main__":
    asyncio.run(run())
