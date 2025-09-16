import re
import time
import random
import pandas as pd
from pathlib import Path
import asyncio
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

FACULTY_DIR = "https://www.douglascollege.ca/faculty-directory"
INPUT_CSV = "inputs.csv"
OUTPUT_CSV = "output.csv"
HEADLESS = False  # set True once it's stable
SKIP_TBA = False  # set False if you want rows like "(Faculty) TBA" included with blank emails

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
    # normalize headers -> original name
    headers = {c.strip().lower(): c for c in df.columns}

    # known aliases
    name_keys   = ["name", "prof name", "professor", "professor name",
                   "instructor", "instructor name", "faculty", "faculty name"]
    course_keys = ["course", "course number", "course code", "course id", "course title"]

    # find actual columns
    name_col = next((headers[k] for k in name_keys   if k in headers), None)
    crs_col  = next((headers[k] for k in course_keys if k in headers), None)

    if not name_col or not crs_col:
        raise ValueError(
            f"CSV needs a name-like column (e.g., 'Prof Name') and a course-like column "
            f"(e.g., 'Course Number'). Found: {list(df.columns)}"
        )

    names = []
    for _, row in df.iterrows():
        full   = norm(str(row[name_col]))
        course = norm(str(row[crs_col]))

        # skip empty or TBA names if desired
        if not full:
            continue
        if SKIP_TBA and re.search(r"\bTBA\b", full, re.I):
            continue

        f, l = split_name(full)
        names.append({"first": f, "last": l, "full": full, "course": course})

    return names

async def fill_and_submit_on_faculty_dir(page, first, last):
    """
    On the faculty directory page:
      - fill the single 'Enter Faculty name' box (or fallback to first input)
      - click the Submit button (robust selectors)
      - capture the results page (popup OR same tab)
    Returns: the results Page object.
    """
    await page.goto(FACULTY_DIR, wait_until="domcontentloaded", timeout=30000)

    # 1) Fill the search box
    search_text = (f"{first} {last}".strip() or last or first or "").strip()
    filled = False

    # Try by label text shown on the page
    try:
        box = page.get_by_label(re.compile(r"Enter\s*Faculty\s*name", re.I))
        if await box.count() > 0:
            await box.first.fill(search_text)
            filled = True
    except Exception:
        pass

    # Fallbacks: placeholder/name/first input
    if not filled:
        for sel in [
            'input[placeholder*="Enter Faculty name" i]',
            'input[aria-label*="Enter Faculty name" i]',
            'input[name*="name" i]',
            'input[type="search"]',
            'input[type="text"]',
        ]:
            loc = page.locator(sel).first
            if await loc.count() > 0:
                await loc.fill(search_text)
                filled = True
                break

    if not filled:
        any_input = page.locator("input").first
        if await any_input.count() > 0:
            await any_input.fill(search_text)
            filled = True

    # 2) Click the Submit button and capture the results
    # The site opens a new tab; we’ll expect a 'page' event. If not, check same-tab.
    button_selectors = [
        "button:has-text('Submit')",
        "input[type=submit][value='Submit']",
        "input[type=submit][value*='Submit' i]",
        "input[value='Submit']",
        "input[value*='Submit' i]"
    ]

    # Try popup path first
    for sel in button_selectors:
        btn = page.locator(sel).first
        if await btn.count() > 0:
            try:
                async with page.context.expect_page(timeout=7000) as new_page_info:
                    await btn.scroll_into_view_if_needed()
                    await btn.click()
                new_page = await new_page_info.value
                await new_page.wait_for_load_state("domcontentloaded")
                return new_page
            except PWTimeout:
                # Button clicked but no popup detected; maybe same-tab navigation
                try:
                    await page.wait_for_load_state("load", timeout=3000)
                    # If same tab, just use current page
                    return page
                except Exception:
                    pass
            except Exception:
                # Try the next selector
                pass

    # 3) Fallback: press Enter and expect popup
    try:
        async with page.context.expect_page(timeout=7000) as new_page_info:
            await page.keyboard.press("Enter")
        new_page = await new_page_info.value
        await new_page.wait_for_load_state("domcontentloaded")
        return new_page
    except PWTimeout:
        # Maybe same-tab
        try:
            await page.wait_for_load_state("load", timeout=3000)
            return page
        except Exception:
            pass

    # 4) Last resort: submit the first form via JS, then check popup or same-tab
    try:
        async with page.context.expect_page(timeout=7000) as new_page_info:
            await page.evaluate("""() => {
                const f = document.querySelector('form'); 
                if (f) f.submit();
            }""")
        new_page = await new_page_info.value
        await new_page.wait_for_load_state("domcontentloaded")
        return new_page
    except PWTimeout:
        # Same-tab again
        try:
            await page.wait_for_load_state("load", timeout=4000)
            return page
        except Exception:
            pass

    # If we reach here, nothing opened or navigated; return current page so caller continues gracefully
    return page

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
            out.append({"name": full, "email": email, "course": entry["course"]})

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
