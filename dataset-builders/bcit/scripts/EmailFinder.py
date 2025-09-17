# Notes:
# - Uses Playwright with Edge channel.
# - Input CSV is expected to have headers: "Prof Name,Course Number"
# - Output CSV: "name,email,course"

import asyncio
import csv
import re
import sys
import argparse
from pathlib import Path
from typing import List, Tuple, Optional

from playwright.async_api import async_playwright, TimeoutError as PWTimeout

SEARCH_URL = "https://search.bcit.ca/s/search.html?collection=bcit~sp-search&profile=_default"

# polite pauses (ms)
PAUSE_SHORT = 200
PAUSE_MED = 400
PAUSE_LONG = 800

def norm_ws(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def normalize_name_for_match(name: str) -> str:
    # Keep letters, digits, spaces, and some common punctuation in names
    s = norm_ws(name)
    s = re.sub(r"[^A-Za-z0-9\s\-\.'`]", "", s)
    return s.lower()

def name_tokens(name: str) -> List[str]:
    return [t for t in re.split(r"[\s\-]+", normalize_name_for_match(name)) if t]

def match_score(query_name: str, candidate_name: str) -> float:
    """Simple token-overlap score to pick the best 'People' card."""
    q = set(name_tokens(query_name))
    c = set(name_tokens(candidate_name))
    if not q or not c:
        return 0.0
    overlap = len(q & c)
    return overlap / max(len(q), 1)

async def ensure_people_tab(page) -> None:
    """
    On the BCIT search results page, click the 'People' tab if present.
    This dramatically improves precision.
    """
    try:
        # The tab is a link that contains text 'People' and often includes a results count
        people_tab = page.locator("a", has_text=re.compile(r"^\s*People\s*", re.I))
        if await people_tab.count():
            await people_tab.first.click()
            await page.wait_for_timeout(PAUSE_MED)
    except Exception:
        pass

async def search_person_and_get_email(page, full_name: str) -> Optional[str]:
    """
    Type the name into the big search bar, submit, switch to People tab,
    scan person cards, and extract the best-matching email.
    """
    # Focus the search input (id='query'); it exists on landing & results pages
    try:
        query = page.locator("#query")
        await query.wait_for(timeout=5000)
        await query.fill("")                       # clear any previous term
        await page.wait_for_timeout(PAUSE_SHORT)
        await query.type(full_name, delay=25)
        await query.press("Enter")
    except PWTimeout:
        # If the search box isn't visible for some reason, navigate directly to the base URL
        await page.goto(SEARCH_URL, wait_until="domcontentloaded")
        await page.wait_for_timeout(PAUSE_MED)
        query = page.locator("#query")
        await query.fill(full_name)
        await query.press("Enter")

    # Results load
    await page.wait_for_load_state("domcontentloaded")
    await page.wait_for_timeout(PAUSE_MED)

    # Prefer the People tab
    await ensure_people_tab(page)

    # Wait a bit for people cards, but don't hang forever
    try:
        await page.wait_for_selector(".bcit-people-card", timeout=3000)
    except PWTimeout:
        # No person cards shown
        return None

    cards = page.locator(".bcit-people-card")
    count = await cards.count()
    if count == 0:
        return None

    best_email: Optional[str] = None
    best_score = -1.0

    for i in range(count):
        card = cards.nth(i)
        # Extract visible full name from the card head
        name_el = card.locator(".bcit-people-card__name")
        email_el = card.locator(".bcit-people-card__email a")

        try:
            name_text = norm_ws(await name_el.inner_text())
        except Exception:
            name_text = ""

        # Email may be missing for some results
        email_text = None
        if await email_el.count():
            try:
                email_text = norm_ws(await email_el.first.inner_text())
            except Exception:
                email_text = None

        if not email_text:
            continue

        score = match_score(full_name, name_text)
        if score > best_score:
            best_score = score
            best_email = email_text

    return best_email

async def run(input_csv: Path, output_csv: Path, headful: bool = False):
    # Load unique (name, course) pairs from input
    pairs: List[Tuple[str, str]] = []
    with input_csv.open("r", encoding="utf-8", newline="") as f:
        r = csv.DictReader(f)
        # Accept either "Prof Name" or "name" for robustness; same for "Course Number" or "course"
        for row in r:
            name = row.get("Prof Name") or row.get("name") or ""
            course = row.get("Course Number") or row.get("course") or ""
            name = norm_ws(name)
            course = norm_ws(course)
            if name:
                pairs.append((name, course))

    # Dedup pairs
    seen = set()
    uniq_pairs: List[Tuple[str, str]] = []
    for p in pairs:
        if p not in seen:
            seen.add(p)
            uniq_pairs.append(p)
    pairs = uniq_pairs

    results: List[Tuple[str, str, str]] = []  # (name, email, course)

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(channel="msedge", headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        # Open search landing once
        await page.goto(SEARCH_URL, wait_until="domcontentloaded")
        await page.wait_for_timeout(PAUSE_MED)

        for idx, (name, course) in enumerate(pairs, 1):
            try:
                email = await search_person_and_get_email(page, name)
                results.append((name, email or "", course))
                # Light pause to be courteous
                await page.wait_for_timeout(PAUSE_SHORT)
                if idx % 25 == 0:
                    print(f"[INFO] Processed {idx}/{len(pairs)} names...")
            except Exception as e:
                print(f"[WARN] lookup failed for '{name}': {e}")
                results.append((name, "", course))
                # Try to recover page state if something weird happened
                try:
                    await page.goto(SEARCH_URL, wait_until="domcontentloaded")
                    await page.wait_for_timeout(PAUSE_SHORT)
                except Exception:
                    pass

        await browser.close()

    # Write output
    with output_csv.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["name", "email", "course"])
        for name, email, course in results:
            w.writerow([name, email, course])

    print(f"[DONE] Wrote {len(results)} rows to {output_csv.resolve()}")

def parse_args():
    ap = argparse.ArgumentParser(description="BCIT email lookup from bcit_courses.csv")
    ap.add_argument("--in", dest="input_csv", default="bcit_courses.csv", help="Input CSV path")
    ap.add_argument("--out", dest="output_csv", default="output.csv", help="Output CSV path")
    ap.add_argument("--headful", action="store_true", help="Run with browser window (Edge)")
    return ap.parse_args()

if __name__ == "__main__":
    args = parse_args()
    in_path = Path(args.input_csv)
    out_path = Path(args.output_csv)
    try:
        asyncio.run(run(in_path, out_path, headful=args.headful))
    except KeyboardInterrupt:
        sys.exit(1)
