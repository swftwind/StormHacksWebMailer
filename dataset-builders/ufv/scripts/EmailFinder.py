import csv
import re
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

DIR_URL = "https://www.ufv.ca/directory/?showall=1&showall=1"
IN_CSV  = "../datasets/ufv_classes_fall_2025.csv"         # Prof Name,Course Number
OUT_CSV = "../datasets/outputs.csv"                       # name,email,course

def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def exact_match(a: str, b: str) -> bool:
    return norm_space(a) == norm_space(b)

def fetch_email_for_name(page, name: str) -> str:
    """Search the UFV directory for `name`. Return email if there is an EXACT match; else ''."""
    # Clear search box and enter name
    search = page.locator("#searchField input#search")
    search.click()
    search.fill("")
    search.type(name, delay=30)
    search.press("Enter")

    # Wait for search results
    try:
        page.wait_for_selector("#search-results .staff-card", timeout=8000)
    except PWTimeout:
        return ""

    cards = page.locator("#search-results .staff-card")
    count = cards.count()
    if count == 0:
        return ""

    # Iterate through cards, look for exact heading match
    for i in range(count):
        card = cards.nth(i)
        heading = card.locator(".card-heading").first
        if not heading.count():
            continue
        htext = norm_space(heading.inner_text())
        if exact_match(htext, name):
            email_link = card.locator(".card-email a[href^='mailto:']").first
            if email_link.count():
                email = email_link.get_attribute("href") or ""
                return email.replace("mailto:", "").strip()
            return ""  # exact match but no email

    return ""

def main(headless=False):
    # Read input CSV
    rows = []
    with open(IN_CSV, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            rows.append({"name": r["Prof Name"], "course": r["Course Number"]})

    cache = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=headless)
        page = browser.new_page(viewport={"width": 1400, "height": 900})
        page.goto(DIR_URL, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_selector("#searchField input#search", timeout=20000)

        for r in rows:
            name = r["name"]
            if name not in cache:
                try:
                    email = fetch_email_for_name(page, name)
                    cache[name] = email
                    if email:
                        print(f"[FOUND] {name}, {email}, {r['course']}")
                    else:
                        print(f"[MISS ] {name}, (no email), {r['course']}")
                except Exception as e:
                    cache[name] = ""
                    print(f"[ERROR] {name}, error={e}, {r['course']}")

        browser.close()

    # Write output
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["name", "email", "course"])
        w.writeheader()
        for r in rows:
            w.writerow({"name": r["name"], "email": cache.get(r["name"], ""), "course": r["course"]})

    print(f"Wrote {OUT_CSV} with {len(rows)} rows.")

if __name__ == "__main__":
    main(headless=("--headless" in sys.argv))
