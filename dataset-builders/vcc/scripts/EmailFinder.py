import csv
import re
import sys
import time
from pathlib import Path
from typing import Dict, List, Tuple
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

DIR_URL = "https://www.vcc.ca/about/college-information/contact-us/employee-directory/"
IN_CSV  = "../datasets/vcc_classes_fall_2025.csv"        # Prof Name,Course Number
OUT_CSV = "../datasets/output.csv"                       # name,email,course

# -------------------- helpers --------------------

def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def exact_match(a: str, b: str) -> bool:
    return norm_space(a) == norm_space(b)

def wait_for_results(page, target_name: str, timeout_ms: int = 15000) -> None:
    try:
        page.wait_for_function(
            """() => {
                const tbl = document.querySelector('table.table-striped-gray');
                const body = tbl?.querySelector('tbody');
                const hasRows = !!(body && body.querySelector('tr'));
                const noRes = document.body.innerText.toLowerCase().includes('no result');
                return hasRows || noRes;
            }""",
            timeout=timeout_ms
        )
    except PWTimeout:
        pass

def extract_email_from_rows(rows_locator, target_name: str) -> str:
    count = rows_locator.count()
    for i in range(count):
        row = rows_locator.nth(i)
        first_td = row.locator("td").first
        if not first_td.count():
            continue

        html = first_td.inner_html()
        parts = re.split(r"<br\s*/?>", html, flags=re.I)
        name_html = parts[0] if parts else ""
        disp_name = norm_space(re.sub(r"<[^>]+>", "", name_html))

        if exact_match(disp_name, target_name):
            link = first_td.locator("a[href^='mailto:']").first
            if link.count():
                href = link.get_attribute("href") or ""
                email = href.replace("mailto:", "").strip()
                return email
            return ""
    return ""

def search_and_get_email(page, name: str) -> str:
    page.wait_for_selector("input#q-directory", timeout=20000)
    q = page.locator("input#q-directory")
    btn = page.locator("#btnSubmit")

    q.click()
    q.fill("")
    q.type(name, delay=20)
    time.sleep(0.1)
    btn.click()

    wait_for_results(page, name, timeout_ms=20000)
    time.sleep(2.5)

    table = page.locator("table.table-striped-gray")
    if not table.count():
        return ""

    rows = table.locator("tbody tr")
    return extract_email_from_rows(rows, name)

# -------------------- main --------------------

def main(headless=False, throttle_sec=0.25):
    rows: List[Tuple[str, str]] = []
    with open(IN_CSV, newline="", encoding="utf-8") as f:
        r = csv.DictReader(f)
        name_field = "Prof Name" if "Prof Name" in r.fieldnames else "Professor Name"
        for row in r:
            name = norm_space(row[name_field])
            course = norm_space(row["Course Number"])
            if name:
                rows.append((name, course))

    cache: Dict[str, str] = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=headless)
        page = browser.new_page(viewport={"width": 1400, "height": 900})
        page.goto(DIR_URL, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_selector("input#q-directory", timeout=20000)

        unique_names = sorted({n for n, _ in rows})
        total = len(unique_names)
        for idx, name in enumerate(unique_names, 1):
            try:
                time.sleep(throttle_sec)
                email = search_and_get_email(page, name)
            except Exception:
                email = ""
            cache[name] = email
            log_email = email if email else "(none)"
            print(f"[TRY  ] {idx}/{total} {name} -> {log_email}")

        browser.close()

    out_path = Path(OUT_CSV)
    with out_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["name", "email", "course"])
        for name, course in rows:
            email = cache.get(name, "")
            log_email = email if email else "(none)"
            print(f"[WRITE] {name},{log_email},{course}")
            w.writerow([name, email, course])

    print(f"\n[DONE ] Wrote {out_path} with {len(rows)} rows.")

if __name__ == "__main__":
    main(headless=("--headless" in sys.argv))
