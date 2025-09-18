import csv
import re
import sys
import time
from pathlib import Path
from typing import List, Tuple, Set, Optional
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

BASE = "https://www.vcc.ca"
INDEX_URL = f"{BASE}/courses/"
OUT_CSV = "../datasets/vcc_classes_fall_2025.csv"  # Prof Name,Course Number

# --- helpers ---------------------------------------------------------------

def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def extract_course_code_from_text(text: str) -> str:
    """
    Link text on the index page typically ends with "(SUBJ NNNN)".
    Example: "ABE Computer Studies-Adv Level (COMP 0863)"
    """
    text = norm_space(text)
    m = re.search(r"\(([A-Z]{2,}\s*\d{3,4}[A-Z]?)\)\s*$", text)
    return m.group(1).replace("  ", " ") if m else ""

def get_all_courses_from_index(page) -> List[Tuple[str, str]]:
    """
    Return list of (course_url, course_code) from the index.
    We scan all <li class="ln-*> a[href^='/courses/']</a>.
    """
    items = []
    links = page.locator("li[class^='ln-'] a[href^='/courses/']")
    n = links.count()
    seen_urls = set()
    for i in range(n):
        a = links.nth(i)
        href = a.get_attribute("href") or ""
        text = a.inner_text()
        code = extract_course_code_from_text(text)
        url = href if href.startswith("http") else f"{BASE}{href}"
        if url in seen_urls:
            continue
        seen_urls.add(url)
        items.append((url, code))
    return items

def wait_for_schedule_or_nosched(page, timeout=8000) -> None:
    """
    For a course page, wait for either:
      - 'No schedule is currently available...' block, or
      - A 'Schedule' header block present ('.hide-full h2' contains 'Schedule')
    We don't assert which one; just ensure one of them appears to avoid racing.
    """
    try:
        page.wait_for_selector(
            "div.col-12:has(i.fa-calendar-xmark), .hide-full h2",
            timeout=timeout
        )
    except PWTimeout:
        # Not fatal — some pages might be fast/slow; continue to parse heuristically
        pass

def has_schedule_header(page) -> bool:
    """
    True if '.hide-full h2' contains the word 'Schedule' (case-insensitive).
    """
    headers = page.locator(".hide-full h2")
    n = headers.count()
    for i in range(n):
        txt = norm_space(headers.nth(i).inner_text()).lower()
        if "schedule" in txt:
            return True
    return False

def find_schedule_tables(page) -> List:
    """
    Find schedule tables on the page. We prefer tables near/after the 'Schedule' header,
    but as a fallback, return any 'table.cr-schedule-table' on the page.
    """
    tables = []
    # First, try to find tables that are within the same main content as the schedule header.
    if has_schedule_header(page):
        # If the page is structured, these tables will be globally under that section anyway.
        candidates = page.locator("table.cr-schedule-table")
        for i in range(candidates.count()):
            tables.append(candidates.nth(i))
    else:
        # Fallback: any schedule-like table
        candidates = page.locator("table.cr-schedule-table")
        for i in range(candidates.count()):
            tables.append(candidates.nth(i))
    return tables

def parse_course_offerings(page) -> List[str]:
    """
    Extract instructor names from any schedule table(s) if present.
    If 'No schedule...' block is visible and no tables, return [].
    """
    # Give the DOM a moment to settle (lazy bits)
    time.sleep(0.15)
    wait_for_schedule_or_nosched(page, timeout=6000)

    # If explicit "No schedule..." message is visible and no tables, we bail
    no_sched = page.locator("div.col-12:has(i.fa-calendar-xmark)")
    any_no_sched = no_sched.count() and no_sched.first.is_visible()

    tables = find_schedule_tables(page)

    if not tables:
        # No tables — if the nosched message is there, treat as no schedule.
        return [] if any_no_sched else []

    instructors: List[str] = []
    for table in tables:
        rows = table.locator("tbody tr")
        rc = rows.count()
        for i in range(rc):
            row = rows.nth(i)
            cell = row.locator("td[data-th='Instructor']")
            if not cell.count():
                continue
            # Instructor is typically inside <span class="cr-sched-instructor">
            name = ""
            span = cell.locator(".cr-sched-instructor").first
            if span.count():
                name = span.inner_text()
            else:
                name = cell.inner_text()
            name = norm_space(name)
            if name and name.lower() != "tba":
                instructors.append(name)

    return instructors

# --- main ------------------------------------------------------------------

def main(headless=False, throttle_sec=0.35):
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=headless)
        page = browser.new_page(viewport={"width": 1400, "height": 900})

        # 1) Open index
        print(f"[SCAN] Opening {INDEX_URL}")
        page.goto(INDEX_URL, wait_until="domcontentloaded", timeout=60000)

        # Ensure lazy lists render
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(1.0)

        # 2) Collect all course links + codes
        courses = get_all_courses_from_index(page)
        print(f"[SCAN] Found {len(courses)} course links on index.")

        # 3) Visit each course page and extract instructors
        pairs: Set[Tuple[str, str]] = set()  # (Prof Name, Course Number)
        total = len(courses)
        for idx, (url, code) in enumerate(courses, 1):
            print(f"[PARSE] {idx}/{total} {url}  code='{code or 'UNKNOWN'}'")
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=60000)
            except PWTimeout:
                print(f"[WARN ] Timeout loading {url}")
                continue

            # Be gentle; allow dynamic content to load
            time.sleep(throttle_sec)

            try:
                instructors = parse_course_offerings(page)
            except Exception as e:
                print(f"[WARN ] Error parsing {url}: {e}")
                instructors = []

            if instructors and code:
                for name in instructors:
                    pairs.add((name, code))
                print(f"[FOUND] {len(instructors)} instructor(s) for {code}: {', '.join(instructors)}")
            else:
                # If we saw a Schedule header but no names, call it 'no instructors'
                reason = "no schedule"
                if has_schedule_header(page):
                    reason = "schedule header, no rows"
                elif code == "":
                    reason = "no code"
                print(f"[SKIP ] {url} ({reason})")

        browser.close()

    # 4) Write CSV
    out_path = Path(OUT_CSV)
    with out_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Prof Name", "Course Number"])
        for name, code in sorted(pairs, key=lambda t: (t[1], t[0])):
            w.writerow([name, code])

    print(f"\n[DONE ] Wrote {out_path} with {len(pairs)} unique (Prof, Course) pairs.")

if __name__ == "__main__":
    main(headless=("--headless" in sys.argv))
