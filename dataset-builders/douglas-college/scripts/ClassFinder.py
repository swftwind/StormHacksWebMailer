import csv
import re
import sys
from datetime import datetime
from urllib.parse import urljoin
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

BASE = "https://www.douglascollege.ca"
COURSES_URL = f"{BASE}/courses"

# Process order you asked for:
# CATEGORY_ORDER = [
#     "Science & Technology",
#     "Commerce & Business Administration",
#     "Humanities & Social Sciences",
# ]
CATEGORY_ORDER = [
    "Health Sciences",
]

TARGET_GROUPS = set(CATEGORY_ORDER)

def now_term_label():
    y = datetime.now().year
    m = datetime.now().month
    if 1 <= m <= 4:
        return f"Winter {y}"
    if 5 <= m <= 8:
        return f"Summer {y}"
    return f"Fall {y}"

TERM_NOW = now_term_label()

def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def get_department_links(page):
    """Return ordered list of (group, dept_url, dept_name) for only target groups."""
    page.goto(COURSES_URL, wait_until="domcontentloaded")
    page.wait_for_selector(".view-content .grouped", timeout=20000)

    # Build map group -> links to preserve your order
    grouped = {g: [] for g in CATEGORY_ORDER}

    groups = page.locator(".view-content .grouped")
    for i in range(groups.count()):
        g = groups.nth(i)
        title = norm(g.locator("h2").inner_text()) if g.locator("h2").count() else ""
        if title not in TARGET_GROUPS:
            continue
        for a in g.locator('.views-row .field-content a[href^="/courses/"]').all():
            href = a.get_attribute("href") or ""
            grouped[title].append((title, urljoin(BASE, href), norm(a.inner_text())))

    # Flatten in the requested order, dedup by URL
    seen = set()
    ordered = []
    for cat in CATEGORY_ORDER:
        for tup in grouped.get(cat, []):
            if tup[1] not in seen:
                seen.add(tup[1])
                ordered.append(tup)
    return ordered

def get_course_links_from_department(page, dept_url):
    """Extract all /course/... links from the department's table(s)."""
    course_links = set()
    page.goto(dept_url, wait_until="domcontentloaded")
    # The table you pasted:
    rows = page.locator("tbody tr")
    if rows.count() == 0:
        # Some departments render differently; fallback to generic course links
        for a in page.locator('a[href^="/course/"]').all():
            href = a.get_attribute("href")
            if href and re.search(r"^/course/[a-z0-9\-]+$", href):
                course_links.add(urljoin(BASE, href))
        return sorted(course_links)

    # Preferred: second <td> has the <a> to the course page
    for i in range(rows.count()):
        r = rows.nth(i)
        a = r.locator('td.views-field-field-description a[href^="/course/"]').first
        if a.count():
            href = a.get_attribute("href")
            if href:
                course_links.add(urljoin(BASE, href))

    return sorted(course_links)

def read_course_code(page):
    """
    Pull course code from the facts grid (label 'Course code' -> value),
    fallback to first ABCD 1234-ish token.
    """
    # Try a tight selector: div.field with label 'Course code'
    try:
        label = page.locator("div.field:has(div.field__label:has-text('Course code'))").first
        if label.count():
            val = label.locator("div.field__item").first.inner_text()
            return norm(val)
    except Exception:
        pass

    # Fallback heuristic
    m = re.search(r"\b([A-Z]{3,5}\s*\d{3,4}[A-Z]?)\b", page.inner_text("body"))
    return norm(m.group(1)) if m else ""

def click_course_offerings(page):
    """Open the Course Offerings tab reliably."""
    try:
        page.get_by_role("tab", name=re.compile(r"Course Offerings", re.I)).click(timeout=7000)
        return
    except PWTimeout:
        pass
    for sel in [
        "a#tabset-0-tab-4",
        "a[role='tab']:has-text('Course Offerings')",
        "text=Course Offerings",
        "label:has-text('Course Offerings')",
    ]:
        try:
            page.locator(sel).first.click(timeout=5000)
            return
        except PWTimeout:
            continue
    # If nothing clickable, it might already be visible — do nothing.

def offerings_for_current_term(page, course_code):
    """
    Inside the Course Offerings pane, find the current term header and read instructor names.
    Returns list[(prof, course_code)].
    """
    out = []

    # Scope to the offerings pane
    offerings_section = page.locator("#tabset-0-section-4, section#tabset-0-section-4").first
    if not offerings_section.count():
        # Sometimes content is mounted under block id
        offerings_section = page.locator("#block-courseofferingentitiesblock").first
    if not offerings_section.count():
        return out

    # Ensure tables are loaded
    offerings_section.wait_for(state="visible", timeout=10000)

    # Find the header for the current term
    term_header = offerings_section.locator(
        f"h2:has-text('{TERM_NOW}'), h3:has-text('{TERM_NOW}')"
    ).first

    # If no explicit term header, fall back to searching any table in the offerings section
    tables_scope = offerings_section
    tables = []
    if term_header.count():
        # the table is usually the next sibling after the header
        tables = term_header.locator("xpath=following-sibling::div//table|following-sibling::table").all()
    if not tables:
        tables = offerings_section.locator("table").all()
    if not tables:
        return out

    def name_from_row(row):
        # Preferred: the two fields shown in your HTML
        first = row.locator(".field--name-instructor-first-name .field__item").first
        last = row.locator(".field--name-instructor-last-name .field__item").first
        if first.count() and last.count():
            return f"{norm(first.inner_text())} {norm(last.inner_text())}"
        # Fallback: use the Instructor <td> text
        # Find instructor cell by header text if thead exists
        tds = row.locator("td")
        # crude fallback: grab the 3rd cell
        if tds.count() >= 3:
            return norm(tds.nth(2).inner_text())
        return ""

    for table in tables:
        # Grab rows
        rows = table.locator("tbody tr")
        for i in range(rows.count()):
            r = rows.nth(i)
            prof = name_from_row(r)
            if prof and course_code:
                out.append((prof, course_code))

    return out

def scrape_course(page, course_url):
    """Open a course, open offerings, collect (prof, course)."""
    page.goto(course_url, wait_until="domcontentloaded")
    # Defensive: ignore listing-like pages
    if "/courses/" in page.url and "/course/" not in page.url:
        return []

    code = read_course_code(page)
    click_course_offerings(page)

    # Wait for a CRN table (or at least the offerings heading) to ensure content mounted
    try:
        page.wait_for_selector("#tabset-0-section-4 h2, #tabset-0-section-4 h3, #block-courseofferingentitiesblock h2", timeout=10000)
    except PWTimeout:
        pass

    pairs = offerings_for_current_term(page, code)
    return pairs

def main(headless=False):
    all_pairs = []
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=headless)
        context = browser.new_context(viewport={"width": 1400, "height": 900})
        page = context.new_page()

        departments = get_department_links(page)

        for group, dept_url, dept_name in departments:
            print(f"[{group}] {dept_name} -> {dept_url}")
            try:
                course_links = get_course_links_from_department(page, dept_url)
            except Exception as e:
                print(f"  ! Failed dept list: {e}")
                continue

            for c in course_links:
                try:
                    pairs = scrape_course(page, c)
                    if pairs:
                        all_pairs.extend(pairs)
                        print(f"    + {len(pairs)} rows from {c}")
                    else:
                        print(f"    - no current-term rows at {c}")
                except Exception as e:
                    print(f"    ! Failed course {c}: {e}")
                    continue

        browser.close()

    # Dedup & write
    dedup = []
    seen = set()
    for prof, course in all_pairs:
        key = (prof, course)
        if key not in seen:
            seen.add(key)
            dedup.append({"Prof Name": prof, "Course Number": course})

    with open("inputs.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["Prof Name", "Course Number"])
        w.writeheader()
        w.writerows(dedup)

    print(f"\nWrote {len(dedup)} rows for {TERM_NOW} -> inputs.csv")

if __name__ == "__main__":
    main(headless=("--headless" in sys.argv))
