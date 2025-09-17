import csv
import re
import sys
from playwright.sync_api import sync_playwright

URL = "https://www.ufv.ca/arfiles/includes/202509-timetable-with-changes.htm"
OUT_CSV = "../datasets/ufv_classes_fall_2025.csv"

# e.g. ABT 110, BUS 100, ENGL 052, FREN 460C
COURSE_RE = re.compile(r"^\s*([A-Z]{2,5})\s+(\d{3,4}[A-Z]?)\b")

# Instructor on section line after CRN + section code, when NOT wrapped in a span
# Examples:
#   "90950 AB1 Samantha Hannah          ABK 149 ..."
#   "90250 AB3 Melanie Opmeer           ABA 234 ..."
SECTION_NAME_RE = re.compile(
    r"""^\s*\d{5}\s+            # CRN
        [A-Z0-9#]{2,3}\s+       # section code (AB1, ON2, A#A, CH1, etc.)
        (?P<name>[A-Z][A-Za-z'.-]+(?:\s+[A-Z][A-Za-z'.-]+)+) # instructor name
        \s{2,}                   # gap before building/campus column
    """,
    re.VERBOSE,
)

def looks_like_name(s: str) -> bool:
    s = s.strip()
    # Two+ capitalized words, allow hyphens, apostrophes, dots
    return bool(re.fullmatch(r"[A-Z][A-Za-z'.-]+(?: [A-Z][A-Za-z'.-]+)+", s))

def main(headless=False):
    pairs = []
    seen = set()
    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=headless)
        page = browser.new_page()
        page.goto(URL, wait_until="domcontentloaded", timeout=120_000)

        # All relevant rows are in <pre class="style173">
        page.wait_for_selector("pre.style173", timeout=60_000)
        blocks = page.locator("pre.style173")
        count = blocks.count()

        current_course = None

        for i in range(count):
            pre = blocks.nth(i)
            txt = pre.inner_text().rstrip()

            # 1) Course header?
            m = COURSE_RE.match(txt)
            if m:
                current_course = f"{m.group(1)} {m.group(2)}"
                continue

            if not current_course:
                continue  # ignore anything before the first course header

            # 2) Try spans that UFV uses for names
            grabbed_any = False
            for span_sel in ("span.style14", "span.style273"):
                spans = pre.locator(span_sel)
                for j in range(spans.count()):
                    name = spans.nth(j).inner_text().strip()
                    if looks_like_name(name):
                        key = (name, current_course)
                        if key not in seen:
                            seen.add(key)
                            pairs.append({"Prof Name": name, "Course Number": current_course})
                        grabbed_any = True

            # 3) If no valid span-based name found, try plain text pattern on the line
            if not grabbed_any:
                m2 = SECTION_NAME_RE.match(txt)
                if m2:
                    name = m2.group("name").strip()
                    if looks_like_name(name):
                        key = (name, current_course)
                        if key not in seen:
                            seen.add(key)
                            pairs.append({"Prof Name": name, "Course Number": current_course})

        browser.close()

    # Write CSV
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["Prof Name", "Course Number"])
        writer.writeheader()
        writer.writerows(pairs)

    print(f"Wrote {len(pairs)} rows to {OUT_CSV}")

if __name__ == "__main__":
    main(headless=("--headless" in sys.argv))
