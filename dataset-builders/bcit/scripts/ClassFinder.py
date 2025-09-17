# bcit_scraper.py
import asyncio
import csv
import re
import sys
from pathlib import Path
from typing import List, Tuple, Set

from playwright.async_api import async_playwright, TimeoutError as PWTimeout, expect

START_URL = "https://www.bcit.ca/study/"
OUTPUT_CSV = "../datasets/bcit_courses.csv"

PAUSE_SHORT = 50
PAUSE_MED = 100

# Normalize whitespace
def norm(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())

def is_placeholder_instructor(name: str) -> bool:
    t = (name or "").lower()
    return t in {"tba", "tbd", "(faculty) tba", "faculty tba", "instructor tba", "instructor tbd"}

async def dismiss_cookie_banner(page):
    try:
        for label in ["Accept", "I agree", "Allow all", "Accept all", "Got it"]:
            btn = page.get_by_role("button", name=label)
            if await btn.count():
                await btn.first.click(timeout=2000)
                await page.wait_for_timeout(300)
                break
    except Exception:
        pass

async def get_section_courses_links(page) -> List[str]:
    """From main study page, find each 'Courses' link."""
    links = []
    blocks = page.locator("div.content-block")
    count = await blocks.count()
    for i in range(count):
        block = blocks.nth(i)
        courses = block.locator("a", has_text="Courses")
        if await courses.count():
            href = await courses.first.get_attribute("href")
            if href and href.startswith("http"):
                links.append(href)
    return links

async def collect_course_links_on_courses_page(page) -> List[str]:
    links = []
    anchors = page.locator("a.ptscourseslist--link")
    count = await anchors.count()
    for i in range(count):
        href = await anchors.nth(i).get_attribute("href")
        if href and href.startswith("/courses/"):
            links.append("https://www.bcit.ca" + href)
    return links

async def scrape_instructors_from_course(page) -> List[Tuple[str, str]]:
    results = []

    # Grab course number from header
    try:
        header_text = await page.locator("main h1, header h1").first.text_content()
        course_num_match = re.search(r"\b([A-Z]{3,5}\s*\d{3,4})\b", header_text or "")
        course_num = norm(course_num_match.group(1)) if course_num_match else ""
    except Exception:
        course_num = ""

    sections = page.locator("div.sctn")
    count = await sections.count()
    for i in range(count):
        sctn = sections.nth(i)

        # Expand details if needed
        view_btn = sctn.locator("button.clicktoshow.course-section-details")
        if await view_btn.count():
            try:
                await view_btn.click()
                instr = sctn.locator(".sctn-instructor p")
                await expect(instr).to_be_visible(timeout=3000)
            except Exception:
                pass

        # Grab instructor
        instr = sctn.locator(".sctn-instructor p")
        if await instr.count():
            name = norm(await instr.first.text_content())
            if name and not is_placeholder_instructor(name):
                results.append((name, course_num))
    return results

async def run():
    out_path = Path(OUTPUT_CSV)
    seen: Set[Tuple[str, str]] = set()
    rows: List[Tuple[str, str]] = []

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(channel="msedge", headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        await page.goto(START_URL, wait_until="domcontentloaded")
        await dismiss_cookie_banner(page)
        await page.wait_for_timeout(PAUSE_MED)

        # Step 1: Collect section-level "Courses" links
        section_links = await get_section_courses_links(page)
        print(f"[INFO] Found {len(section_links)} section course pages.")

        # Step 2: Collect all course links
        all_course_links: List[str] = []
        for section_url in section_links:
            try:
                await page.goto(section_url, wait_until="domcontentloaded")
                await page.wait_for_timeout(PAUSE_SHORT)
                links = await collect_course_links_on_courses_page(page)
                all_course_links.extend(links)
            except Exception as e:
                print(f"[WARN] Could not scrape section {section_url}: {e}")

        all_course_links = sorted(set(all_course_links))
        print(f"[INFO] Found {len(all_course_links)} course pages.")

        # Step 3: Visit each course page
        for idx, url in enumerate(all_course_links, 1):
            try:
                await page.goto(url, wait_until="domcontentloaded")
                await page.wait_for_timeout(PAUSE_SHORT)

                pairs = await scrape_instructors_from_course(page)
                for name, course in pairs:
                    key = (name, course)
                    if key not in seen:
                        seen.add(key)
                        rows.append(key)

                if idx % 25 == 0:
                    print(f"[INFO] Processed {idx}/{len(all_course_links)} course pages...")
            except Exception as e:
                print(f"[WARN] Failed course {url}: {e}")

        await browser.close()

    # Step 4: Write to CSV
    rows.sort(key=lambda x: (x[1], x[0]))
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Prof Name", "Course Number"])
        for name, course in rows:
            w.writerow([name, course])

    print(f"[DONE] Wrote {len(rows)} rows to {out_path.resolve()}")

if __name__ == "__main__":
    try:
        asyncio.run(run())
    except KeyboardInterrupt:
        sys.exit(1)
