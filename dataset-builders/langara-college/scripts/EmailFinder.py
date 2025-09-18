import csv
import re
import sys
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

START_URL = "https://www.langara.ca/contact/departments"
OUT_CSV   = "../datasets/langara_faculty_emails.csv"

FACULTY_LINK_TEXTS = [
    r"\bfaculty\b",
    r"\bfaculty\s*&\s*staff\b",
    r"\bfaculty\s*/\s*staff\b",
    r"\bstaff\b",
    r"\binstructors?\b",
    r"\bpeople\b",
    r"\bour\s+people\b",
    r"\bdepartment\s+contacts?\b",
    r"\bcontacts?\b",
    r"\bcontact\s+us\b",
    r"\bteam\b",
    r"\bwho\s+we\s+are\b",
]

def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def clean_email(raw: str) -> str:
    if not raw:
        return ""
    email = raw.strip()
    email = email.replace("mailto:", "").replace("%20", "").strip()
    email = email.replace(" [at] ", "@").replace("[at]", "@")
    email = re.sub(r"\s*@\s*", "@", email)
    return email

def looks_like_email(s: str) -> bool:
    return bool(re.search(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", s or ""))

def name_from_email(email: str) -> str:
    """
    Very gentle fallback: 'jane.doe@langara.ca' -> 'Jane Doe'.
    Handles hyphens/underscores too. Only used if we couldn't find a nearby name.
    """
    try:
        local = email.split("@", 1)[0]
    except Exception:
        return ""
    if not local:
        return ""
    parts = re.split(r"[._\-]+", local)
    parts = [p for p in parts if p and not p.isdigit()]
    if not parts:
        return ""
    return " ".join(w.capitalize() for w in parts)

def absolutize(base: str, href: str) -> str:
    try:
        return urljoin(base, href)
    except Exception:
        return href

def same_site_or_students(u: str) -> bool:
    try:
        host = urlparse(u).netloc.lower()
    except Exception:
        return False
    return ("langara.ca" in host)

def log(msg: str) -> None:
    print(msg, flush=True)

def get_department_links(page) -> list[str]:
    links = set()
    anchors = page.locator("a.icon-link")
    n = anchors.count()
    for i in range(n):
        href = anchors.nth(i).get_attribute("href") or ""
        href = href.strip()
        if not href:
            continue
        url = absolutize(page.url, href)
        if same_site_or_students(url):
            links.add(url)
    if not links:
        anchors = page.locator("main a[href]")
        n = anchors.count()
        for i in range(min(n, 1000)):
            href = anchors.nth(i).get_attribute("href") or ""
            url = absolutize(page.url, href)
            if same_site_or_students(url):
                links.add(url)
    return sorted(links)

def click_faculty_nav_if_present(page) -> bool:
    for pat in FACULTY_LINK_TEXTS:
        try:
            locator = page.locator("nav a", has_text=re.compile(pat, re.I))
            if locator.count():
                locator.first.click(timeout=3000)
                page.wait_for_load_state("domcontentloaded", timeout=15000)
                return True
        except Exception:
            pass
        try:
            locator = page.locator("a", has_text=re.compile(pat, re.I))
            if locator.count():
                locator.first.click(timeout=3000)
                page.wait_for_load_state("domcontentloaded", timeout=15000)
                return True
        except Exception:
            pass
    return False

def extract_name_near_anchor_handle(a_handle) -> str:
    """
    Robust name finder:
    - If inside a table: walk to containing <tr>, then scan previous <tr> siblings for strong/b/h*.
    - Otherwise: search previous siblings in the same container; if none, walk up parents and
      scan their previous siblings for strong/b/h*.
    - If still missing: search any strong/b/h* within the same container (fallback, may catch 'above').
    Returns normalized text or ''.
    """
    js = """
    (node) => {
      const isHeading = (el) => el && el.matches && el.matches('strong,b,h1,h2,h3,h4,h5,h6');

      const textOf = (el) => (el && el.textContent) ? el.textContent.trim() : '';

      // 1) If in a table row, look in previous rows for a heading/strong/b (like your sample)
      const tr = node.closest('tr');
      if (tr) {
        let prev = tr.previousElementSibling;
        while (prev) {
          // direct heading in the row:
          let hit = prev.querySelector('strong,b,h1,h2,h3,h4,h5,h6');
          if (hit && textOf(hit)) return textOf(hit);
          // name might be inside the first cell:
          const firstCell = prev.querySelector('td,th');
          if (firstCell) {
            let hit2 = firstCell.querySelector('strong,b,h1,h2,h3,h4,h5,h6');
            if (hit2 && textOf(hit2)) return textOf(hit2);
            // sometimes just the first line is the name:
            const raw = textOf(firstCell);
            if (raw) {
              const firstLine = raw.split('\\n')[0].trim();
              if (firstLine && !firstLine.includes('@')) return firstLine;
            }
          }
          prev = prev.previousElementSibling;
        }
      }

      // 2) General case: search previous siblings within same container, walking up if needed
      const container = node.closest('td, li, article, section, div, tbody, table, main, body') || node.parentElement;
      let el = node;
      while (el && el !== container) {
        el = el.parentElement;
      }
      // Now el == container (or null). Start from the anchor and scan previous siblings up the tree
      let cursor = node;
      while (cursor && cursor !== container) {
        let prev = cursor.previousElementSibling;
        while (prev) {
          if (isHeading(prev) && textOf(prev)) return textOf(prev);
          const within = prev.querySelector && prev.querySelector('strong,b,h1,h2,h3,h4,h5,h6');
          if (within && textOf(within)) return textOf(within);
          prev = prev.previousElementSibling;
        }
        cursor = cursor.parentElement;
      }

      // 3) As a very last try, grab the first heading/strong in the container
      if (container) {
        const any = container.querySelector('strong,b,h1,h2,h3,h4,h5,h6');
        if (any && textOf(any)) return textOf(any);
      }

      return '';
    }
    """
    try:
        name = a_handle.evaluate(js)
        return norm_space(name)
    except Exception:
        return ""

def extract_people_from_page(page) -> list[tuple[str, str]]:
    out = []
    anchors = page.locator("a[href^='mailto:'], a.spamspan[href^='mailto:']")
    n = anchors.count()
    for i in range(n):
        a = anchors.nth(i)
        href = a.get_attribute("href") or ""
        email = clean_email(href)
        if not looks_like_email(email):
            text = a.inner_text().strip()
            if looks_like_email(text):
                email = text
        if not looks_like_email(email):
            continue

        # Find name by walking DOM from the anchor handle
        try:
            a_handle = a.element_handle()
        except Exception:
            a_handle = None

        name = ""
        if a_handle:
            name = extract_name_near_anchor_handle(a_handle)

        # If empty, and anchor text itself looks like a name, take it
        if not name:
            txt = a.inner_text().strip()
            if re.search(r"[A-Za-z]+\s+[A-Za-z]+", txt) and not looks_like_email(txt):
                name = norm_space(txt)

        # Final fallback: derive from email (only if nothing else found)
        if not name:
            name = name_from_email(email)

        # Clean up "Email Address:" type prefixes if they slipped in
        name = re.sub(r"(?i)^email\s*address:?\s*", "", name).strip()

        out.append((name, email))
    return out

def main(headless=False, throttle=0.4):
    out_rows = []
    seen = set()

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=headless)
        page = browser.new_page(viewport={"width": 1400, "height": 900})

        log(f"[OPEN ] {START_URL}")
        page.goto(START_URL, wait_until="domcontentloaded", timeout=60000)

        dept_links = get_department_links(page)
        log(f"[FOUND] {len(dept_links)} department links")

        for idx, dept_url in enumerate(dept_links, 1):
            try:
                log(f"\n[DEPT ] {idx}/{len(dept_links)} {dept_url}")
                page.goto(dept_url, wait_until="domcontentloaded", timeout=60000)
                time.sleep(throttle)

                clicked = click_faculty_nav_if_present(page)
                if clicked:
                    log("[NAV  ] Faculty/People page opened")
                    time.sleep(throttle)

                people = extract_people_from_page(page)
                if not people:
                    log("[MISS ] No mailto links found on this page")
                else:
                    for (name, email) in people:
                        key = (name, email.lower())
                        if key not in seen:
                            seen.add(key)
                            out_rows.append({"name": name, "email": email})
                            log(f"[GRAB ] {name or '(no name)'} -> {email}")

            except Exception as e:
                log(f"[ERROR] {dept_url} :: {e}")

        browser.close()

    out_path = Path(OUT_CSV)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["name", "email"])
        w.writeheader()
        for r in out_rows:
            w.writerow(r)

    log(f"\n[DONE ] Wrote {out_path} with {len(out_rows)} rows.")

if __name__ == "__main__":
    main(headless=("--headless" in sys.argv))
