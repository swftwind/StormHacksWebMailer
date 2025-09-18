# Parse the uploaded Langara HTML offline and extract (Professor, Course Number)
import re, html, pandas as pd
from pathlib import Path

in_path = Path("../datasets/Course Search.html")
out_path = Path("../datasets/langara_classes_fall_2025.csv")

# Helpers
def strip_tags(s: str) -> str:
    s = re.sub(r"<[^>]+>", "", s, flags=re.S)
    s = s.replace("&nbsp;", " ")
    return html.unescape(s).strip()

def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()

def looks_like_subj(x: str) -> bool:
    return bool(re.fullmatch(r"[A-Z]{2,5}", x))

def looks_like_cnum(x: str) -> bool:
    return bool(re.fullmatch(r"\d{3,4}[A-Z]?", x))

def looks_like_name(x: str) -> bool:
    # Allow letters, hyphens, apostrophes, periods and spaces; avoid tokens like WWW or dashes
    if x.upper() in {"WWW", "TBA", "TBD"}: 
        return False
    if re.fullmatch(r"[-–—]+", x):
        return False
    # Must have at least one space (first + last) or comma
    return bool(re.search(r"[A-Za-z]", x)) and (" " in x or "," in x)

# Read file
html_text = in_path.read_text(encoding="utf-8", errors="ignore")

# Extract table rows
rows_html = re.findall(r"<tr[^>]*>(.*?)</tr>", html_text, flags=re.S|re.I)

records = []
for tr in rows_html:
    tds = re.findall(r"<td[^>]*>(.*?)</td>", tr, flags=re.S|re.I)
    if not tds:
        continue
    # Clean texts
    tvals = [norm_space(strip_tags(td)) for td in tds]
    # Find subject + course number
    subj = None
    cnum = None
    for v in tvals:
        if subj is None and looks_like_subj(v):
            subj = v
        elif subj and cnum is None and looks_like_cnum(v):
            cnum = v
            break
    if not subj or not cnum:
        continue
    # Instructor: choose the last plausible name-like cell
    instr = ""
    for v in reversed(tvals):
        if looks_like_name(v):
            instr = v
            break
    if not instr:
        continue
    records.append((instr, f"{subj} {cnum}"))

# Deduplicate
records = sorted(set(records))

df = pd.DataFrame(records, columns=["Prof Name", "Course Number"])

# Save CSV
df.to_csv(out_path, index=False)

out_path.as_posix()
