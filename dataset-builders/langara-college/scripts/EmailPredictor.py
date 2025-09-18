# Predict Langara emails from names using pattern:
# first initial + last name + @langara.ca  (lowercase)
#
# Input CSV:  Prof Name,Course Number
# Output CSV: name,email,course
#
# Example:
#  "Lan, Gabrielle",ACCT 1045  ->  Gabrielle Lan,glan@langara.ca,ACCT 1045

import csv
import re
import sys
import unicodedata
from pathlib import Path
from typing import Tuple

IN_CSV  = "../datasets/langara_classes_fall_2025_clean.csv"  # Prof Name, Course Number
OUT_CSV = "../datasets/output.csv"          # name,email,course

# ---- helpers ----

TITLE_RE = re.compile(
    r"^\s*(dr\.?|prof\.?|mr\.?|mrs\.?|ms\.?|mx\.?)\s+", re.IGNORECASE
)

def strip_titles(s: str) -> str:
    """Remove common honorifics at the start."""
    s = s.strip()
    # repeatedly remove leading title tokens
    prev = None
    while prev != s:
        prev = s
        s = TITLE_RE.sub("", s, count=1)
    return s

def remove_diacritics(s: str) -> str:
    """Turn José → Jose, Zoë → Zoe."""
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", s)
        if not unicodedata.combining(ch)
    )

def clean_name_token(token: str) -> str:
    """Lowercase, remove diacritics and non-letters (keep letters only)."""
    token = remove_diacritics(token)
    token = token.lower()
    token = re.sub(r"[^a-z]", "", token)  # drop punctuation, spaces, digits
    return token

def parse_name(full: str) -> Tuple[str, str, str]:
    """
    Return (display_name, first, last).
    Accepts 'Last, First [Middle]' or 'First [Middle] Last'.
    """
    original = full.strip()
    s = strip_titles(original)

    # Multiple-prof rows? Skip gently (return empty parts)
    if any(sep in s for sep in [" / ", " & ", " and ", ";"]):
        return original, "", ""

    # Try 'Last, First' form
    if "," in s:
        parts = [p.strip() for p in s.split(",") if p.strip()]
        if len(parts) >= 2:
            last = parts[0]
            first_part = parts[1]
            first = first_part.split()[0] if first_part else ""
        else:
            # degenerate, fall back to space logic
            parts = s.split()
            first = parts[0] if parts else ""
            last  = parts[-1] if len(parts) > 1 else ""
    else:
        # 'First Middle Last'
        parts = s.split()
        first = parts[0] if parts else ""
        last  = parts[-1] if len(parts) > 1 else ""

    # Build nice display name "First Last"
    disp_first = first.strip()
    disp_last  = last.strip()
    display = " ".join(p for p in [disp_first, disp_last] if p)

    return display if display else original, first, last

def predict_langara_email(first: str, last: str) -> str:
    fi = clean_name_token(first[:1])  # first initial
    ln = clean_name_token(last)
    if not fi or not ln:
        return ""
    return f"{fi}{ln}@langara.ca"

# ---- main ----

def main(in_csv=IN_CSV, out_csv=OUT_CSV):
    rows_in = []
    with open(in_csv, newline="", encoding="utf-8") as f:
        r = csv.DictReader(f)
        # Resolve column names flexibly
        name_field = None
        course_field = None
        lower = [c.lower() for c in r.fieldnames]
        for i, c in enumerate(lower):
            if "name" in c or "prof" in c or "instructor" in c or "faculty" in c:
                name_field = r.fieldnames[i]
            if ("course" in c) or ("number" in c) or ("code" in c):
                course_field = r.fieldnames[i]
        if name_field is None:
            name_field = r.fieldnames[0]
        if course_field is None:
            course_field = r.fieldnames[1] if len(r.fieldnames) > 1 else r.fieldnames[0]

        for row in r:
            name = (row.get(name_field) or "").strip()
            course = (row.get(course_field) or "").strip()
            if name:
                rows_in.append((name, course))

    out_rows = []
    for prof, course in rows_in:
        display, first, last = parse_name(prof)
        email = predict_langara_email(first, last)
        out_rows.append({
            "name": display,
            "email": email,
            "course": course
        })

    Path(out_csv).parent.mkdir(parents=True, exist_ok=True)
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["name", "email", "course"])
        w.writeheader()
        w.writerows(out_rows)

    print(f"[DONE] Wrote {out_csv} with {len(out_rows)} rows.")

if __name__ == "__main__":
    # Allow `--in path` `--out path`
    args = sys.argv[1:]
    in_csv  = IN_CSV
    out_csv = OUT_CSV
    if "--in" in args:
        in_csv = args[args.index("--in")+1]
    if "--out" in args:
        out_csv = args[args.index("--out")+1]
    main(in_csv, out_csv)
