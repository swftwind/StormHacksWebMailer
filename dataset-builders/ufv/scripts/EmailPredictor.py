import csv
import re
import unicodedata
from pathlib import Path

IN_CSV  = "../datasets/ufv_classes_fall_2025.csv"   # Prof Name,Course Number
OUT_CSV = "../datasets/outputs.csv"                 # name,email,course

# Choose how the local-part should be cased:
# - "title": Christine.Nehring@ufv.ca
# - "lower": christine.nehring@ufv.ca
LOCAL_PART_CASE = "title"

# Common prefixes/suffixes to ignore
TITLES = {
    "mr", "mrs", "ms", "miss", "mx", "dr", "prof", "professor",
    "sir", "madam", "madame", "rev", "reverend", "fr", "father",
    "sr", "sister", "dean", "provost"
}
SUFFIXES = {"jr", "sr", "ii", "iii", "iv", "phd", "md", "dvm", "mfa", "mba"}

# Obvious non-person or junk patterns (case-insensitive substring check)
SKIP_SUBSTRINGS = [
    "online", "section", "pending", "staff", "instructor", "tba", "tbd",
    "ufv", "room", "building", "campus", "abbotsford", "chilliwack",
    "mission", "tel:", "x ", "cepa", "aba", "abb", "abk", "abc", "abd"
]

def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def strip_accents(s: str) -> str:
    # Convert to ASCII by removing diacritics
    nkfd = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in nkfd if not unicodedata.combining(ch))

def clean_token(tok: str) -> str:
    """
    Keep letters & hyphens only. Drop apostrophes/periods/commas, etc.
    Convert to ASCII (strip accents).
    """
    t = strip_accents(tok)
    t = t.replace("’", "").replace("'", "")  # drop apostrophes
    t = t.replace(".", "").replace(",", "")  # drop dots/commas
    # Keep letters and hyphens
    t = re.sub(r"[^A-Za-z\-]", "", t)
    return t

def is_initial(tok: str) -> bool:
    # e.g., "J.", "A", "R." — short tokens often used as middle initials
    return bool(re.fullmatch(r"[A-Za-z]\.?", tok.strip()))

def looks_like_person_name(name: str) -> bool:
    s = name.lower()
    if any(sub in s for sub in SKIP_SUBSTRINGS):
        return False
    # If the whole name has digits, it's suspicious
    if re.search(r"\d", s):
        return False
    # Needs at least two word-like parts
    parts = re.findall(r"[A-Za-zÀ-ÖØ-öø-ÿ'\-]+", name)
    return len(parts) >= 2

def split_name(name_raw: str):
    """
    Return (first, last) or (None, None) if we can't confidently parse.
    - Handle 'Last, First Middle' and 'First Middle Last'
    - Drop titles/suffixes/initials
    """
    name = norm_space(name_raw)

    # Early reject non-person-like rows
    if not looks_like_person_name(name):
        return None, None

    # If comma style: "Last, First Middle ..."
    if "," in name:
        last_part, first_part = [norm_space(x) for x in name.split(",", 1)]
        left = [p for p in last_part.split() if p]
        right = [p for p in first_part.split() if p]
        left = [t for t in left if t.lower() not in TITLES and t.lower() not in SUFFIXES]
        right = [t for t in right if t.lower() not in TITLES and t.lower() not in SUFFIXES and not is_initial(t)]
        if not right or not left:
            return None, None
        first = right[0]
        last = left[-1]
    else:
        # "First Middle Last"
        parts = [p for p in name.split() if p]
        # remove titles/suffixes and initials
        parts = [t for t in parts if t.lower() not in TITLES and t.lower() not in SUFFIXES and not is_initial(t)]
        if len(parts) < 2:
            return None, None
        first = parts[0]
        last = parts[-1]

    # Clean tokens to email-safe (letters + hyphen)
    first = clean_token(first)
    last = clean_token(last)

    if not first or not last:
        return None, None

    return first, last

def format_local_part(first: str, last: str) -> str:
    if LOCAL_PART_CASE == "lower":
        return f"{first.lower()}.{last.lower()}"
    # Default "title" case: First.Last (like directory’s display)
    return f"{first.capitalize()}.{last.capitalize()}"

def make_email(name: str) -> str:
    first, last = split_name(name)
    if not first or not last:
        return ""
    return f"{format_local_part(first, last)}@ufv.ca"

def main():
    input_path = Path(IN_CSV)
    output_path = Path(OUT_CSV)

    rows_in = []
    with input_path.open(newline="", encoding="utf-8") as f:
        rdr = csv.DictReader(f)
        for r in rdr:
            rows_in.append({"name": r["Prof Name"], "course": r["Course Number"]})

    with output_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["name", "email", "course"])
        w.writeheader()
        for r in rows_in:
            name, course = r["name"], r["course"]
            email = make_email(name)
            if email:
                print(f"[GEN ] {name}, {email}, {course}")
            else:
                print(f"[SKIP] {name}, (no pattern), {course}")
            w.writerow({"name": name, "email": email, "course": course})

    print(f"\nWrote {output_path} with {len(rows_in)} rows.")

if __name__ == "__main__":
    main()
