# Clean Langara dataset:
#  - Remove rows with section headings in the "Prof Name" column
#  - Remove rows where multiple profs are listed in a single row
#  - Keep all other rows as-is

import re
import pandas as pd
from pathlib import Path

IN_CSV  = "../datasets/langara_classes_fall_2025.csv"
OUT_CSV = "../datasets/langara_classes_fall_2025_clean.csv"

def norm_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def is_heading(name: str) -> bool:
    """Detect obvious section headings instead of names."""
    if not name:
        return True
    t = norm_space(name)
    # headings are usually single words or generic labels
    if len(t.split()) == 1 and t.isalpha() and t[0].isupper():
        return True
    # common section terms
    if t.lower() in {"arts","science","sciences","business","health","humanities","social sciences","program","programs","departments","department","school","schools"}:
        return True
    return False

def has_multiple_names(name: str) -> bool:
    """Detect if multiple names are listed in the same cell."""
    if not name:
        return False
    t = name.strip()
    # separators that usually mean multiple profs
    if any(sep in t for sep in ["/", "&", " and ", ";", ", "]):
        return True
    return False

def main():
    df = pd.read_csv(IN_CSV, dtype=str).fillna("")

    # Normalize whitespace
    for c in df.columns:
        df[c] = df[c].map(norm_space)

    name_col = "Prof Name" if "Prof Name" in df.columns else df.columns[0]

    # Apply filters
    mask_heading = df[name_col].map(is_heading)
    mask_multi   = df[name_col].map(has_multiple_names)

    df_clean = df[~(mask_heading | mask_multi)].copy()

    # Write out
    Path(OUT_CSV).parent.mkdir(parents=True, exist_ok=True)
    df_clean.to_csv(OUT_CSV, index=False)

    print(f"[DONE] Wrote {OUT_CSV} with {len(df_clean)} rows (from {len(df)}).")

if __name__ == "__main__":
    main()
