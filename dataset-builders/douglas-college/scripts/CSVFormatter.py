# Quick CSV cleaner and stacker
# - Removes rows with faculty TBD (case-insensitive) or missing email
# - Groups by professor name + email, stacking multiple courses so name/email only appear once
# - Writes to /mnt/data/output-clean.csv

import pandas as pd
import re
from pathlib import Path

in_path = Path("output.csv")
out_path = Path("output-clean.csv")

# Read CSV with flexible options
df = pd.read_csv(in_path, dtype=str).fillna("")

# Heuristically detect columns
cols_lower = [c.lower() for c in df.columns]

def pick_col(candidates):
    for i, c in enumerate(cols_lower):
        if any(tok in c for tok in candidates):
            return df.columns[i]
    return None

name_col = pick_col(["name", "faculty", "instructor", "prof"])
email_col = pick_col(["email", "e-mail", "mail"])
course_col = pick_col(["course", "code", "catalog", "subject", "number"])

# Fallbacks if detection fails
if name_col is None and len(df.columns) >= 1:
    name_col = df.columns[0]
if email_col is None and len(df.columns) >= 2:
    email_col = df.columns[1]
if course_col is None and len(df.columns) >= 3:
    course_col = df.columns[2]

# Normalize whitespace
for c in [name_col, email_col, course_col]:
    df[c] = df[c].astype(str).str.strip()

# Filter: remove rows with faculty TBD or missing email
mask_tbd = df[name_col].str.lower().str.contains(r"\btbd\b") | df[name_col].str.lower().str.contains("to be determined")
mask_missing_email = (df[email_col] == "") | (~df[email_col].str.contains("@"))
df_clean = df[~mask_tbd & ~mask_missing_email].copy()

# Group by name + email and stack courses
out_rows = []
for (name, email), g in df_clean.groupby([name_col, email_col]):
    courses = [c for c in g[course_col].tolist() if c != ""]
    if not courses:
        continue
    # First row with name/email + first course
    out_rows.append({name_col: name, email_col: email, course_col: courses[0]})
    # Subsequent rows with only course
    for c in courses[1:]:
        out_rows.append({name_col: "", email_col: "", course_col: c})

out_df = pd.DataFrame(out_rows, columns=[name_col, email_col, course_col])

# Write output
out_df.to_csv(out_path, index=False)