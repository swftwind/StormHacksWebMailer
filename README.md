# StormHacks Mailer (Outlook Web + Playwright)

Automate composing personalized emails in **Outlook on the Web** from a **CSV**.  
The script opens a browser window (Edge by default), fills **To / Subject / Body**, and either **saves a Draft** or **schedules** each message at a time you choose — all **without** Microsoft Graph or admin permissions.

---

## Prerequisites

- **Windows** with **Microsoft Edge** (or Chrome) and access to Outlook on the Web.
- **Python 3.9+**
- Python packages and browser drivers:
  ```bash
  pip install playwright pandas
  python -m playwright install
  ```
  > The second line downloads the browser binaries that Playwright uses.

---

## Configure

Open the script and edit the config block near the top:

```python
# -------- CONFIG --------
CSV_FILE    = "email-lists/Professor Outreach List - MSE.csv"  # your recipients file
YOUR_NAME   = "Arseniy"                             # appears in the email body
YOUR_ROLE   = "Marketing Coordinator"               # appears in the email body
SUBJECT     = "SFU Surge StormHacks In-Class Presentation"

# Scheduling
SCHEDULE_EMAILS = False                  # False = draft only; True = schedule send
SCHEDULE_AT     = "2025-09-15 09:00"     # local date/time for 'Send later'
SCHEDULE_TZ     = "America/Vancouver"    # timezone for the schedule time (IANA name)

# Browser profile (used to reuse your login session)
PROFILE_NAME = "Default"  # or "Profile 1", etc.
```

**What to expect:**  
- The script opens an **Edge** window using your existing Windows profile folder. If Edge isn’t found, it tries Chrome.  
- If you’re not already signed into **Outlook on the Web** in that profile, the first run may require you to **log in once** (SSO/MFA). The session is then reused for subsequent runs.

---

## CSV format (follow the sample)

Your CSV should follow the same shape as **`Sample Professor Outreach List - MSE.csv`**. Minimal example:

```csv
Name,Email,Course,Times,Campus,Room
Krishna Vijayaraghavan,krishna@sfu.ca,MSE 103,Tu 10:30AM - 11:20AM,ALL SURREY,SRYE3016
,,MSE 352,Mo 2:30PM - 4:20PM,SRY...,SRYC5240
Erik Kjeang,ekjeang@sfu.ca,MSE 210,Mo 2:30PM - 4:20PM,SRYE2016,
```

- **Multiple courses per professor:** leave **Name** and **Email** blank on the extra rows — the script automatically attaches those to the professor above and will phrase the message like “your **MSE 103 and MSE 352** classes.”  
- Only **Name**, **Email**, and **Course** are required; other columns may be blank or omitted.

---

## Run

From the project folder:

```bash
python StormHacksMailerWeb.py
```

**What happens:**  
- A browser window opens to Outlook Web.  
- For each professor in the CSV, the script:
  - opens a **compose** window (very fast via keyboard shortcut after the first message),
  - fills **To**, **Subject**, and **Body** (with clickable links),
  - optionally triggers the rich **link preview** for the StormHacks URL,
  - then either:
    - **saves a Draft**, or
    - **Schedule send → Custom time**, fills your configured date/time, and **schedules** it.

You can **cancel** at any time from the terminal with **Ctrl + C**.

---

## Drafts vs. Scheduled

- `SCHEDULE_EMAILS = False` → **Drafts only** (review and send manually).  
- `SCHEDULE_EMAILS = True`  → uses Outlook Web’s **Schedule send**:
  - `SCHEDULE_AT` must be in `YYYY-MM-DD HH:MM` (local time for the mailbox).  
  - `SCHEDULE_TZ` should be a valid IANA timezone (e.g., `America/Vancouver`).  
  - If the schedule dialog can’t be completed for any reason, the script **falls back to saving a Draft** and logs a message.

---

## Tips / Troubleshooting

- If fields aren’t found, ensure **PROFILE_NAME** matches the Edge profile you actually use (e.g., `Profile 1`).  
- If a “phantom” recipient appears, make sure your CSV **doesn’t include stray header/blank rows**; follow the sample format exactly.  
- Link preview adds a small delay; the script can be configured to skip it if you need maximum speed.

---

## Dependencies Recap

- Python 3.9+  
- Packages: `playwright`, `pandas`  
- One-time driver install: `python -m playwright install`

---