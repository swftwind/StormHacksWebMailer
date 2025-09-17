import os, time
import pandas as pd
from pathlib import Path
from urllib.parse import quote
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo  # Python 3.9+
import platform
from collections import OrderedDict

# -------- CONFIG --------
# CSV_FILE    = "email-lists/tests-and-samples/Sample Professor Outreach List - MSE.csv"
CSV_FILE    = "email-lists/bcit/output-clean.csv"
YOUR_NAME   = "Josie"
YOUR_ROLE   = "Marketing Director"
SUBJECT     = "Student Event Opportunity from SFU Surge"

SCHEDULE_EMAILS = True  # False = create Drafts only; True = uses "Send later" to schedule
SCHEDULE_AT     = "2025-09-17 08:30"          # local clock time to schedule emails for
SCHEDULE_TZ     = "America/Vancouver"       # python classifies us as america >:/

# staggering (only applies when SCHEDULE_EMAILS = True)
STAGGER_SCHEDULE = True          # True = add delays by batch
STAGGER_BATCH_SIZE = 10           # every N emails…
STAGGER_INCREMENT_MINUTES = 5     # …add this many minutes

# ========================= DO NOT TOUCH
OWA_HOME    = "https://outlook.office.com/mail/"
OWA_COMPOSE = "https://outlook.office.com/mail/deeplink/compose"
# =========================

# Use the browser/profile you already use for Outlook Web (so SSO/MFA is already signed in)
EDGE_DATA   = Path(rf"C:\Users\{os.getlogin()}\AppData\Local\Microsoft\Edge\User Data")
CHROME_DATA = Path(rf"C:\Users\{os.getlogin()}\AppData\Local\Google\Chrome\User Data")
PROFILE_NAME = "Default"   # or "Profile 1", etc.

EMAIL_TEMPLATE = """Dear [Professor’s Name],
My name is [Your Name], I am a [Position] at SFU Surge. 

SFU Surge is a student-led organization at Simon Fraser University that empowers students to connect with the tech industry and gain practical experience through meaningful initiatives. 

We are excited to invite students from [COURSE_PHRASE] at BCIT to participate in our annual flagship event StormHacks, one of Western Canada’s largest 24-hour hackathons. 

StormHacks provides a unique opportunity for students to kickstart their careers in tech by transforming their ideas into innovative projects. 
We welcome students from all institutions and backgrounds as we believe innovation comes from a multitude of perspectives and experiences. 

With the support of our sponsors, including Major League Hacking, Microsoft, Safe Software, and more, students can gain access to mentorship, workshops, and networking opportunities.



If you’re open to helping us, we have, for your convenience, included a short announcement below that you can share with your students via your preferred communication platform:

*Title: SFU Surge StormHacks Hackathon, Oct 4-5th*

StormHacks is the sandbox for innovators to brew up their boldest ideas, where 24 hours of intensive building transform ambitious concepts into projects that can shape the future. 

Hosted by SFU Surge, StormHacks is a 24-hour event with substantial opportunities for students to build their own projects, network with veteran industry professionals, and engage with a rapidly growing tech community. Admission is free, and food is provided!

Attending professionals and companies include Major League Hacking, Microsoft, AMD, Safe Software, Huawei, Vercel, Scalar, and many more.

Applications are now live! Apply before September 22nd @ 11:59. For any inquiries, please contact @sfusurge on Instagram.

Link to our official website: https://www.stormhacks.com/

We welcome any questions, concerns, and inquiries. Please let us know if you’re able to help get this out to as many students as possible. Your support and time are greatly appreciated.. 
"""

def prof_salutation(name: str) -> str:
    """
    use 'Professor <Full Name>' by default.
    if the provided name already starts with a title (Prof/Dr/Doctor),
    keep it as-is. if name is missing, fallback to 'Professor'.
    """
    if not name:
        return "Professor"
    clean = str(name).strip()
    lowered = clean.lower()
    if lowered.startswith(("prof", "dr", "doctor")):
        return clean  # already titled; preserve exactly
    return f"Professor {clean}"

def course_phrase(courses):
    if not courses: return "your class"
    if len(courses) == 1: return f"your {courses[0]} class"
    if len(courses) == 2: return f"your {courses[0]} and {courses[1]} classes"
    return f"your {', '.join(courses[:-1])}, and {courses[-1]} classes"

def render_body_text(name, courses):
    return (EMAIL_TEMPLATE
            .replace("[Professor’s Name]", prof_salutation(name))
            .replace("[Your Name]", YOUR_NAME)
            .replace("[Position]", YOUR_ROLE)
            .replace("[COURSE_PHRASE]", course_phrase(courses)))

def load_recipients(csv_path):
    # Read headered CSV if present; otherwise assign headers
    try:
        df_try = pd.read_csv(csv_path)
        if [c.strip().lower() for c in df_try.columns[:2]] == ["name", "email"]:
            df = df_try.rename(columns=lambda c: c.strip())
        else:
            raise ValueError
    except Exception:
        df = pd.read_csv(csv_path, header=None)
        df = df.rename(columns={0:"Name",1:"Email",2:"Course",3:"Times",4:"Campus",5:"Room"})

    # Keep only known cols and normalize
    df = df[["Name","Email","Course","Times","Campus","Room"][:len(df.columns)]]
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = (df[col].astype(str).str.strip()
                                  .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA}))
    df = df.dropna(how="all")

    # Skip an accidental header-as-row
    df = df[~(
        df["Name"].str.fullmatch(r"(?i)name", na=False) &
        df["Email"].str.fullmatch(r"(?i)email", na=False)
    )]

    # Stream rows, carrying forward last seen Name/Email to attach course-only rows
    people = OrderedDict()
    current_name = None
    current_email = None

    for _, r in df.iterrows():
        name  = r.get("Name")
        email = r.get("Email")
        course = r.get("Course")

        if pd.notna(name):  current_name  = str(name).strip()
        if pd.notna(email): current_email = str(email).strip()

        # Only proceed once we have a usable email and a course
        if not current_email or "@" not in current_email:
            continue
        c = (str(course) if pd.notna(course) else "").strip()
        if not c:
            continue

        key = (current_name or "Professor", current_email)
        if key not in people:
            people[key] = {"Name": key[0], "Email": key[1], "Courses": []}
        if c not in people[key]["Courses"]:
            people[key]["Courses"].append(c)

    return list(people.values())

def save_and_close(page):
    page.keyboard.press("Control+S")   # force save
    time.sleep(1.2)
    # close composer
    for sel in ['button[aria-label="Close"]','button[title="Close"]','button[aria-label*="Close"]']:
        try:
            page.locator(sel).first.click(timeout=1000)
            break
        except Exception:
            pass
    # save prompt, if shown
    for sel in ['button:has-text("Save")','button[aria-label="Save"]']:
        try:
            page.locator(sel).first.click(timeout=800)
            break
        except Exception:
            pass

def pick_profile():
    if (EDGE_DATA / PROFILE_NAME).exists():
        return "msedge", str(EDGE_DATA), PROFILE_NAME
    if (CHROME_DATA / PROFILE_NAME).exists():
        return "chrome", str(CHROME_DATA), PROFILE_NAME
    return "msedge", None, None

def parse_schedule(when_str: str, tz_name: str) -> tuple[str, str]:
    """
    Returns (date_str, time_str) in formats OWA accepts in the 'Send later' dialog.
    We'll try multiple formats when typing, so just compute primary ones here.
    """
    tz = ZoneInfo(tz_name)
    dt = datetime.strptime(when_str, "%Y-%m-%d %H:%M").replace(tzinfo=tz)

    # Date & time strings Outlook Web commonly accepts in the dialog inputs
    date_mmddyyyy = dt.strftime("%m/%d/%Y")
    # time without leading zero hour (Windows uses %#I, others %-I)
    hour_fmt = "%#I:%M %p" if platform.system() == "Windows" else "%-I:%M %p"
    time_ampm = dt.strftime(hour_fmt)
    return date_mmddyyyy, time_ampm

def click_first(page, selectors, timeout=1000):
    for sel in selectors:
        try:
            page.locator(sel).first.wait_for(timeout=timeout)
            page.locator(sel).first.click()
            return True
        except Exception:
            continue
    return False

def schedule_send_owa(page, when_date_str: str, when_time_str: str) -> bool:
    """
    Outlook Web: open Send menu -> Schedule send -> Custom time,
    set date/time, confirm. Returns True on success, False on any miss.
    """

    def click_first(selectors, timeout=1000):
        for sel in selectors:
            try:
                page.locator(sel).first.wait_for(timeout=timeout)
                page.locator(sel).first.click()
                return True
            except Exception:
                continue
        return False

    # 1) Open Send menu (chevron). Try several ways.
    opened = click_first([
        # Chevron next to Send
        'button[aria-label*="Send options"]',
        'button[title*="Send options"]',
        'button[aria-haspopup="menu"][aria-label*="Send"]',
        'button[aria-label="Send"] + button',        # Send then chevron
        'button:has(svg[data-icon-name="ChevronDown"])'
    ], timeout=300)
    if not opened:
        try:
            page.locator('button[aria-label="Send"]').first.focus()
            page.keyboard.press("ArrowDown")  # opens the menu on some builds
            opened = True
        except Exception:
            pass
    if not opened:
        print("Could not open Send menu.")
        return False

    # 2) Click "Schedule send"
    if not click_first([
        'div[role="menuitem"]:has-text("Schedule send")',
        'button:has-text("Schedule send")',
    ], timeout=300):
        print("Could not find 'Schedule send' in menu.")
        return False

    # 3) In the first dialog, click "Custom time"
    if not click_first([
        'div[role="dialog"] button:has-text("Custom time")',
        'button:has-text("Custom time")',
        'div[role="dialog"] a:has-text("Custom time")',
    ], timeout=300):
        print("Could not find 'Custom time' option.")
        return False

    # 4) In the "Set custom date and time" dialog, fill date and time
    # Date
    date_filled = False
    for sel in [
        'div[role="dialog"] input[aria-label="Select a date"]',
        'div[role="dialog"] input[placeholder*="date"]',
        'div[role="dialog"] input[type="text"]'
    ]:
        try:
            inp = page.locator(sel).first
            inp.wait_for(timeout=300)
            inp.click()
            inp.fill(when_date_str)   # e.g., 09/15/2025
            inp.press("Enter")
            date_filled = True
            break
        except Exception:
            continue

    # Time
    time_filled = False
    for sel in [
        'div[role="dialog"] input[aria-label="Select a time"]',
        'div[role="dialog"] input[placeholder*="time"]',
        'div[role="dialog"] input[type="text"]'
    ]:
        try:
            # if multiple inputs exist, the 2nd is usually time
            candidates = page.locator(sel)
            inp = candidates.nth(1) if candidates.count() > 1 else candidates.first
            inp.wait_for(timeout=300)
            inp.click()
            inp.fill(when_time_str)   # e.g., 9:00 AM
            inp.press("Enter")
            time_filled = True
            break
        except Exception:
            continue

    if not (date_filled and time_filled):
        print("Could not set custom date/time.")
        return False

    # 5) Confirm (Send)
    if not click_first([
        'div[role="dialog"] button:has-text("Send")',
        'div[role="dialog"] button:has-text("Schedule send")',
        'div[role="dialog"] button[aria-label="Send"]',
    ], timeout=300):
        print("Could not confirm schedule.")
        return False

    page.wait_for_timeout(300)  # let OWA finish
    return True

def schedule_fields_for_index(index: int) -> tuple[str, str]:
    """
    For email #index (0-based), return (date_str, time_str) adjusted for staggering.
    Uses SCHEDULE_AT/SCHEDULE_TZ and STAGGER_* config.
    """
    tz = ZoneInfo(SCHEDULE_TZ)
    base_dt = datetime.strptime(SCHEDULE_AT, "%Y-%m-%d %H:%M").replace(tzinfo=tz)

    extra_minutes = 0
    if STAGGER_SCHEDULE:
        extra_minutes = (index // STAGGER_BATCH_SIZE) * STAGGER_INCREMENT_MINUTES

    adj = base_dt + timedelta(minutes=extra_minutes)
    # Reuse your existing formatter to produce the strings OWA expects
    when_str = adj.strftime("%Y-%m-%d %H:%M")
    return parse_schedule(when_str, SCHEDULE_TZ)

def main():
    recips = load_recipients(CSV_FILE)
    # recips.pop()
    print(f"Loaded {len(recips)} professors")
    channel, user_data_dir, profile = pick_profile()
    print(f"Using browser: {channel} | profile: {profile or '(none)'}")

    from playwright.sync_api import TimeoutError as PWTimeout
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        launch_kwargs = dict(channel=channel, headless=True)
        if user_data_dir: launch_kwargs["user_data_dir"] = os.path.join(user_data_dir, profile)
        ctx = p.chromium.launch_persistent_context(**launch_kwargs)
        page = ctx.new_page()

        # ensure mailbox session
        page.goto(OWA_HOME, wait_until="domcontentloaded")
        page.wait_for_timeout(2500)

        made = 0
        scheduled = False
        scheduled_when = None

        for r in recips:
            email = (r["Email"] or "").strip()
            if not email or email.lower() in ("nan","none"): continue
            body_text = render_body_text(r["Name"], r["Courses"])

            # prefill To + Subject via query (Body via typing; body=... can be too long for URL)
            compose_url = f"{OWA_COMPOSE}?to={quote(email)}&subject={quote(SUBJECT)}"
            page.goto(compose_url, wait_until="domcontentloaded")

            # Subject (in case OWA didn't apply it)
            for sel in ['input[aria-label="Add a subject"]',
                        'input[placeholder="Add a subject"]',
                        'input[aria-label="Subject"]']:
                try:
                    subj = page.locator(sel).first
                    subj.wait_for(timeout=1000); subj.click(); subj.fill(SUBJECT)
                    break
                except PWTimeout:
                    continue

            # Body
            body_ok = False
            for sel in [
                '[aria-label="Message body"]',
                'div[contenteditable="true"][role="textbox"]',
            ]:
                try:
                    body = page.locator(sel).first
                    body.wait_for(timeout=1000)
                    body.click()

                    # --- replace this line ---
                    # body.fill(body_text)

                    # --- with this block (makes the URLs clickable) ---
                    body_html = (
                        body_text
                        .replace(
                            "https://sfusurge.com/",
                            '<a href="https://sfusurge.com/">https://sfusurge.com/</a>'
                        )
                        .replace(
                            "https://www.stormhacks.com/",
                            '<a href="https://www.stormhacks.com/">https://www.stormhacks.com/</a>'
                        )
                    )

                    handle = body.element_handle()
                    handle.evaluate(
                        "(el, html) => { el.innerHTML = html.replace(/\\n/g,'<br>'); }",
                        body_html,
                    )
                    # --- end replacement ---
                    
                    # ensure focus is in the editor, jump to end, and trigger the preview
                    body.click()                         # refocus the compose editor
                    page.keyboard.press("Control+End")   # caret to end (right after the link)
                    page.keyboard.press("Backspace")     # remove any trailing space after the URL
                    page.keyboard.press("Enter")         # this makes OWA create the rich preview
                    page.wait_for_timeout(1200)          # small wait for the card to render
                    # (optional) send Enter again if it occasionally misses:
                    # page.keyboard.press("Enter")

                    # === Schedule or Draft ===
                    if SCHEDULE_EMAILS:
                        date_str, time_str = schedule_fields_for_index(made)
                        scheduled = schedule_send_owa(page, date_str, time_str)
                        if scheduled:
                            scheduled_when = f"{date_str} {time_str}"
                        else:
                            print("Scheduling failed; saving as Draft instead.")
                            save_and_close(page)
                    else:
                        save_and_close(page)

                    body_ok = True
                    break
                except PWTimeout:
                    continue
            if not body_ok:
                print("Could not find message body; skipping this one.")
                continue

            # save_and_close(page)
            made += 1
            print(f"{'Scheduled' if scheduled else 'Draft'} for {email}" + (f" at {scheduled_when}" if scheduled_when else ""))

        print(f"\nDone. Drafts created: {made}")
        ctx.close()

if __name__ == "__main__":
    main()
