# ==============================================================
#   JOB HUNTING AI AUTOMATION TOOL  v2.0
#   ─────────────────────────────────────────────────────────
#   ✅ Salesforce roles ONLY  (title whitelist + blacklist)
#   ✅ Remote jobs ONLY       (API flag + post-filter)
#   ✅ No duplicates across daily runs  (seen_jobs.json)
#   ✅ searchTerms batching   (saves Apify credits)
#   ✅ Formatted Excel        (6 sheets, clickable links)
#
#   Powered by: Apify + openclawai/job-board-scraper
#   API Docs  : https://apify.com/openclawai/job-board-scraper
# ==============================================================

from apify_client import ApifyClient
import pandas as pd
from datetime import datetime, timedelta
import hashlib, json, os, time, re

# ── Load .env file for local development (gitignored) ──────────
# On GitHub Actions, secrets are injected as real env vars instead.
_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
if os.path.exists(_env_path):
    with open(_env_path, encoding="utf-8") as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())

# ── Google Sheets (cloud mode) ─────────────────────────────────
# Auto-enabled when GitHub Actions sets these two env vars.
# Leave blank for local runs — falls back to Excel output.
_GSHEETS_CREDS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON", "")
_GSHEETS_SHEET_ID   = os.getenv("GOOGLE_SHEET_ID", "")
USE_GOOGLE_SHEETS   = bool(_GSHEETS_CREDS_JSON and _GSHEETS_SHEET_ID)

if USE_GOOGLE_SHEETS:
    try:
        import gspread
        from google.oauth2.service_account import Credentials as GCreds
    except ImportError:
        print("⚠️  Google Sheets mode needs: pip install gspread google-auth")
        USE_GOOGLE_SHEETS = False


# ==============================================================
#   ⚙️  CONFIGURATION  — Edit only this section
# ==============================================================

# Loaded from .env locally, or from GitHub Secret on Actions
APIFY_API_TOKEN = os.getenv("APIFY_API_TOKEN", "")

# ── Platforms ──────────────────────────────────────────────────
SITES = ["indeed", "linkedin", "glassdoor", "google", "zip_recruiter"]

# ── Locations ─────────────────────────────────────────────────
LOCATIONS = [
    "United States",
    "Remote",
]

# ── Keywords (batched ≤5 per API call — saves credits) ────────
# All 6 are split into batches automatically
SEARCH_KEYWORDS = [
    "Salesforce Engineer",
    "Salesforce Developer",
    "Salesforce Architect",
    "Salesforce Administrator",
    "Lead Salesforce",
    "Salesforce CPQ Developer",
]

# ── Filters ───────────────────────────────────────────────────
MAX_RESULTS_PER_SITE  = 20      # per site per API call (max 100)
HOURS_OLD             = 24      # only jobs posted last 24 h
JOB_TYPE              = "fulltime"
COUNTRY_INDEED        = "usa"
ENFORCE_ANNUAL_SALARY = True

# ── Output files ──────────────────────────────────────────────
TIMESTAMP       = datetime.now().strftime("%Y%m%d_%H%M")
OUTPUT_FILE     = f"jobs_{TIMESTAMP}.xlsx"
MASTER_FILE     = "jobs_master.xlsx"
SEEN_JOBS_FILE  = "seen_jobs.json"   # tracks jobs across daily runs (local mode)

# ── Google Sheets tab names ────────────────────────────────────
GSHEETS_ALL_JOBS_TAB  = "All Jobs"   # new jobs appended here every run
GSHEETS_SEEN_JOBS_TAB = "Seen Jobs"  # replaces seen_jobs.json in cloud mode


# ==============================================================
#   🎯  SALESFORCE ROLE FILTER
#   ─────────────────────────────────────────────────────────
#   KEEP a job only if title contains ≥1 word from MUST_INCLUDE
#   SKIP a job if title contains any word from MUST_EXCLUDE
# ==============================================================

# Title MUST contain at least one of these (case-insensitive)
SALESFORCE_MUST_INCLUDE = [
    "salesforce",
    "sfdc",
    "apex developer",
    "apex engineer",
    "lwc developer",
    "lightning web component",
    "visualforce",
    "service cloud",
    "sales cloud",
    "marketing cloud",
    "pardot",
    "salesforce cpq",
    "salesforce admin",
    "salesforce architect",
    "salesforce consultant",
    "salesforce engineer",
    "salesforce developer",
    "salesforce integration",
]

# Title MUST NOT contain any of these — even if "salesforce" appears elsewhere
SALESFORCE_MUST_EXCLUDE = [
    "customer success manager",
    "customer success",
    "sales consultant",
    "sales representative",
    "sales rep",
    "sales associate",
    "dental",
    "insurance",
    "mortgage",
    "real estate",
    "nurse",
    "driver",
    "warehouse",
    "account executive",
    "account manager",
    "business development",
    "marketing manager",
    "hr manager",
    "recruiter",
    "staffing",
    "medical",
    "pharmaceutical",
    "financial advisor",
    "Salesforce Business Analyst",
]

def is_salesforce_role(title: str) -> bool:
    """
    Returns True only if the job title is a genuine Salesforce
    technical/functional role — not a sales or unrelated job.
    """
    if not title:
        return False

    t = title.lower().strip()

    # Step 1 — Must contain at least one Salesforce keyword
    has_sf_keyword = any(kw in t for kw in SALESFORCE_MUST_INCLUDE)
    if not has_sf_keyword:
        return False

    # Step 2 — Must not contain any exclusion keyword
    has_bad_keyword = any(kw in t for kw in SALESFORCE_MUST_EXCLUDE)
    if has_bad_keyword:
        return False

    return True


def is_truly_remote(job: dict) -> bool:
    """
    Double-checks remote status using multiple job fields.
    The API isRemote flag sometimes misses remote jobs or
    includes hybrid/on-site. We check 3 signals.
    """
    # Signal 1: API remote flag
    api_flag = job.get("is_remote", False) is True

    # Signal 2: Location string contains remote indicators
    loc = str(job.get("location", "") or "").lower()
    loc_remote = any(w in loc for w in [
        "remote", "work from home", "wfh", "anywhere", "virtual"
    ])

    # Signal 3: Title or description contains remote indicators
    title = str(job.get("title", "") or "").lower()
    desc  = str(job.get("description", "") or "").lower()[:500]
    text_remote = any(w in title + " " + desc for w in [
        "remote", "work from home", "100% remote", "fully remote"
    ])

    # Signal 4: Explicitly NOT remote — reject these
    on_site_signals = any(w in loc for w in [
        "on-site", "on site", "onsite", "in-office", "in office"
    ])
    if on_site_signals:
        return False

    return api_flag or loc_remote or text_remote


# ==============================================================
#   ☁️  GOOGLE SHEETS CLIENT  (cloud mode only)
# ==============================================================

_gc     = None   # gspread.Client  — set by init_gsheets()
_gsheet = None   # gspread.Spreadsheet


def init_gsheets():
    """Connect to Google Sheets using service-account credentials from env var."""
    global _gc, _gsheet
    if not USE_GOOGLE_SHEETS:
        return
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds   = GCreds.from_service_account_info(
                  json.loads(_GSHEETS_CREDS_JSON), scopes=SCOPES)
    _gc     = gspread.authorize(creds)
    _gsheet = _gc.open_by_key(_GSHEETS_SHEET_ID)
    print(f"  ☁️  Google Sheets connected: '{_gsheet.title}'")


# ==============================================================
#   🗂️  SEEN JOBS TRACKER  — cross-run deduplication
# ==============================================================

class SeenJobsTracker:
    """
    Saves/loads a JSON file that remembers every job URL
    we have fetched. Before saving a job we check this set.
    New jobs are added to the file after each run.
    """

    def __init__(self, filepath: str = SEEN_JOBS_FILE):
        self.filepath   = filepath
        self.seen_urls  : set = set()
        self.seen_hash  : set = set()
        self._new_urls  : set = set()   # only entries added THIS run
        self._new_hashes: set = set()   # only entries added THIS run
        self._load()

    def _load(self):
        # ── Cloud mode: load from Google Sheets "Seen Jobs" tab ──
        if USE_GOOGLE_SHEETS and _gsheet is not None:
            try:
                ws   = _gsheet.worksheet(GSHEETS_SEEN_JOBS_TAB)
                rows = ws.get_all_values()
                for row in rows[1:]:          # skip header row
                    if len(row) >= 1 and row[0]:
                        self.seen_urls.add(row[0])
                    if len(row) >= 2 and row[1]:
                        self.seen_hash.add(row[1])
                print(f"  ☁️  Loaded {len(self.seen_urls):,} seen jobs "
                      f"from Google Sheets")
            except gspread.WorksheetNotFound:
                ws = _gsheet.add_worksheet(
                    GSHEETS_SEEN_JOBS_TAB, rows=50000, cols=3)
                ws.append_row(["url", "hash", "added_on"])
                print(f"  ☁️  Created '{GSHEETS_SEEN_JOBS_TAB}' tab "
                      f"(first run)")
            except Exception as e:
                print(f"  ⚠️  Could not load seen jobs from Sheets: {e}")
            return

        # ── Local mode: load from seen_jobs.json ──────────────────
        if os.path.exists(self.filepath):
            try:
                with open(self.filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.seen_urls = set(data.get("seen_urls", []))
                self.seen_hash = set(data.get("seen_hash", []))
                print(f"  📂 Loaded {len(self.seen_urls):,} seen jobs "
                      f"from {self.filepath}")
            except Exception as e:
                print(f"  ⚠️  Could not load seen jobs: {e}")
        else:
            print(f"  📂 No seen-jobs file yet — starting fresh")

    def save(self):
        # ── Cloud mode: append NEW entries to Google Sheets ───────
        if USE_GOOGLE_SHEETS and _gsheet is not None:
            if not self._new_urls and not self._new_hashes:
                print(f"  ☁️  No new seen-job entries to append")
                return
            try:
                ws   = _gsheet.worksheet(GSHEETS_SEEN_JOBS_TAB)
                now  = datetime.now().strftime("%Y-%m-%d %H:%M")
                rows = []
                for url, h in zip(
                    sorted(self._new_urls),
                    sorted(self._new_hashes)
                ):
                    rows.append([url, h, now])
                # pad if lengths differ
                for h in sorted(self._new_hashes)[len(rows):]:
                    rows.append(["", h, now])
                ws.append_rows(rows, value_input_option="USER_ENTERED")
                print(f"  ☁️  Appended {len(rows)} new seen-job "
                      f"entries to Google Sheets")
            except Exception as e:
                print(f"  ⚠️  Could not save seen jobs to Sheets: {e}")
            return

        # ── Local mode: overwrite seen_jobs.json ──────────────────
        data = {
            "seen_urls"   : list(self.seen_urls),
            "seen_hash"   : list(self.seen_hash),
            "total"       : len(self.seen_urls),
            "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }
        with open(self.filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        print(f"  💾 Saved {len(self.seen_urls):,} seen jobs → {self.filepath}")

    @staticmethod
    def _make_hash(title: str, company: str) -> str:
        """Fingerprint from title + company (fallback when no URL)"""
        raw = f"{title.lower().strip()}|{company.lower().strip()}"
        return hashlib.md5(raw.encode()).hexdigest()

    def is_seen(self, job: dict) -> bool:
        """Returns True if this job was already fetched in a previous run"""
        url  = job.get("job_url") or job.get("url") or ""
        h    = self._make_hash(
                   job.get("title", ""),
                   job.get("company", "")
               )
        return url in self.seen_urls or h in self.seen_hash

    def mark_seen(self, job: dict):
        """Adds a job to the seen set (call after keeping a job)"""
        url = job.get("job_url") or job.get("url") or ""
        h   = self._make_hash(
                  job.get("title", ""),
                  job.get("company", "")
              )
        if url:
            self.seen_urls.add(url)
            self._new_urls.add(url)
        self.seen_hash.add(h)
        self._new_hashes.add(h)


# ==============================================================
#   🔍  FETCH JOBS FROM APIFY
# ==============================================================

def batch_keywords(keywords: list, size: int = 5) -> list:
    """Splits keywords into batches of ≤5 (API limit for searchTerms)"""
    return [keywords[i:i + size] for i in range(0, len(keywords), size)]


def fetch_jobs_batch(keyword_batch: list, location: str) -> list:
    """
    Calls openclawai/job-board-scraper with searchTerms[] array.
    One API call fetches up to 5 keywords at once — saves credits.
    isRemote=True is set at API level.
    """
    client = ApifyClient(APIFY_API_TOKEN)

    label = " | ".join(keyword_batch[:2]) + \
            (f" +{len(keyword_batch)-2} more" if len(keyword_batch) > 2 else "")
    print(f"\n    🔍 [{label}]  📍 {location}")

    run_input = {
        # ── Use searchTerms (array) for batching multiple keywords ──
        "searchTerms"         : keyword_batch,   # up to 5 keywords/call
        "location"            : location,
        "sites"               : SITES,
        "maxResults"          : MAX_RESULTS_PER_SITE,
        "hoursOld"            : HOURS_OLD,
        "isRemote"            : True,            # ✅ remote only at API level
        "jobType"             : JOB_TYPE,
        "countryIndeed"       : COUNTRY_INDEED,
        "enforceAnnualSalary" : ENFORCE_ANNUAL_SALARY,
        "descriptionFormat"   : "markdown",
    }

    try:
        run = client.actor("openclawai/job-board-scraper").call(
            run_input=run_input,
            wait_duration=timedelta(minutes=10)
        )

        if not run:
            print("      ⚠️  No run returned")
            return []

        dataset_id = (
            run["defaultDatasetId"]
            if isinstance(run, dict)
            else run.default_dataset_id
        )

        jobs = list(client.dataset(dataset_id).iterate_items())
        print(f"      ✅ {len(jobs)} raw jobs returned by API")
        return jobs

    except Exception as e:
        print(f"      ❌ API Error: {e}")
        return []


# ==============================================================
#   🐛  DEBUG HELPER — dumps first N raw jobs to JSON
#       so you can inspect every field the API actually returns
# ==============================================================

DEBUG_MODE      = True          # set False to disable
DEBUG_DUMP_FILE = "debug_raw_jobs.json"
_debug_saved    = 0             # internal counter
_debug_limit    = 5             # save first 5 raw jobs per run


def debug_dump(job: dict):
    """Appends raw job dict to debug_raw_jobs.json (first 5 jobs only)."""
    global _debug_saved
    if not DEBUG_MODE or _debug_saved >= _debug_limit:
        return
    existing = []
    if os.path.exists(DEBUG_DUMP_FILE):
        try:
            with open(DEBUG_DUMP_FILE, "r", encoding="utf-8") as f:
                existing = json.load(f)
        except Exception:
            existing = []
    existing.append(job)
    with open(DEBUG_DUMP_FILE, "w", encoding="utf-8") as f:
        json.dump(existing, f, indent=2, default=str)
    _debug_saved += 1


# ==============================================================
#   💰  ROBUST SALARY PARSER
#   Handles all formats LinkedIn / Indeed / Glassdoor return:
#     • Structured numbers  →  salary_min=120000, salary_max=150000
#     • Already-formatted   →  "120,000 - 150,000"
#     • Hourly/monthly      →  enforceAnnualSalary converts these
#     • Salary from text    →  salary_source="description"
#     • Zero / None         →  explicit None checks (avoids 0==False)
# ==============================================================

def parse_salary(job: dict) -> str:
    """
    Extracts salary from every known field variation across
    LinkedIn, Indeed, Glassdoor, Google Jobs, ZipRecruiter.
    Returns a human-readable string or 'Not Listed'.
    """

    # ── Step 1: try all known numeric field names ─────────────
    # Use explicit `is not None` so a $0 value isn't skipped
    s_min = (job.get("salary_min")
             if job.get("salary_min") is not None
             else job.get("min_amount"))

    s_max = (job.get("salary_max")
             if job.get("salary_max") is not None
             else job.get("max_amount"))

    s_cur = (job.get("salary_currency")
             or job.get("currency") or "USD")

    s_int = (job.get("salary_interval")
             or job.get("salary_period")
             or job.get("compensation_type") or "year")

    # Normalise interval label
    interval_map = {
        "yearly": "year", "annual": "year", "annually": "year",
        "monthly": "month", "hourly": "hour", "weekly": "week",
        "hour": "hour", "month": "month", "year": "year",
    }
    s_int = interval_map.get(str(s_int).lower(), str(s_int))

    # ── Step 2: convert to float, build string ────────────────
    def to_num(v):
        if v is None:
            return None
        try:
            cleaned = str(v).replace(",", "").replace("$", "").strip()
            return float(cleaned)
        except Exception:
            return None

    min_f = to_num(s_min)
    max_f = to_num(s_max)

    if min_f is not None and max_f is not None and max_f > 0:
        return f"${min_f:,.0f} – ${max_f:,.0f} {s_cur}/{s_int}"
    if min_f is not None and min_f > 0:
        return f"${min_f:,.0f}+ {s_cur}/{s_int}"
    if max_f is not None and max_f > 0:
        return f"Up to ${max_f:,.0f} {s_cur}/{s_int}"

    # ── Step 3: fallback — salary buried in a text field ──────
    # Indeed/LinkedIn sometimes return salary as a raw string
    for field in ["salary", "salary_text", "compensation",
                  "salary_range", "pay", "wage"]:
        val = job.get(field)
        if val and str(val).strip() not in ("", "None", "null"):
            return str(val).strip()

    # ── Step 4: scan description for salary patterns ──────────
    # e.g.  "$120,000 – $150,000"  or  "$85/hr"
    desc = str(job.get("description", "") or "")[:800]
    pattern = r'\$[\d,]+(?:\.\d+)?(?:\s*[-–]\s*\$[\d,]+(?:\.\d+)?)?(?:\s*/\s*(?:yr|year|hr|hour|mo|month))?'
    matches = re.findall(pattern, desc, re.IGNORECASE)
    if matches:
        return matches[0].strip()          # first salary mention

    return "Not Listed"


# ==============================================================
#   📅  ROBUST DATE PARSER
#   Handles all formats LinkedIn / Indeed / Glassdoor return:
#     • ISO datetime  →  "2026-05-22T09:15:00+00:00"
#     • Date only     →  "2026-05-22"
#     • Relative      →  "15 hours ago", "2 days ago", "1 week ago"
#     • Unix stamp    →  1716371200
#     • None / ""     →  "Unknown"
# ==============================================================

def parse_date(date_raw) -> str:
    """
    Returns YYYY-MM-DD string for any date format the
    job boards may return. Relative strings are converted
    to absolute dates using today's date as the anchor.
    """
    if date_raw is None:
        return "Unknown"

    s = str(date_raw).strip()

    if not s or s.lower() in ("none", "null", "n/a", ""):
        return "Unknown"

    now = datetime.now()
    sl  = s.lower()

    # ── Relative strings ──────────────────────────────────────
    if any(x in sl for x in ["just posted", "today", "moments ago"]):
        return now.strftime("%Y-%m-%d")

    if "yesterday" in sl:
        return (now - timedelta(days=1)).strftime("%Y-%m-%d")

    # "15 hours ago", "2 hour ago"
    m = re.search(r'(\d+)\s*hour', sl)
    if m:
        return (now - timedelta(hours=int(m.group(1)))).strftime("%Y-%m-%d")

    # "2 days ago", "1 day ago"
    m = re.search(r'(\d+)\s*day', sl)
    if m:
        return (now - timedelta(days=int(m.group(1)))).strftime("%Y-%m-%d")

    # "1 week ago", "2 weeks ago"
    m = re.search(r'(\d+)\s*week', sl)
    if m:
        return (now - timedelta(weeks=int(m.group(1)))).strftime("%Y-%m-%d")

    # "1 month ago", "2 months ago"
    m = re.search(r'(\d+)\s*month', sl)
    if m:
        return (now - timedelta(days=int(m.group(1)) * 30)).strftime("%Y-%m-%d")

    # "30+ days ago" — treat as 30 days
    m = re.search(r'(\d+)\+?\s*days?\s*ago', sl)
    if m:
        return (now - timedelta(days=int(m.group(1)))).strftime("%Y-%m-%d")

    # ── Unix timestamp (integer or string of digits) ───────────
    if re.fullmatch(r'\d{9,11}', s):          # 10-digit unix stamp
        try:
            return datetime.fromtimestamp(int(s)).strftime("%Y-%m-%d")
        except Exception:
            pass

    if re.fullmatch(r'\d{13}', s):            # 13-digit ms timestamp
        try:
            return datetime.fromtimestamp(int(s) / 1000).strftime("%Y-%m-%d")
        except Exception:
            pass

    # ── ISO / standard formats — let pandas handle them ───────
    try:
        return pd.to_datetime(date_raw, utc=True).strftime("%Y-%m-%d")
    except Exception:
        pass

    try:
        return pd.to_datetime(date_raw).strftime("%Y-%m-%d")
    except Exception:
        pass

    # ── Last resort: take first 10 chars if looks like a date ─
    if re.match(r'\d{4}-\d{2}-\d{2}', s):
        return s[:10]

    return s[:20]   # return raw (truncated) so it's not lost


# ==============================================================
#   🧹  FILTER + CLEAN A SINGLE JOB
# ==============================================================

def filter_and_clean(job: dict, tracker: SeenJobsTracker) -> dict | None:
    """
    Applies all 3 filters then returns clean dict or None.

    Filter 1 — Salesforce role check (title whitelist/blacklist)
    Filter 2 — Remote-only check  (API flag + location/description)
    Filter 3 — Deduplication      (seen_jobs.json cross-run check)
    """

    title   = job.get("title", "") or ""
    company = job.get("company", "") or ""

    # ── Filter 1: Salesforce role ─────────────────────────────
    if not is_salesforce_role(title):
        return None

    # ── Filter 2: Remote check ────────────────────────────────
    if not is_truly_remote(job):
        return None

    # ── Filter 3: Duplicate check ─────────────────────────────
    if tracker.is_seen(job):
        return None

    # ── All filters passed — dump raw data for debug ──────────
    debug_dump(job)

    # ── Mark as seen immediately ──────────────────────────────
    tracker.mark_seen(job)

    # ── Salary  (robust parser) ───────────────────────────────
    salary = parse_salary(job)

    # ── Date posted  (robust parser) ──────────────────────────
    date_posted = parse_date(job.get("date_posted"))

    # ── Location ──────────────────────────────────────────────
    city    = job.get("city",  "") or ""
    state   = job.get("state", "") or ""
    country = job.get("country", "") or ""
    loc     = ", ".join(filter(None, [city, state, country])) \
              or job.get("location", "") or "Remote"

    # ── Apply link ────────────────────────────────────────────
    apply_link = (job.get("job_url_direct")
                  or job.get("job_url")
                  or job.get("url")
                  or "N/A")

    # ── Description snippet ───────────────────────────────────
    desc      = (job.get("description") or "")
    desc_short = desc[:400] + "…" if len(desc) > 400 else desc

    # ── Skills ────────────────────────────────────────────────
    skills     = job.get("skills") or []
    skills_str = ", ".join(skills[:10]) if isinstance(skills, list) \
                 else str(skills)

    # ── Salary source tag (LinkedIn/Indeed indicate origin) ───
    sal_src    = job.get("salary_source") or ""   # "direct_data" or "description"

    return {
        "Job Title"        : title,
        "Company Name"     : company,
        "Platform"         : str(job.get("site", "")).capitalize(),
        "Location"         : loc,
        "Remote"           : "✅ Yes",
        "Job Type"         : str(job.get("job_type", JOB_TYPE)).capitalize(),
        "Salary"           : salary,
        "Salary Source"    : sal_src,            # new — shows where salary came from
        "Date Posted"      : date_posted,
        "Job Link"         : apply_link,
        "Company Rating"   : job.get("company_rating", ""),
        "Company Size"     : (job.get("company_employees_label")
                              or job.get("company_num_employees") or ""),
        "Skills Required"  : skills_str,
        "Description"      : desc_short,
        "Scraped On"       : datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


# ==============================================================
#   📊  SAVE TO FORMATTED EXCEL
# ==============================================================

def save_to_excel(df: pd.DataFrame, filename: str):
    print(f"\n  💾 Saving {len(df)} jobs → {filename}")

    writer   = pd.ExcelWriter(filename, engine="xlsxwriter")
    wb       = writer.book

    # ── Formats ───────────────────────────────────────────────
    hdr = wb.add_format({"bold": True, "bg_color": "#1A3C5E",
                          "font_color": "#FFFFFF", "border": 1,
                          "align": "center", "valign": "vcenter",
                          "font_size": 11, "text_wrap": True})
    lnk = wb.add_format({"font_color": "#0563C1", "underline": True})
    ttl = wb.add_format({"bold": True, "font_size": 16,
                          "font_color": "#1A3C5E"})
    sub = wb.add_format({"italic": True, "font_size": 10,
                          "font_color": "#555555"})
    lbl = wb.add_format({"bold": True, "bg_color": "#EEF2F7", "border": 1})
    val = wb.add_format({"border": 1})
    big = wb.add_format({"bold": True, "font_size": 15,
                          "font_color": "#1A3C5E", "align": "center"})

    WIDTHS = {
        "Job Title"       : 40, "Company Name"  : 26,
        "Platform"        : 14, "Location"      : 22,
        "Remote"          : 11, "Job Type"      : 13,
        "Salary"          : 32, "Salary Source" : 16,
        "Date Posted"     : 13, "Job Link"      : 55,
        "Company Rating"  : 15, "Company Size"  : 18,
        "Skills Required" : 40, "Description"   : 55,
        "Scraped On"      : 18,
    }

    def _write_sheet(df_s: pd.DataFrame, name: str):
        if df_s.empty:
            return
        df_s.to_excel(writer, sheet_name=name, index=False)
        ws = writer.sheets[name]
        for i, col in enumerate(df_s.columns):
            ws.set_column(i, i, WIDTHS.get(col, 20))
            ws.write(0, i, col, hdr)
        # Clickable job links
        if "Job Link" in df_s.columns:
            li = list(df_s.columns).index("Job Link")
            for ri, url in enumerate(df_s["Job Link"], start=1):
                if str(url).startswith("http"):
                    ws.write_url(ri, li, str(url), lnk, str(url)[:50])
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(df_s), len(df_s.columns) - 1)

    # Sheet 1 — All Jobs
    _write_sheet(df, "All Jobs")

    # Sheet 2 — By Platform
    for platform in df["Platform"].dropna().unique():
        _write_sheet(df[df["Platform"] == platform].copy(),
                     str(platform)[:28])

    # Sheet 3 — Summary
    ws3 = wb.add_worksheet("📊 Summary")
    ws3.set_column("A:A", 32)
    ws3.set_column("B:D", 22)
    ws3.merge_range("A1:D1",
                    "🤖 Salesforce Remote Job Hunt — Daily Summary", ttl)
    ws3.write("A2",
              f"Run: {datetime.now().strftime('%Y-%m-%d %H:%M')} | "
              f"Source: Apify openclawai/job-board-scraper | "
              f"Filter: Salesforce roles + Remote only", sub)

    stats = [
        ("🎯 New Jobs Today",   len(df)),
        ("🏢 Unique Companies", df["Company Name"].nunique()),
        ("💰 With Salary",
         (df["Salary"] != "Not Listed").sum() if "Salary" in df.columns else 0),
        ("📡 Platforms Used",   df["Platform"].nunique()),
    ]
    for ci, (lbl_txt, v) in enumerate(stats):
        ws3.write(3, ci, lbl_txt, lbl)
        ws3.write(4, ci, v, big)

    ws3.write("A7", "Platform",  lbl)
    ws3.write("B7", "New Jobs",  lbl)
    ws3.write("C7", "% of Total", lbl)
    for ri, (p, c) in enumerate(
            df["Platform"].value_counts().items(), start=7):
        ws3.write(ri, 0, p, val)
        ws3.write(ri, 1, c, val)
        ws3.write(ri, 2, f"{c/len(df)*100:.1f}%", val)

    ws3.write(7 + df["Platform"].nunique() + 1, 0,
              "Filters applied", lbl)
    filters = [
        "✅ Salesforce roles only (title whitelist + blacklist)",
        "✅ Fully remote jobs only (API flag + location/description check)",
        "✅ New jobs only (cross-run deduplication via seen_jobs.json)",
        f"✅ Posted within last {HOURS_OLD} hours",
    ]
    for ri, f in enumerate(filters, start=7 + df["Platform"].nunique() + 2):
        ws3.write(ri, 0, f)

    writer.close()
    print(f"  ✅ Saved → {os.path.abspath(filename)}")


# ==============================================================
#   📚  UPDATE MASTER FILE  (all-time accumulation)
# ==============================================================

def update_master(new_df: pd.DataFrame):
    # ── Check if master file is locked (open in Excel) ────────
    if os.path.exists(MASTER_FILE):
        lock_file = os.path.join(
            os.path.dirname(os.path.abspath(MASTER_FILE)),
            "~$" + os.path.basename(MASTER_FILE)
        )
        if os.path.exists(lock_file):
            backup = MASTER_FILE.replace(".xlsx", f"_backup_{TIMESTAMP}.xlsx")
            print(f"\n  ⚠️  '{MASTER_FILE}' is open in Excel — cannot overwrite.")
            print(f"  ℹ️  Close Excel and re-run to update the master file.")
            print(f"  💾 Saving today's new jobs to backup: {backup}")
            save_to_excel(new_df, backup)
            return

    if os.path.exists(MASTER_FILE):
        try:
            old = pd.read_excel(MASTER_FILE, sheet_name="All Jobs")
            combined = pd.concat([old, new_df], ignore_index=True)
            combined.drop_duplicates(
                subset=["Job Title", "Company Name", "Platform"],
                keep="first", inplace=True)
            combined.sort_values("Date Posted",
                                 ascending=False, inplace=True)
            print(f"\n  📚 Master: {len(old)} existing + "
                  f"{len(new_df)} new = {len(combined)} total")
            save_to_excel(combined, MASTER_FILE)
        except PermissionError:
            backup = MASTER_FILE.replace(".xlsx", f"_backup_{TIMESTAMP}.xlsx")
            print(f"\n  ⚠️  Permission denied writing '{MASTER_FILE}'.")
            print(f"  ℹ️  Close the file in Excel, then re-run.")
            print(f"  💾 Saving today's new jobs to backup: {backup}")
            save_to_excel(new_df, backup)
        except Exception as e:
            print(f"  ⚠️  Master update error: {e}")
            save_to_excel(new_df, MASTER_FILE)
    else:
        print("\n  📚 Creating master file for the first time...")
        save_to_excel(new_df, MASTER_FILE)


# ==============================================================
#   ☁️  SAVE TO GOOGLE SHEETS  (cloud mode)
#   Appends new rows — never deletes or overwrites previous data
# ==============================================================

def save_to_sheets(df: pd.DataFrame):
    """
    Appends today's new jobs to the 'All Jobs' tab in Google Sheets.
    Creates the tab (with headers) on first run.
    Previous rows are never touched — pure append.
    """
    if not USE_GOOGLE_SHEETS or _gsheet is None:
        return

    print(f"\n  ☁️  Saving {len(df)} jobs → Google Sheets '{GSHEETS_ALL_JOBS_TAB}'")

    try:
        try:
            ws = _gsheet.worksheet(GSHEETS_ALL_JOBS_TAB)
        except gspread.WorksheetNotFound:
            ws = _gsheet.add_worksheet(GSHEETS_ALL_JOBS_TAB, rows=50000, cols=20)
            ws.append_row(list(df.columns))   # write header on first run
            print(f"  ☁️  Created '{GSHEETS_ALL_JOBS_TAB}' tab (first run)")

        # If tab exists but is empty, write the header first
        if not ws.get_all_values():
            ws.append_row(list(df.columns))

        # Convert all values to strings (Sheets doesn't accept NaN/None)
        rows = df.fillna("").astype(str).values.tolist()
        ws.append_rows(rows, value_input_option="USER_ENTERED")
        print(f"  ✅ Appended {len(rows)} rows → "
              f"https://docs.google.com/spreadsheets/d/{_GSHEETS_SHEET_ID}")

    except Exception as e:
        print(f"  ⚠️  Google Sheets save error: {e}")
        print(f"  💾 Falling back to local Excel: {OUTPUT_FILE}")
        save_to_excel(df, OUTPUT_FILE)


# ==============================================================
#   🚀  MAIN
# ==============================================================

def main():
    t0 = datetime.now()

    mode_label = "☁️  CLOUD (Google Sheets)" if USE_GOOGLE_SHEETS \
                 else "💻 LOCAL (Excel)"
    print("=" * 62)
    print("   🤖 JOB HUNTING AI AUTOMATION TOOL  v2.0")
    print(f"   📡 Apify · openclawai/job-board-scraper")
    print(f"   {mode_label}")
    print("=" * 62)
    print(f"  🌐 Platforms   : {', '.join(SITES)}")
    print(f"  📍 Locations   : {len(LOCATIONS)}")
    print(f"  🔍 Keywords    : {len(SEARCH_KEYWORDS)}")
    print(f"  🕒 Posted ≤    : {HOURS_OLD} hours ago")
    print(f"  🏠 Remote only : YES")
    print(f"  🎯 Role filter : Salesforce only")
    print("=" * 62)

    # ── Init Google Sheets connection (cloud mode only) ───────
    if USE_GOOGLE_SHEETS:
        init_gsheets()

    # ── Load seen-jobs tracker ────────────────────────────────
    tracker = SeenJobsTracker(SEEN_JOBS_FILE)

    # ── Batch keywords (≤5 per call) ─────────────────────────
    kw_batches = batch_keywords(SEARCH_KEYWORDS, size=5)
    total_calls = len(kw_batches) * len(LOCATIONS)
    call_n = 0

    raw_all: list[dict] = []

    for batch in kw_batches:
        for location in LOCATIONS:
            call_n += 1
            print(f"\n  [{call_n}/{total_calls}]", end="")
            jobs = fetch_jobs_batch(batch, location)
            raw_all.extend(jobs)
            if call_n < total_calls:
                time.sleep(2)   # small courtesy delay

    print(f"\n{'─' * 62}")
    print(f"  📦 Raw jobs from API   : {len(raw_all)}")

    # ── Apply all 3 filters + clean ──────────────────────────
    kept, skipped_role, skipped_remote, skipped_dup = [], 0, 0, 0

    for job in raw_all:
        title = job.get("title", "") or ""

        # check each filter stage for stats
        if not is_salesforce_role(title):
            skipped_role += 1
            continue
        if not is_truly_remote(job):
            skipped_remote += 1
            continue
        if tracker.is_seen(job):
            skipped_dup += 1
            continue

        cleaned = filter_and_clean(job, tracker)
        if cleaned:
            kept.append(cleaned)

    print(f"  🚫 Skipped (wrong role) : {skipped_role}")
    print(f"  🚫 Skipped (not remote) : {skipped_remote}")
    print(f"  🚫 Skipped (duplicate)  : {skipped_dup}")
    print(f"  ✅ New Salesforce remote : {len(kept)}")

    if not kept:
        print("\n  ℹ️  No new jobs found today. "
              "All matching jobs were already collected.")
        tracker.save()
        return

    df = pd.DataFrame(kept)
    df.drop_duplicates(
        subset=["Job Title", "Company Name", "Platform"],
        keep="first", inplace=True)
    df.sort_values("Date Posted", ascending=False, inplace=True)
    df.reset_index(drop=True, inplace=True)

    # ── Save outputs ──────────────────────────────────────────
    if USE_GOOGLE_SHEETS:
        save_to_sheets(df)      # append to Google Sheets
        tracker.save()          # append new seen entries to Sheets
    else:
        save_to_excel(df, OUTPUT_FILE)
        update_master(df)
        tracker.save()          # overwrite seen_jobs.json

    # ── Final summary ─────────────────────────────────────────
    secs = (datetime.now() - t0).seconds
    print(f"\n{'=' * 62}")
    print(f"  ✅ Done in {secs}s")
    if USE_GOOGLE_SHEETS:
        print(f"  ☁️  Data → Google Sheets (appended, never overwritten)")
        print(f"  🔗 https://docs.google.com/spreadsheets/d/{_GSHEETS_SHEET_ID}")
    else:
        print(f"  📊 Today   : {OUTPUT_FILE}")
        print(f"  📚 Master  : {MASTER_FILE}")
    print(f"  🎯 New jobs: {len(df)}")
    print(f"\n  📡 Breakdown by platform:")
    for p, c in df["Platform"].value_counts().items():
        print(f"     {p:<18} : {c}")
    print("=" * 62)


if __name__ == "__main__":
    main()
