# ==============================================================
#   APIFY CLOUD SCHEDULER
#   ─────────────────────────────────────────────────────────
#   ⚠️  IMPORTANT — what this script does vs doesn't do:
#
#   ✅ Schedules the RAW openclawai/job-board-scraper on Apify
#      cloud. Raw results go to your Apify dataset.
#
#   ❌ Does NOT run job_hunter.py (filtering / dedup / Sheets).
#      For full automation use GitHub Actions (see README).
#
#   Use this file only if you want Apify to independently
#   keep a raw dataset of jobs in your Apify console.
# ==============================================================

import os
import requests
import json

# ── Load .env for local use ────────────────────────────────────
_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
if os.path.exists(_env_path):
    with open(_env_path, encoding="utf-8") as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())

# ── YOUR SETTINGS ──────────────────────────────────────────────
APIFY_API_TOKEN = os.getenv("APIFY_API_TOKEN", "")

# Cron schedule examples:
#   "0 8 * * *"     = Every day at 8:00 AM
#   "0 8 * * 1-5"   = Monday–Friday at 8:00 AM
#   "0 9,17 * * *"  = Twice daily (9 AM and 5 PM)
CRON_SCHEDULE = "0 13 * * 1-5"   # Mon–Fri at 8 AM EST (13:00 UTC)

TIMEZONE = "America/New_York"

# Actor input — mirrors job_hunter.py configuration
ACTOR_INPUT = {
    "searchTerms"         : [            # ✅ array (was single string — bug fixed)
        "Salesforce Engineer",
        "Salesforce Developer",
        "Salesforce Architect",
        "Salesforce Administrator",
        "Lead Salesforce",
    ],
    "location"            : "United States",
    "sites"               : ["indeed", "linkedin", "glassdoor",
                              "google", "zip_recruiter"],
    "maxResults"          : 20,
    "hoursOld"            : 24,
    "isRemote"            : True,        # ✅ was False — bug fixed
    "jobType"             : "fulltime",
    "countryIndeed"       : "usa",
    "enforceAnnualSalary" : True,
    "descriptionFormat"   : "markdown",
}


def create_apify_schedule():
    """Creates a schedule on Apify cloud — raw scraper runs without your PC."""

    url     = "https://api.apify.com/v2/schedules"
    headers = {
        "Authorization": f"Bearer {APIFY_API_TOKEN}",
        "Content-Type" : "application/json",
    }

    payload = {
        "name"          : "DailySalesforceJobScraper",
        "cronExpression": CRON_SCHEDULE,
        "timezone"      : TIMEZONE,
        "isEnabled"     : True,
        "isExclusive"   : False,
        "description"   : "Raw Salesforce job scrape — Mon-Fri 8 AM EST",
        "actions"       : [
            {
                "type"    : "RUN_ACTOR",
                "actorId" : "openclawai/job-board-scraper",
                "input"   : ACTOR_INPUT,
                "options" : {
                    "build"        : "latest",
                    "timeoutSecs"  : 300,
                    "memoryMbytes" : 1024,
                }
            }
        ]
    }

    response = requests.post(url, headers=headers, json=payload)

    if response.status_code == 201:
        data        = response.json()
        schedule_id = data["data"]["id"]
        print(f"\n  ✅ Schedule CREATED successfully!")
        print(f"  🆔 Schedule ID : {schedule_id}")
        print(f"  🕒 Runs        : {CRON_SCHEDULE} ({TIMEZONE})")
        print(f"  🌐 View at     : https://console.apify.com/schedules/{schedule_id}")
        print(f"\n  ⚠️  Note: this only schedules the raw scraper.")
        print(f"  For full automation (filter+dedup+Sheets) use GitHub Actions.")
        return schedule_id
    else:
        print(f"\n  ❌ Failed: {response.status_code}")
        print(f"  {response.text}")
        return None


def list_schedules():
    """List all your existing Apify schedules."""
    url     = "https://api.apify.com/v2/schedules"
    headers = {"Authorization": f"Bearer {APIFY_API_TOKEN}"}
    r       = requests.get(url, headers=headers)

    if r.status_code == 200:
        schedules = r.json()["data"]["items"]
        print(f"\n  📋 Your Apify Schedules ({len(schedules)} total):")
        for s in schedules:
            status = "✅ Active" if s.get("isEnabled") else "⏸️  Paused"
            print(f"  {status} | {s['name']} | {s.get('cronExpression')}")
    else:
        print(f"  ❌ Error: {r.status_code}")


def delete_schedule(schedule_id: str):
    """Delete a schedule by ID."""
    url     = f"https://api.apify.com/v2/schedules/{schedule_id}"
    headers = {"Authorization": f"Bearer {APIFY_API_TOKEN}"}
    r       = requests.delete(url, headers=headers)
    if r.status_code == 204:
        print(f"  ✅ Schedule {schedule_id} deleted")
    else:
        print(f"  ❌ Error: {r.status_code}")


if __name__ == "__main__":
    print("=" * 58)
    print("   ☁️  APIFY RAW SCRAPER SCHEDULER")
    print("=" * 58)
    create_apify_schedule()
    print()
    list_schedules()
