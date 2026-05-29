# 🤖 Job Hunting AI Automation Tool

Fetches jobs from **Indeed + LinkedIn + Glassdoor + Google + ZipRecruiter**
using **Apify cloud** — runs automatically without your computer.

---

## 📁 Files in This Project

| File | Purpose |
|---|---|
| `job_hunter.py` | Main script — fetches jobs & saves to Excel |
| `apify_scheduler.py` | Sets up cloud schedule on Apify |
| `README.md` | This guide |

---

## 🚀 Setup Guide (Step by Step)

### Step 1 — Install Python Libraries

Open PowerShell and run:
```powershell
pip install apify-client pandas openpyxl xlsxwriter requests
```

### Step 2 — Get Your Free Apify API Token

1. Go to 👉 https://apify.com — Sign up free (no credit card)
2. Go to 👉 https://console.apify.com/account/integrations
3. Copy your **Personal API Token**
4. Paste it in `job_hunter.py` where it says `YOUR_APIFY_API_TOKEN_HERE`

### Step 3 — Configure Your Search (in job_hunter.py)

Edit the configuration section at the top:
```python
SEARCH_KEYWORDS = ["Salesforce Engineer", "Lead Salesforce", ...]
LOCATIONS       = ["United States", "Remote", ...]
SITES           = ["indeed", "linkedin", "glassdoor", ...]
HOURS_OLD       = 72      # jobs posted in last 72 hours
REMOTE_ONLY     = False   # True = remote jobs only
```

### Step 4 — Run Manually (Test It)

```powershell
cd C:\Users\MS102 9926461\Documents\JobHunter
python job_hunter.py
```

### Step 5 — Schedule on Apify Cloud (No PC Needed)

**Option A: Apify Dashboard (Easiest)**
1. Go to 👉 https://console.apify.com/schedules
2. Click **"Create new"**
3. Set cron: `0 8 * * *` (runs every day at 8 AM)
4. Add Actor: `openclawai/job-board-scraper`
5. Paste your search input
6. Click **Enable** ✅
7. Done — Apify runs it daily forever without your PC

**Option B: Via Python Script**
```powershell
python apify_scheduler.py
```

---

## 📊 Excel Output Structure

| Sheet | Contents |
|---|---|
| **All Jobs** | Every job found, all columns, sortable |
| **Remote Only** | Only remote jobs, green highlighted |
| **Indeed** | Jobs from Indeed only |
| **Linkedin** | Jobs from LinkedIn only |
| **Glassdoor** | Jobs from Glassdoor only |
| **📊 Summary** | Stats: total, remote, by platform, by keyword |

### Columns in Excel

| Column | Example |
|---|---|
| Job Title | Senior Salesforce Engineer |
| Company Name | Accenture |
| Platform | Indeed / LinkedIn / Glassdoor |
| Location | New York, NY, US |
| Remote | ✅ Yes / ❌ No |
| Salary | $110,000 - $150,000 USD/year |
| Date Posted | 2026-05-21 |
| Job Link | https://indeed.com/job/... (clickable) |
| Company Rating | 4.2 |
| Skills Required | Salesforce, Apex, LWC |
| Description | First 400 chars of job description |

---

## 💰 Cost on Apify Free Plan

- Free plan: **$5 credit/month** (resets monthly)
- Actor cost: **$0.003 per job**
- Jobs per month FREE: **$5 ÷ $0.003 = ~1,666 jobs/month**

---

## 🔗 Official Links

- Apify Homepage: https://apify.com
- Apify Console: https://console.apify.com
- Actor Used: https://apify.com/openclawai/job-board-scraper
- Apify Pricing: https://apify.com/pricing
- API Token: https://console.apify.com/account/integrations
- Schedules: https://console.apify.com/schedules
