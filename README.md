# ACE-SAMGOV-Automation

Weekly collector that pulls newly-signed construction contract awards from SAM.gov and appends them to a running Excel workbook for business development.

Run manually, or on a schedule (Windows Task Scheduler / cron) with no arguments for a normal weekly run.

## Quick Start

One-time setup (skip anything already done — e.g. `.env` may already exist):
```
pip install -r requirements.txt
playwright install chromium
```
Copy `env.example` to `.env` and fill in values if you haven't already.

Every time you want to run it:
1. Run `launch_chrome_debug.bat` and leave the window open. First time only: log into SAM.gov in that window — the session is saved for future runs.
2. Run `python sam_automation.py` (add `--dry-run` to preview without writing to Excel or sending email).

For unattended/scheduled runs, point Task Scheduler at `run_sam_automation.bat` instead of step 2 — but the Chrome debug window from step 1 must already be running and logged in.

See [Setup](#setup) below for more detail.

## Configuration

All configuration is via environment variables, loaded from a `.env` file in the project root (copy `env.example` to `.env` and fill in values — `.env` is gitignored and must never be committed).

| Variable | Default | Purpose |
|---|---|---|
| `NAICS_CODES` | `23,237,2362,2373,2379` | Comma-separated NAICS codes/prefixes to filter the Construction family search. |
| `MODIFIED_DATE_LOOKBACK_DAYS` | `30` | How far before the target week to widen the search window (see [Why the date window is padded](#why-the-date-window-is-padded)). |
| `CDP_URL` | `http://localhost:9222` | Chrome DevTools Protocol URL of the already-running, already-logged-in Chrome window the script attaches to. |
| `EXCEL_OUTPUT_PATH` | `./sam_construction_awards.xlsx` | Workbook that award rows are appended to. Created on first run if missing. |
| `EMAIL_ENABLED` | `false` | Send a success/failure notification email after each run. |
| `SMTP_HOST` | `smtp.gmail.com` | SMTP server (works with Gmail, Microsoft 365 / ace-consulting.net, or any SMTP provider). |
| `SMTP_PORT` | `465` | SMTP port. |
| `SMTP_USE_SSL` | `true` | Use implicit SSL (`SMTP_SSL`) vs. STARTTLS. |
| `EMAIL_ADDRESS` | — | Sending account (also used as SMTP login). |
| `EMAIL_PASSWORD` | — | SMTP password / app password for the sending account. |
| `NOTIFY_EMAIL_TO` | — | Recipient address for the notification. |
| `LOG_LEVEL` | `INFO` | Python logging level. |
| `LOG_PATH` | `./sam_automation.log` | Log file location; also mirrored to stdout. |

Command-line flags (`sam_automation.py`):

| Flag | Purpose |
|---|---|
| `--reference-date YYYY-MM-DD` | Compute the target award week relative to this date instead of today. Used for backfills/testing. |
| `--dry-run` | Fetch and parse records but skip the Excel write and the email send. |

## What it does

Every run computes the **previous full Monday–Sunday week** relative to today (or `--reference-date`), searches SAM.gov for Construction-NAICS contract awards signed in that week, and appends one Excel row per award with:

`Awardee, Title, Owner, Project Value, Contract Award Date, Contract Award Number, Description`

New rows are appended to the existing workbook at `EXCEL_OUTPUT_PATH` rather than overwriting it, so the file accumulates history run over run.

## How it works

### Session bootstrap: a real, human-logged-in Chrome window

SAM.gov's data comes from its own **internal search UI API** (`sgs/v1/search` + `opps/v2/opportunities`) — the same calls the SAM.gov website itself makes — not the public `api.sam.gov` REST API. The public API with a standard key proved unreachable during development, so this script instead reverse-engineers the exact browser workflow used to pull this data manually.

That internal API also appears to detect and reject automation: SAM.gov's session-setup call never fires for a browser Playwright launches itself. To work around this, `bootstrap_session()` does **not** launch its own browser. Instead it connects over the Chrome DevTools Protocol (`playwright.chromium.connect_over_cdp`) to a Chrome window you start and log into yourself, and lifts the `XSRF-TOKEN`/`SESSION` cookie from that genuinely human-initiated session. `launch_chrome_debug.bat` starts that Chrome window with `--remote-debugging-port=9222` and a dedicated profile directory (`%LOCALAPPDATA%\sam_automation_chrome_profile`), so the SAM.gov login persists between runs — you only need to log in the first time.

Because the script attaches to a browser it doesn't own, it never calls `browser.close()` — closing it is the operator's responsibility.

### Fetch: bulk search, then per-record detail lookup

1. **`fetch_bulk_records`** pages through the search endpoint filtered by NAICS code and `modified_date`, and collects every candidate notice.
2. **`fetch_award_detail`** then calls the per-notice detail endpoint for each candidate to get the fields the search index doesn't return: the real contract award date, award number, and dollar amount.
3. Each candidate is kept only if its real award date falls inside the target Monday–Sunday week; everything else is discarded.

#### Why the date window is padded

SAM.gov's search index only supports filtering by `modified_date` (last-touched timestamp), not the actual award date. A notice awarded in the target week may have been last modified earlier or later than that week. To compensate, the search queries a window starting `MODIFIED_DATE_LOOKBACK_DAYS` (default 30) before the target week's start, then the per-record detail lookup narrows results down to notices whose *real* award date falls in-range. Increase this if awards near the boundary start going missing; decrease it to cut down on wasted detail-lookup calls.

### Parsing

`parse_bulk_record` extracts the awardee, title, owner (first entry in `organizationHierarchy`), and description (HTML-stripped and unescaped) from the search result. `merge_award_detail` then overlays the award date, award number, project value, and — if present — a more specific awardee name from the detail endpoint's `data2.award` block.

### Resilience

All HTTP calls to SAM.gov go through `_pw_get_json`, which retries on `429` (rate limited) and `5xx` responses with exponential backoff (via `tenacity`, up to 5 attempts, 2–60s). A failed detail lookup for a single notice is logged and skipped rather than aborting the whole run.

### Excel output

`load_or_create_workbook` creates a new workbook with a styled header row (`Awards` sheet, frozen header, column widths tuned per field) if `EXCEL_OUTPUT_PATH` doesn't exist yet; otherwise it opens the existing file and `append_records` adds new rows to the bottom. If the workbook is open in Excel at save time, this raises a clear `PermissionError` telling you to close it and re-run.

### Email notifications

If `EMAIL_ENABLED=true`, `send_email` sends a plain-text summary (rows added, workbook path, log path) on success, or the exception message on failure, over SMTP (SSL or STARTTLS depending on `SMTP_USE_SSL`). Disabled and silently skipped by default.

## Setup

1. **Install dependencies** (Python venv recommended):
   ```
   pip install -r requirements.txt
   playwright install chromium
   ```
2. **Configure**: copy `env.example` to `.env` and adjust values as needed.
3. **Start the Chrome debug window**: run `launch_chrome_debug.bat`. On first run, log into SAM.gov in the window it opens and leave it running. The session persists in its dedicated profile for future runs.
4. **Run the collector**: `python sam_automation.py` (add `--dry-run` to preview without writing Excel or sending email).
5. **Schedule it**: `run_sam_automation.bat` activates the project's venv and runs `sam_automation.py`, logging its exit code to `scheduler_wrapper.log`. Point Windows Task Scheduler at this `.bat` for unattended weekly runs — the Chrome debug window from step 3 must already be running and logged in for the scheduled run to succeed.

## Files

| File | Purpose |
|---|---|
| `sam_automation.py` | The production script — everything described above. |
| `launch_chrome_debug.bat` | Launches the human-driven Chrome window the script attaches to. |
| `run_sam_automation.bat` | Scheduler entry point: activates the venv, runs the script, logs the exit code. |
| `requirements.txt` | Python dependencies. |
| `env.example` | Template for `.env`. |
| `sam_construction_awards.xlsx` | Default output workbook (gitignored in practice; regenerated as needed). |
| `test_sam_scrape.py` | Standalone diagnostic script probing whether an anonymous (non-logged-in) session can reach the internal API directly — not part of the production flow. |

## Known fragility

This depends entirely on an **undocumented, internal SAM.gov API** (`sgs/v1/search`, `opps/v2/opportunities`) that mirrors what the website's own UI calls — it is not a supported public interface and could change or break without notice. If the script starts failing, checking whether these endpoints or their response shapes changed is the first debugging step.
