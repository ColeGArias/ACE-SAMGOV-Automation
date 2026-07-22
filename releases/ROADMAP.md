# Desktop App Roadmap — SAM.gov Award Collector

Turns the current `sam_automation.py` into a single `.exe` any user on the team can double-click. This version supersedes the first draft: while reviewing it, a direct check of SAM.gov's own Terms of Use turned up a real compliance problem with the *current* data-collection technique, and a viable, sanctioned replacement. Fixing that comes first — packaging a Terms-of-Use violation into an app and handing it to more people would be moving in the wrong direction, not "legitimizing" anything.

## Why this changed: the current technique isn't compliant

Fetched directly from [SAM.gov's Terms of Use](https://sam.gov/about/terms-of-use):

> "Automated data gathering, web scraping tools are prohibited and, if detected, will result in the associated account(s) being denied access to SAM.gov via Login.gov."
>
> "Do not use your SAM.gov login for data mining, bots, or other data gathering and extraction tools."

`sam_automation.py`'s `bootstrap_session()` connects over Chrome DevTools Protocol to a human-logged-in Chrome window and lifts its session cookie to drive SAM.gov's internal, undocumented search API (`sgs/v1/search` + `opps/v2/opportunities`). The fact that a real human logs in first doesn't exempt what happens after — that's still "using your SAM.gov login for... data gathering and extraction tools." The stated consequence is the account being denied SAM.gov/Login.gov access, not a CFAA matter (that clause covers unauthorized *modification*; this pipeline only ever reads).

**The fix**: SAM.gov publishes a real, sanctioned public API that appears to cover the same data — see Phase 0 below. Once that's validated, the CDP/Chrome dependency and this compliance risk both go away in the same move.

---

## Phase 0: Migrate the data source (do this before any app packaging)

Replace the internal-endpoint/CDP approach with SAM.gov's official [Get Opportunities Public API](https://open.gsa.gov/api/get-opportunities-public-api/) (`https://api.sam.gov/opportunities/v2/search`). It maps cleanly onto what `sam_automation.py` already does:

| Today (internal endpoint, via CDP) | Public API |
|---|---|
| `sgs/v1/search` with `naics=`, `notice_type=a`, `modified_date.from/to` | Same search, via `ncode=` (NAICS), `ptype=a` ("Award Notice"), `postedFrom`/`postedTo` (max 1-year span, `MM/dd/yyyy`) |
| `opps/v2/opportunities/{notice_id}` per-candidate detail call for `data2.award.{date,number,amount,awardee}` | Search response reportedly already includes `data.award.number` / `.amount` / `.date` / `.awardee.name` inline — **confirm this during prototyping**; if true, `fetch_award_detail`'s per-record lookup loop can be deleted entirely |
| Cookie lifted from a human-logged-in Chrome window | A personal `api_key`, generated from the user's own SAM.gov **Account Details** page |
| No documented rate limit (relies on evading bot detection) | Reported ~10 requests/day on an unlinked personal key, ~1,000/day if the account is linked to ACE's registered SAM.gov entity. Real usage here is ~100/week typical, 200–300/week at peak a couple months a year — comfortably inside the entity tier even worst-case (one call per record), and inside the unlinked tier too if the inline-award-data question above resolves favorably |
| No key expiry | Individual API keys rotate every ~90 days |

**Checklist:**

* [ ] Confirm the SAM.gov account is actually linked to ACE's registered entity (Account Details page) — believed to be true, not yet confirmed.
* [ ] Generate a Public API key and prototype a single `ptype=a` + `ncode=` + `postedFrom/postedTo` call against a known historical week.
* [ ] Confirm whether `data.award.*` is populated inline on search results — decides whether the detail-lookup loop survives.
* [ ] Re-target the padded-lookback-window logic (currently built around `modified_date`) at `postedFrom/postedTo` instead — same shape as today: search a wider window, then keep only records whose real `award.date` falls in the target Monday–Sunday week.
* [ ] Cross-check output against a real historical run of the current script (via the existing `--reference-date` flag) to confirm parity before cutting over.
* [ ] Store the API key in the app's config (Phase 3), same treatment as the SMTP password.
* [ ] Once validated, delete: `bootstrap_session()`'s CDP logic, `launch_chrome_debug.bat`, and `test_sam_scrape.py` (it diagnoses the anonymous-session approach being retired, and stops being useful once it's gone).

### Key-expiry UX

When the stored key stops working, the app shows a modal: *"Your SAM.gov API key needs updating — paste the new key from your Account Details page,"* a single input field, Save. Saving writes the new key to config and the app resumes normally — this repeats every ~90 days indefinitely, no code changes needed each cycle.

* **Reactive trigger (build this first)**: the API call fails with a 401/403 → show the modal immediately instead of a raw error.
* **Proactive trigger (nice-to-have)**: store the date the current key was entered; once real usage confirms whether rotation is a hard cutover or has a grace period, add a soft "expiring soon" banner a few days ahead of the ~90-day mark.

### Bonus this unlocks: real unattended scheduling

The current script requires a human to already be logged into a Chrome window before a scheduled run can succeed, which makes `run_sam_automation.bat` + Task Scheduler a fragile "stretch goal." Once the CDP/Chrome dependency is gone, that constraint disappears — a scheduled weekly run becomes a fully supportable, unattended path, not a workaround.

### Open items — unconfirmed, not blockers to this plan

* Whether search results embed award data inline, or a per-notice detail call is still required (drives real request volume).
* Exact current daily rate limits per account tier — SAM.gov's own docs page is vague; secondary sources gave the ~10/~1,000 figures above. Confirm once a key exists.
* Exact key-rotation mechanics (advance email warning? grace period? silent regeneration?).

---

## Repo structure cleanup

Current layout is everything flat at the repo root, plus a couple of leftovers found while working through this: `.env.example` and `env.example` are byte-identical duplicates, and `test.py` at the root is one unrelated junk line, not real code. Proposed layout, grouped by concern:

```text
ACE-SAMGOV-Automation/
├── sam_automation.py
├── requirements.txt
├── requirements-app.txt       # new: eel, pyinstaller, pyinstaller-hooks-contrib (Phase 2)
├── README.md
├── .env.example                # env.example (duplicate) removed
├── .gitignore
├── output/                     # sam_construction_awards.xlsx (master) + per-run dated files
├── logs/                       # sam_automation.log
├── scripts/                    # run_sam_automation.bat (launch_chrome_debug.bat retired in Phase 0)
├── app/                        # Eel UI + PyInstaller entry point
│   ├── main.py
│   ├── config_store.py
│   └── web/
│       ├── index.html
│       ├── styles.css
│       └── app.js
└── releases/
    └── ROADMAP.md
```

* [ ] Delete `test.py`.
* [ ] Delete the duplicate `env.example`, keep `.env.example`.
* [ ] Move `sam_construction_awards.xlsx` → `output/`, `sam_automation.log` → `logs/`; update `EXCEL_OUTPUT_PATH` / `LOG_PATH` defaults to match.
* [ ] Move `run_sam_automation.bat` → `scripts/`; delete `launch_chrome_debug.bat` and `test_sam_scrape.py` once Phase 0 lands.

---

## New capability: per-run downloadable dated spreadsheet

In addition to the existing ever-growing master workbook (kept exactly as-is), each run now also produces a standalone file for just that run's date range.

* [ ] Write a second file per run to `output/`, named `SAM_Construction_Awards_{start}_to_{end}.xlsx` (ISO dates — e.g. `SAM_Construction_Awards_2026-06-29_to_2026-07-05.xlsx`), alongside appending to the master workbook as today.
* [ ] Auto-save this file locally (so it's never lost even if a browser download is missed or misplaced) **and** offer it as a real browser download from the app.
* [ ] Download mechanism: Eel has no built-in file-download route, so the exposed Python function reads the file, base64-encodes it, and returns the bytes to JS over Eel's normal call bridge. `app.js` turns that into a `Blob` and triggers a synthetic `<a download="SAM_Construction_Awards_...xlsx">` click — the standard client-side "download a blob" pattern, no custom HTTP routes needed.

---

## Corrected Architecture Overview

```text
User Double-Clicks SAM-Award-Collector.exe
  ├── 1. PyInstaller extracts bundled web/ assets to a temp path (sys._MEIPASS)
  ├── 2. Python launches the Eel background server & opens the app window
  ├── 3. On first launch, UI reads/creates per-user config at %LOCALAPPDATA%\ACE-SAMGOV-Automation\config.json
  │      (NAICS codes, Excel/output settings, email settings, SAM.gov API key)
  ├── 4. User reviews/edits settings and clicks Run (or Dry Run)
  ├── 5. JS calls an exposed Python function; Python runs the pipeline against
  │      api.sam.gov on a background thread (imported directly — no subprocess, no .bat)
  ├── 6. Python streams progress lines back to the UI via an exposed JS callback
  ├── 7. On success: master workbook updated, a new dated per-run file saved to
  │      output/ and offered as a one-click browser download
  └── 8. If the stored API key has expired, a modal prompts for a fresh one
         (see Phase 0's key-expiry UX) instead of surfacing a raw auth error
```

---

## Directory Structure (app-specific detail)

```text
ACE-SAMGOV-Automation/
├── sam_automation.py         # unchanged in shape — imported as a module; internals updated per Phase 0
├── requirements.txt
├── requirements-app.txt
│
├── app/
│   ├── main.py                # Eel entry point / PyInstaller build target
│   ├── config_store.py        # reads/writes %LOCALAPPDATA%\ACE-SAMGOV-Automation\config.json
│   ├── resource_path.py        # sys._MEIPASS helper for *read-only bundled* assets only (web/)
│   └── web/
│       ├── index.html          # settings form + run/dry-run buttons + log pane + download button
│       ├── styles.css
│       └── app.js
│
├── output/                     # master workbook + per-run dated files (writable, not bundled)
├── logs/
├── scripts/
│   └── run_sam_automation.bat  # Task Scheduler entry point — now fully supportable, see Phase 0
└── releases/
    └── ROADMAP.md
```

---

## Phase 1: Environment & Directory Setup

* [ ] Add `app/` alongside the existing files per the structure above.
* [ ] Reuse the existing `.venv` (already present in this repo) rather than creating a second one.
* [ ] Add app-only dependencies in a separate file so the CLI-only `requirements.txt` doesn't grow deps CLI-only users don't need:
  ```bash
  pip install eel pyinstaller pyinstaller-hooks-contrib
  pip freeze > requirements-app.txt
  ```

---

## Phase 2: Frontend Development (Settings + Status UI)

* [ ] `app/web/index.html` — a form for: `NAICS_CODES`, `MODIFIED_DATE_LOOKBACK_DAYS` (recheck this default once Phase 0 re-targets it at `postedFrom/postedTo`), `EXCEL_OUTPUT_PATH` (native file-save picker), SAM.gov API key, and the "Email notifications" section (`EMAIL_ENABLED`, `SMTP_HOST/PORT/SSL`, `EMAIL_ADDRESS`, `NOTIFY_EMAIL_TO`; password handled via `keyring`, Phase 6).
* [ ] "Run Now" / "Dry Run" buttons (the existing `--dry-run` flag becomes a checkbox).
* [ ] A "Download this run's report" button that appears once a run completes (see the per-run downloadable spreadsheet section above).
* [ ] A scrolling log pane fed by streamed progress (Phase 3) — a multi-minute run with a silent UI reads as "frozen" to a non-technical user.
* [ ] The key-expiry modal from Phase 0.
* [ ] Include the Eel bridge script tag:
  ```html
  <script type="text/javascript" src="/eel.js"></script>
  ```

---

## Phase 3: Backend Bridge (Eel & Threading)

* [ ] `app/main.py` initializes Eel pointing at `web/`:
  ```python
  import eel
  eel.init('web')
  ```
* [ ] `app/resource_path.py` — the `sys._MEIPASS` helper, scoped explicitly to read-only bundled assets (`web/`) only. Never used for config, output files, or logs — those live in writable locations (`%LOCALAPPDATA%`, `output/`, `logs/`).
* [ ] Run the pipeline on a background thread — a synchronous multi-minute call would block Eel's single process and freeze the window:
  ```python
  import threading
  import sam_automation

  @eel.expose
  def run_pipeline(options, dry_run):
      threading.Thread(target=_run_pipeline_worker, args=(options, dry_run), daemon=True).start()
      return {"status": "started"}
  ```
* [ ] Stream progress via a logging handler that forwards each record to an exposed JS function, instead of returning one final blob at the end:
  ```python
  class EelLogHandler(logging.Handler):
      def emit(self, record):
          eel.append_log_line(self.format(record))
  ```
* [ ] Expose a `download_run_file(path)` function that base64-encodes the per-run file for the download flow described above.

---

## Phase 4: Config & Credentials

* [ ] `app/config_store.py` reads/writes JSON at `%LOCALAPPDATA%\ACE-SAMGOV-Automation\config.json`, created with defaults matching `env.example` on first launch. Fully replaces `.env` / `load_dotenv()` for the packaged app; the CLI script keeps using `.env` unchanged.
* [ ] Store the SAM.gov API key and SMTP password via `keyring` (Windows Credential Manager) rather than plaintext JSON — this is going to other people's machines.
* [ ] Call `sam_automation`'s existing functions directly from the worker thread — no subprocess, no `.bat`, no shelling out for the actual data pipeline.

---

## Phase 5: Packaging into a Standalone Executable

* [ ] Test the full pipeline end-to-end in dev mode (`python app/main.py`) before packaging.
* [ ] Package with PyInstaller, bundling only the read-only `web/` assets:
  ```bash
  pyinstaller app/main.py --name SAM-Award-Collector --onedir --noconsole --add-data "app/web;web"
  ```
* [ ] Because `--noconsole` hides all stdout/stderr, wrap `main()`'s startup in a `try/except` that shows a native error dialog (e.g. `ctypes.windll.user32.MessageBoxW`) on failure.
* [ ] Test on a machine with no Python installed (confirms the freeze worked).

---

## Phase 6: Rollout & Multi-User Hardening

* [ ] **Decide Excel output location for multi-user use**: local per-user file (today's default) vs. a shared network path so multiple people's runs accumulate into one workbook — a product decision, not a technical one.
* [ ] **Unsigned-exe warning**: an unsigned PyInstaller `.exe` triggers Windows SmartScreen for other users. Get it signed with an org code-signing cert, or document the "More info → Run anyway" click-through.
* [ ] **Versioning**: put a version string in the window title/UI so support requests can identify which build someone is running.
* [ ] Keep the CLI flow (`.env`, `sam_automation.py` run directly) documented and working until the packaged app has been used successfully by at least one other person.

---

## Documentation & Reference Links

| Resource | Description | URL |
|---|---|---|
| **SAM.gov Terms of Use** | Source of the "no automated data gathering" finding that drove Phase 0 | https://sam.gov/about/terms-of-use |
| **SAM.gov Get Opportunities Public API** | Official docs for the endpoint Phase 0 migrates to | https://open.gsa.gov/api/get-opportunities-public-api/ |
| **Eel Repository** | Official Python-Eel GitHub repo & examples | https://github.com/python-eel/Eel |
| **PyInstaller Manual** | Official PyInstaller documentation | https://pyinstaller.org/en/stable/ |
| **PyInstaller Data Files** | Guide on bundling external assets (`--add-data`) | https://pyinstaller.org/en/stable/spec-files.html#adding-data-files |
| **PyInstaller Runtime Information** | Resolving `sys._MEIPASS` at runtime | https://pyinstaller.org/en/stable/runtime-information.html |
| **keyring (PyPI)** | Cross-platform OS credential-store access — used for the SAM.gov API key and SMTP password | https://pypi.org/project/keyring/ |
| **Python subprocess** | Reference for any remaining process-launch needs | https://docs.python.org/3/library/subprocess.html |
