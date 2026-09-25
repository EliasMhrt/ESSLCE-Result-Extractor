# ESSLCE Result Extractor — Technical Explanation

An automated, CAPTCHA-solving API client for the Ethiopian Secondary School
Leaving Certificate Examination (ESSLCE) results portal, wrapped in a
drag-and-drop GUI for non-technical school staff at Tana Haik Secondary School.
The app is branded Oryares and donated to the school.

## Pipeline Overview

```
Start.bat
    └─> launcher.py  (GUI)  ──spawns subprocess──>  extractor.py (engine)
                                                          │
                                    ┌─────────────────────┼─────────────────────┐
                                    │                     │                     │
                              Chrome browser        High school students.xlsx  results.xlsx
                              solves Turnstile           (input)               (output)
                              + HMAC-signed API           │   progress.csv      │
                              calls via api.eaes.et       └── (resume/backup)   ┘
```

There is **no LLM / AI** anywhere in this app. Every step is deterministic
Python logic: column matching, zero-padding, HMAC signing, JSON parsing.

---

## 1. `Start.bat` — Invisible entry point

- `cd /d "%~dp0"` changes to the folder containing the scripts so all relative
  paths (`.venv`, `students.xlsx`, `results.xlsx`, …) resolve correctly.
- Launches `launcher.py` with `.venv\Scripts\pythonw.exe`.
  - `pythonw.exe` runs Python **without a console window**, so staff only see
    the GUI.
- Falls back to `python.exe` if `pythonw` is absent.
- If the venv is missing it prints an install hint and pauses so the window
  does not vanish instantly.

---

## 2. `launcher.py` — The GUI shell

### Two modes, decided in `main()` (launcher.py:373)

| Condition | Mode | What happens |
|---|---|---|
| no `-i` argument | GUI mode | Opens the tkinter window and runs the event loop |
| `-i` / `--input` present | Extraction mode | `from extractor import main; main()` — i.e. it *becomes* the engine |
| `--selftest` present | Self-test | Imports extractor to validate the import chain, writes `selftest_ok.txt` |

The dual role matters for the frozen PyInstaller EXE: the EXE *is* the launcher,
but when called with `-i <file>` it runs extraction. See `_runner_cmd()`
(launcher.py:26).

### GUI layout (4 cards)

1. **Students list card** — drag-and-drop zone (`tkinterdnd2`) + Browse.
   Dropped paths are stored in `self.input_path`.
2. **Output card** — where `results.xlsx` will be saved (default: same folder).
3. **Actions card** — Start (green), Stop (orange), Open Results buttons.
4. **Log card** — a read-only dark `tk.Text` widget showing live progress.

### Subprocess handshake — how the GUI stays responsive

The extraction is **CPU/socket heavy and slow**. It is deliberately run in a
child process so a hang or crash cannot freeze the GUI:

1. `_start()` (launcher.py:313) validates the chosen input file, then:
   ```
   subprocess.Popen([python, extractor.py, -i <input>, -o <output>],
                    stdout=PIPE, stderr=STDOUT, text=True,
                    bufsize=1, creationflags=CREATE_NO_WINDOW, ...)
   ```
   - `creationflags=CREATE_NO_WINDOW (0x08000000)` hides any console window.
   - `stderr=STDOUT` merges errors into the same stream.
   - `bufsize=1` gives **line-buffered** reading.
2. A **reader thread** blocks on `readline()` and pushes each line into a
   thread-safe `queue.Queue`.
3. `_drain_log()` is a periodic `root.after(120, …)` callback that:
   - drains the queue into the log widget, and
   - `poll()`s the child; when it exits, re-enables buttons.
   - exit code 0 → "Finished ✓" + enables **Open Results**
     (`os.startfile`); non-zero → "Stopped / error".

This *thread + queue + periodic poll* pattern is the standard way to keep a
tkinter UI responsive while a heavy worker runs.

---

## 3. `extractor.py` — The engine (core of the app)

### Step A — Read the student list (lines 101–156)

`read_students_file()`:
- `.xlsx` / `.xls` → `pandas.read_excel`.
- `.csv` / plain text → tries encodings in order:
  `utf-8-sig`, `utf-8`, `cp1252`, `latin-1`; if a UTF-16 BOM is detected,
  `utf-16` is tried first. Comma is assumed; if the file collapses to a single
  column containing tabs, it is re-read with `sep="\t"` (handles phone / Excel
  exports).

`find_columns()` → fuzzy header matching via `_match_key()`:
- `ADMISSION_KEYS` = admission, admission_no, candidate, reg, registration, id number
- `FIRSTNAME_KEYS` = first_name, first name, fname, name
- Matching is case-insensitive and partial (`startswith` or substring).
- **Any other columns (Section, Class, …) are carried through to the output**
  untouched (`extra_cols`).

`normalize_admission()` — **critical fix**:
- Strips non-digits, then **zero-pads to 8 digits**: `1234567 → 01234567`.
- Without this the API returns 404 "Student not found". Confirmed by the
  history in `debug.log`.

### Step B — Launch an undetected, visible Chrome (`main()` line ~464)

```
with SB(uc=True, test=True, incognito=True) as sb:
    sb.activate_cdp_mode(SITE_URL)
```

- SeleniumBase **UC (`uc=True`) mode** + CDP drives a *real* Chrome that looks
  indistinguishable from a human browser, so Cloudflare does not flag it.
- The window **must stay visible** — that is required for the CAPTCHA clicks to
  look human (a real mouse/pointer event).

### Step C — Solve Turnstile and harvest the token (`harvest_token`)

Cloudflare Turnstile protects the portal. The flow per student:
1. `sb.open(SITE_URL)`.
2. `sb.solve_captcha()` — SeleniumBase attempts to auto-solve the widget.
3. Poll the hidden token field (`_token_js`) until a `cf-turnstile-response`
   textarea/input contains a token > 40 chars.
4. If still empty, `sb.uc_gui_click_captcha()` physically clicks the checkbox.
5. Up to 4 attempts; at the call site, if every attempt fails, wait
   `BACKOFF_SECONDS` (20s) and retry forever until a token appears.

The token is stored in the module global `_HARVESTED_TOKEN`, refreshed **for
every student** because Turnstile tokens are single-use.

### Step D — Signed API call (`lookup_student`)

Instead of scraping HTML, the script replicates the portal's own XHR calls:

1. `GET https://api.eaes.et/api/v1/session/key` → `{session_id, session_secret}`.
2. Build `signature_input = f"{ms_timestamp}:{admission_lower}:{first_name_lower}"`.
3. Sign: **HMAC-SHA256** over the input with the session secret (`_sign()`).
4. `POST https://api.eaes.et/api/v1/results/web` with:
   - Headers: `X-Session-Id`, `X-Timestamp`, `X-Signature`,
     a real Chrome `User-Agent`, and `Origin`/`Referer` set to the portal —
     so the request is indistinguishable from the browser's.
   - JSON body: `admission_no`, `first_name` (UPPERCASED by the caller),
     and the harvested `turnstileToken`.
5. Every request/response is appended to `debug.log` (timestamp, status,
   first 600 chars of body) — invaluable for debugging.
6. Errors are classified:
   - **404** → "Student not found" — a *normal, expected* outcome.
   - **403 / 429** → "challenge_403/429" → caller harvests a fresh token and
     sleeps `BACKOFF_SECONDS`.
   - anything else / network → transient → retried (up to `LOOKUP_RETRIES=3`).

### Step E — Parse the JSON (`parse_result`)

Flattens the API response into one output row:

`studentInfo` → Full Name, School, Admission No, Stream, Sex, Age, Total, Average.

**Name splitting** (`split_name`):
- `"HANA TESFAYE GEBRE"` → **First Name** = HANA, **Middle Name** =
  TESFAYE, **Last Name** = GEBRE.
- Positional rule: token 0 → first; tokens 1..n-1 → middle; last token → last.
- 1-token and 2-token names degrade gracefully (missing parts are blank).

`results` array → **one column per subject** (Biology, Chemistry, Physics,
Mathematics, English, Scholastic Aptitude Test, …), so any/new subject the API
returns automatically appears in the output.

Key lookup is via `_find_key()` — case-insensitive with partial matching
(`Avg`, `AVERAGE`, `avg_score` all resolve). Numeric strings are converted to
int/float where clean.

### Step F — Main loop, safety rails, throttling (lines ~468–574)

For each student:
1. Skip if already completed — `load_done()` reads `progress.csv` and filters
   on `Admission|First Name` keys (status `Found` / `Not Found`).
2. Normalize admission + uppercase first name; build the row dict with the
   input First Name and all carried extra columns.
3. Harvest token (retrying forever with 20s backoff on failure).
4. Lookup with up to 3 retries; on a 403/429 *challenge* it harvests a **new**
   token and backs off 20s.
5. First success dumps the raw JSON to `raw_sample.json` (dev check).
6. Append the row with `Status` = `Found` / `Not Found` / `Error`.

Throttling — to stay under the site's radar:
- random 4–8s delay (`MIN_DELAY`/`MAX_DELAY`) between students;
- every 25 students, a longer 20–40s pause.

Resume / incremental save:
- Every 5 students, `save_progress()` **merges everything back into
  `progress.csv`** (deduped by `Admission|First Name` key). A killed run simply
  restarts and skips the already-done students.

### Final write-out

- `load_progress_rows()` reads back all cumulative rows (across runs).
- Merged into one DataFrame → `results.xlsx` via pandas + openpyxl.
- If the output file is locked (open in Excel/WPS), a timestamped fallback is
  written: `results_20260915_123456.xlsx` and the user is notified.
- Prints a final summary: count checked, count resumed, elapsed time.

---

## 4. The data files and their roles

| File | Producer | Consumer | Role |
|---|---|---|---|
| `students.xlsx` | school staff | `extractor.py` | input: identifies each student |
| `progress.csv` | `save_progress()` | `load_done()` / `load_progress_rows()` | resume checkpoint + cumulative raw backup |
| `results.xlsx` | final `to_excel` | teacher | the deliverable |
| `raw_sample.json` | first successful lookup | developer | verify all subjects are captured |
| `debug.log` | `_log_debug()` | developer | full request/response history for diagnostics |
| `downloaded_files/` | SeleniumBase / chromedriver setup | — | auto-managed marker + download cache |
| `selftest_ok.txt` | `--selftest` run | developer | proves the import chain works end-to-end |

---

## 5. Runtime requirements

- **Windows with a visible display and mouse** — the browser must stay visible.
- **Google Chrome** installed (SeleniumBase drives a real instance).
- **Python 3.9+** in a venv with `pandas`, `openpyxl`, `seleniumbase`,
  `tkinterdnd2` (+ `pyinstaller` for building).
- Internet access to `result.eaes.et` and `api.eaes.et`.
- Admission numbers resolve to **8 digits** after zero-padding.

## 6. Build / packaging

- `ESSLCE-Result-Extractor.spec` — PyInstaller single-file **windowed**
  (`console=False`) build; collects `seleniumbase`, `tkinterdnd2`, and `attrs`
  metadata so Chrome automation keeps working inside the frozen EXE.
- Building gives `dist\ESSLCE-Result-Extractor.exe`. That EXE is the launcher,
  and via the `-i` mode in `launcher.py:main()` it also acts as the engine.

---

## Summary

The GUI collects two file choices and launches the engine as a hidden child
process, streaming its stdout into a log panel. The engine opens a
visually-normal Chrome, solves a Turnstile CAPTCHA to obtain a one-time token,
replicates the official API's HMAC-SHA256 request signing to fetch clean JSON
per student, splits each full name into First / Middle / Last, throttles every
request to avoid rate-limiting, and saves incrementally to `progress.csv` so a
school-sized batch run is resumable and crash-safe. No machine learning is used
anywhere — every step is explicit, deterministic logic.