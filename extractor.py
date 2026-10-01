"""
ESSLCE Result Extractor - Oryares (Donated to Tana Haik Secondary School)
=====================================================
Checks official ESSLCE results on https://result.eaes.et for a list of
students and saves the results to an Excel file.

How it works
------------
1. A real (visible) Chrome browser opens the results portal in Stealth/UC mode.
2. Cloudflare Turnstile is solved automatically in the browser and a fresh
   security token is harvested.
3. The script then talks to the official results API directly using that token
   (with the same HMAC signing the website itself uses), so each student lookup
   returns clean structured JSON.
4. Results are written to an Excel file as they are found (incremental).

Requirements
------------
- Python 3.9+ on a real Windows machine with a mouse/display.
- The browser window must stay visible while the script runs.
- Run:  python extractor.py

Configuration
-------------
Edit the settings below (INPUT_FILE, OUTPUT_FILE, delays, retries).
The input file must have a column for Admission Number and a column for
First Name. Extra columns are carried over to the output automatically.
"""

import argparse
import csv
import hashlib
import hmac
import json
import os
import random
import re
import signal
import socket
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

import pandas as pd

from seleniumbase import SB

# ---------------------------------------------------------------------------
# CONFIGURATION - edit these to match your school's setup
# ---------------------------------------------------------------------------

# Every runtime file lives next to the app, not in whatever directory the
# process happened to be started from. The launcher uses the same BASE_DIR, so
# pause.flag / stop.flag / progress.csv are always found by both sides - even
# when the frozen EXE is launched from a desktop shortcut.
_IS_FROZEN = getattr(sys, "frozen", False)
BASE_DIR = os.path.dirname(os.path.abspath(sys.executable if _IS_FROZEN else __file__))

INPUT_FILE = os.path.join(BASE_DIR, "students.xlsx")   # students list (admission no + first name)
OUTPUT_FILE = os.path.join(BASE_DIR, "results.xlsx")   # where the results go
PROGRESS_FILE = os.path.join(BASE_DIR, "progress.csv")  # live progress/backup file
DEBUG_FILE = os.path.join(BASE_DIR, "debug.log")        # request/response history
RAW_SAMPLE_FILE = os.path.join(BASE_DIR, "raw_sample.json")  # first API response

SITE_URL = "https://result.eaes.et/"
API_BASE = "https://api.eaes.et/api/v1"
RESULTS_ENDPOINT = API_BASE + "/results/web"

ADMISSION_DIGITS = 8                   # ESSLCE admission numbers are 8 digits (with leading zeros)

MIN_DELAY = 4                          # seconds, minimum wait between students
MAX_DELAY = 8                          # seconds, maximum wait between students
PAUSE_EVERY = 25                       # take a longer pause every N students
PAUSE_MIN = 20                         # longer pause min seconds
PAUSE_MAX = 40                         # longer pause max seconds

# Pause / Stop cooperation with the GUI launcher. The launcher creates
# pause.flag to ask for a pause, deletes it to resume, and creates stop.flag
# to ask for a clean stop (progress is always flushed before exiting).
PAUSE_FILE = os.path.join(BASE_DIR, "pause.flag")
STOP_FILE = os.path.join(BASE_DIR, "stop.flag")
NETWORK_FAIL_LIMIT = 5                 # consecutive network errors before auto-pausing
NETWORK_RETRY_MIN = 5                  # seconds between connectivity checks while auto-paused

_INTERRUPTED = False


def _on_sigint(sig, frame):
    """Ctrl+C handler: finish the current student, then save and exit cleanly."""
    global _INTERRUPTED
    _INTERRUPTED = True
    print("\nInterrupt received - finishing current work, then saving progress...")

TRY_STUDENTS = 0                       # 0 = process ALL students (full run for the whole batch)

TOKEN_WAIT = 12                        # max seconds to wait for a Turnstile token
LOOKUP_RETRIES = 3                     # retries per student on transient errors
BACKOFF_SECONDS = 20                   # wait after a security/duplicate error

TIMEOUT = 30                           # HTTP timeout for API calls

SAVE_RAW_SAMPLE = True                # dump the first API response to raw_sample.json

_HARVESTED_TOKEN = None                # set before each API lookup

# Which columns to look for in the input file (case-insensitive, partial match)
ADMISSION_KEYS = [
    "admission", "admission_no", "admission number", "candidate", "reg",
    "registration", "id number",
]
FIRSTNAME_KEYS = ["first_name", "first name", "firstname", "fname", "name"]


# ---------------------------------------------------------------------------
# Column detection
# ---------------------------------------------------------------------------

def _match_key(columns, keys):
    low = {str(c).strip().lower(): c for c in columns}
    for key in keys:
        for col, original in low.items():
            if col == key or col.startswith(key) or key in col:
                return original
    return None


def read_students_file(path):
    """Read an Excel or CSV/plain-text students file, handling common
    encodings and tab-separated layouts that Excel/phone exports produce."""
    path = str(path)
    if path.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(path)

    try:
        with open(path, "rb") as fh:
            head = fh.read(2)
    except OSError:
        head = b""

    encodings = ("utf-8-sig", "utf-8", "cp1252", "latin-1")
    if head in (b"\xff\xfe", b"\xfe\xff"):
        encodings = ("utf-16",) + encodings

    for encoding in encodings:
        try:
            df = pd.read_csv(path, encoding=encoding, sep=",")
        except (UnicodeDecodeError, pd.errors.ParserError,
                pd.errors.EmptyDataError):
            continue
        header = str(df.columns[0])
        if len(df.columns) == 1 and "\t" in header:
            try:
                return pd.read_csv(path, encoding=encoding, sep="\t")
            except (UnicodeDecodeError, pd.errors.ParserError,
                    pd.errors.EmptyDataError):
                pass
        return df
    return pd.read_csv(path, encoding="latin-1", sep="\t")


def find_columns(df):
    ad = _match_key(df.columns, ADMISSION_KEYS)
    fn = _match_key(df.columns, FIRSTNAME_KEYS)
    if ad is None or fn is None:
        print("Could not find admission-number and first-name columns in the file.")
        print("Found columns:", list(df.columns))
        print("Expected something like 'Admission Number' and 'First Name'.")
        sys.exit(1)
    extra = [c for c in df.columns if c not in (ad, fn)]
    return ad, fn, extra


def normalize_admission(value):
    """Pad the admission number with leading zeros to the ESSLCE width.

    Example: 1234567 -> 01234567
    """
    raw = str(value).strip()
    digits = re.sub(r"\D", "", raw)
    if digits.isdigit() and len(digits) < ADMISSION_DIGITS:
        return digits.zfill(ADMISSION_DIGITS)
    return raw


# ---------------------------------------------------------------------------
# Turnstile token harvesting (inside the browser)
# ---------------------------------------------------------------------------

def _token_js():
    return (
        "(() => { const el = document.querySelector("
        "'textarea[name=cf-turnstile-response], input[name=cf-turnstile-response]');"
        " return el ? (el.value || '') : ''; })()"
    )


def harvest_token(sb, site_url, max_attempts=4):
    """Get a fresh Turnstile token from the browser. Returns token or None."""
    for attempt in range(max_attempts):
        try:
            sb.open(site_url)
            sb.sleep(2.5 + random.random())
            try:
                sb.solve_captcha()
            except Exception:
                pass
            for _ in range(int(TOKEN_WAIT / 2) + 1):
                tok = sb.execute_script(_token_js()) or ""
                if len(tok) > 40:
                    return tok.strip()
                sb.sleep(2)
            # Maybe an interactive checkbox needs a click
            try:
                sb.uc_gui_click_captcha()
                sb.sleep(3)
                tok = sb.execute_script(_token_js()) or ""
                if len(tok) > 40:
                    return tok.strip()
            except Exception:
                pass
        except Exception as exc:
            print(f"    token attempt {attempt + 1} failed: {exc!r}")
    return None


# ---------------------------------------------------------------------------
# Signed results API call
# ---------------------------------------------------------------------------

def _sign(secret, message):
    return hmac.new(secret.encode(), message.encode(), hashlib.sha256).hexdigest()


def _api_get(url):
    """Fetch a GET endpoint, retrying transient network/TLS errors so a flaky
    connection cannot abort a whole batch. Raises LookupError when exhausted."""
    last_exc = None
    for attempt in range(1, LOOKUP_RETRIES + 1):
        try:
            with urllib.request.urlopen(url, timeout=TIMEOUT) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            last_exc = exc
            if attempt < LOOKUP_RETRIES:
                print(f"    session request attempt {attempt} failed ({exc!r}) "
                      f"- retrying in {BACKOFF_SECONDS}s")
                time.sleep(BACKOFF_SECONDS)
    raise LookupError(f"session: {last_exc}")


def _log_debug(message):
    with open(DEBUG_FILE, "a", encoding="utf-8") as fh:
        fh.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} | {message}\n")


def _stop_requested():
    return os.path.exists(STOP_FILE) or _INTERRUPTED


def _wait_while_paused():
    """Block while pause.flag exists. Returns True if a stop was requested."""
    while os.path.exists(PAUSE_FILE):
        if _stop_requested():
            return True
        print("    PAUSED - waiting for resume (Stop to exit)...", flush=True)
        time.sleep(1)
    return False


def _wait_for_network():
    """Wait until the API host accepts a TCP connection (internet back) or a
    stop is requested. Returns False when the caller should stop."""
    host = urllib.parse.urlparse(API_BASE).hostname
    while True:
        if _stop_requested():
            return False
        try:
            sock = socket.create_connection((host, 443), timeout=5)
            sock.close()
            print("    connectivity restored - resuming.", flush=True)
            return True
        except OSError:
            print("    internet still unreachable - waiting...", flush=True)
            time.sleep(NETWORK_RETRY_MIN)


def lookup_student(admission, first_name):
    """Returns parsed result dict or raises a LookupError with a message."""
    session = _api_get(f"{API_BASE}/session/key")
    sid, sec = session["session_id"], session["session_secret"]

    ts = str(int(time.time() * 1000))
    signature_input = f"{ts}:{admission.strip().lower()}:{first_name.strip().lower()}"
    sig = _sign(sec, signature_input)

    headers = {
        "Content-Type": "application/json",
        "X-Session-Id": sid,
        "X-Timestamp": ts,
        "X-Signature": sig,
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36"
        ),
        "Origin": "https://result.eaes.et",
        "Referer": "https://result.eaes.et/",
    }
    body = json.dumps({
        "admission_no": admission.strip(),
        "first_name": first_name.strip().upper(),
        "turnstileToken": _HARVESTED_TOKEN,
    }).encode("utf-8")
    _log_debug(
        f"REQUEST admission={admission.strip().lower()!r} "
        f"name={first_name.strip().lower()!r} "
        f"token_len={len(_HARVESTED_TOKEN or '')}"
    )

    req = urllib.request.Request(RESULTS_ENDPOINT, data=body, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
            raw = resp.read().decode("utf-8")
        _log_debug(
            f"LOOKUP ok   admission={admission!r} name={first_name!r} "
            f"-> {raw[:600]}"
        )
        return json.loads(raw)
    except urllib.error.HTTPError as err:
        detail = ""
        try:
            detail = err.read().decode("utf-8")
        except Exception:
            pass
        _log_debug(
            f"LOOKUP http{err.code} admission={admission!r} name={first_name!r} "
            f"-> {detail[:600]!r}"
        )
        if err.code == 404:
            raise LookupError(f"not_found: {detail[:300]}")
        if err.code in (403, 429):
            raise LookupError(f"challenge_{err.code}: {detail}")
        raise LookupError(f"http_{err.code}: {detail}")
    except (urllib.error.URLError, TimeoutError, OSError):
        _log_debug(f"LOOKUP network admission={admission!r} name={first_name!r}")
        raise LookupError("network")


# ---------------------------------------------------------------------------
# Result parsing
# ---------------------------------------------------------------------------

def _find_key(d, *candidates):
    """Case-insensitive key lookup with partial matching."""
    if not isinstance(d, dict):
        return None
    low = {str(k).strip().lower(): v for k, v in d.items()}
    flat = {k.replace(" ", "").replace("_", ""): v for k, v in low.items()}
    for cand in candidates:
        c = cand.lower().replace(" ", "").replace("_", "")
        for k, v in list(low.items()) + list(flat.items()):
            if k == c or k in c or c in k:
                return v
    return None


def _extract_result_value(item):
    """A result may be a plain value or nested inside another dict."""
    if not isinstance(item, dict):
        return item
    return _find_key(item, "result", "score", "mark", "grade", "value", "point", "val")


def split_name(full_name):
    """Split a full name into first, middle, and last name parts.

    Handles 1, 2, 3 or more tokens positionally:
      "HANA TESFAYE GEBRE"    -> ("HANA", "TESFAYE", "GEBRE")
      "HANA GEBRE"            -> ("HANA", "", "GEBRE")
      "HANA"                  -> ("HANA", "", "")
      "A B C D"               -> ("A", "B C", "D")
    """
    parts = [p for p in str(full_name or "").split() if p]
    if not parts:
        return "", "", ""
    first = parts[0]
    if len(parts) == 1:
        return first, "", ""
    middle = " ".join(parts[1:-1])
    return first, middle, parts[-1]


def parse_result(data):
    """Turn the API JSON response into a flat dict of columns.

    Extracts the student info (including Total / Average) and ONE column per
    subject, so every subject the API returns ends up in the Excel output.
    """
    payload = data or {}
    info = payload.get("studentInfo") or payload.get("StudentInfo") \
        or payload.get("student") or {}
    if not isinstance(info, dict):
        info = {}

    full_name = _find_key(info, "full_name", "fullname", "name")
    first_name, middle_name, last_name = split_name(full_name)

    row = {
        "First Name": first_name,
        "Middle Name": middle_name,
        "Last Name": last_name,
        "School": _find_key(info, "school"),
        "Admission No": _find_key(info, "admission_no", "admission", "candidate"),
        "Stream": _find_key(info, "stream", "track"),
        "Sex": _find_key(info, "sex", "gender"),
        "Age": _find_key(info, "age"),
        "Total": _find_key(info, "total", "total_score", "total_score_result"),
        "Average": _find_key(info, "avg", "average"),
    }

    subjects = payload.get("results")
    if subjects is None:
        subjects = payload.get("Results") or payload.get("subjectResults") \
            or payload.get("subjects") or payload.get("detail")
    if isinstance(subjects, dict):
        subjects = [{"Subject": k, "Result": v} for k, v in subjects.items()]

    for subj in subjects or []:
        if not isinstance(subj, dict):
            continue
        subject = str(_find_key(subj, "subject", "subject_name", "course", "name") or "Unknown").strip()
        result_val = _extract_result_value(_find_key(subj, "result", "score", "mark", "grade", "point"))
        if subject.lower() in ("total", "total score", "average", "avg"):
            if subject.lower().startswith("average") and row["Average"] is None:
                row["Average"] = result_val
            elif row["Total"] is None:
                row["Total"] = result_val
            continue
        while subject and subject in row and row[subject] is not None:
            subject = f"{subject} (2)"
        if subject:
            row[subject] = result_val

    for key, val in list(row.items()):
        if val is None:
            continue
        if isinstance(val, (str, bool)):
            try:
                num = float(val)
                if str(val).strip().replace(".", "", 1).isdigit():
                    row[key] = int(num) if num.is_integer() else num
            except (TypeError, ValueError):
                pass
    return row


# ---------------------------------------------------------------------------
# Progress / output
# ---------------------------------------------------------------------------

def _row_key(row):
    """Stable identity for a student, used to skip already-collected students.

    The admission number is normalised on both sides, because the input file and
    the API do not agree on the width: the file may hold 65001 while the portal
    reports it zero-padded, or the other way round. Without normalising here, a
    resumed run would never match and would re-fetch the whole batch.
    """
    admission = normalize_admission(row.get("Admission No") or "")
    name = str(row.get("First Name") or "").strip().upper()
    return f"{admission}|{name}".strip()


def _migrate_row(row):
    """Split a legacy 'Full Name' (combined) value into separate First /
    Middle / Last Name columns, so old progress files adopt the current
    output format. The API full name is authoritative.
    """
    full = row.get("Full Name")
    if full:
        first, middle, last = split_name(full)
        if first:
            row["First Name"] = first
        row["Middle Name"] = middle
        row["Last Name"] = last
        row.pop("Full Name", None)
    return row


def load_done():
    done = set()
    if not os.path.exists(PROGRESS_FILE):
        return done
    with open(PROGRESS_FILE, "r", encoding="utf-8", newline="") as fh:
        for row in csv.DictReader(fh):
            if row.get("Status") not in ("Found", "Not Found"):
                continue
            done.add(_row_key(row))
    return done


def load_progress_rows():
    rows = {}
    if not os.path.exists(PROGRESS_FILE):
        return rows
    with open(PROGRESS_FILE, "r", encoding="utf-8", newline="") as fh:
        for row in csv.DictReader(fh):
            key = _row_key(row)
            if key:
                rows[key] = _migrate_row(row)
    return rows


def _csv_value(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return value


def save_progress(rows):
    merged = load_progress_rows()
    for r in rows:
        key = _row_key(r)
        if not key:
            continue
        merged[key] = {k: _csv_value(r.get(k)) for k in r}
    if not merged:
        return
    cols = []
    for r in merged.values():
        for k in r:
            if k not in cols:
                cols.append(k)
    with open(PROGRESS_FILE, "w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=cols)
        writer.writeheader()
        for r in merged.values():
            writer.writerow({k: r.get(k, "") for k in cols})


# ---------------------------------------------------------------------------
# Excel output with one sheet per study stream
# ---------------------------------------------------------------------------

SHEET_NATURAL = "Natural Science"
SHEET_SOCIAL = "Social Science"
SHEET_OTHER = "Other"
SHEET_ORDER = (SHEET_NATURAL, SHEET_SOCIAL, SHEET_OTHER)


def _sheet_name(stream):
    """Map a Stream value to an output sheet name."""
    text = str(stream or "").lower()
    if "natural" in text:
        return SHEET_NATURAL
    if "social" in text or "human" in text:
        return SHEET_SOCIAL
    return SHEET_OTHER


def write_workbook(df, path):
    """Write the results DataFrame to a single workbook with one sheet per
    stream (Natural Science / Social Science / Other). Rows without a
    recognisable stream (e.g. Not Found / Error) land in 'Other'."""
    df = df.copy()
    if "Stream" in df.columns:
        df["_Sheet"] = df["Stream"].fillna("").astype(str).apply(_sheet_name)
    else:
        df["_Sheet"] = SHEET_OTHER
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name in SHEET_ORDER:
            group = df[df["_Sheet"] == name].drop(columns=["_Sheet"])
            group.to_excel(writer, sheet_name=name, index=False)


# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------

def main():
    global INPUT_FILE, OUTPUT_FILE
    parser = argparse.ArgumentParser(description="ESSLCE Result Extractor")
    parser.add_argument("-i", "--input", default=INPUT_FILE,
                        help="path to the students Excel/CSV file")
    parser.add_argument("-o", "--output", default=OUTPUT_FILE,
                        help="path where the results Excel file is saved")
    args = parser.parse_args()

    INPUT_FILE = args.input or INPUT_FILE
    OUTPUT_FILE = args.output or OUTPUT_FILE

    df = read_students_file(INPUT_FILE)
    ad_col, fn_col, extra_cols = find_columns(df)

    records = []
    done_keys = load_done()
    total = len(df)

    print("=" * 60)
    print("ESSLCE Result Extractor")
    print(f"Students: {total}   (collected: {len(done_keys)})")
    if TRY_STUDENTS:
        print(f"TEST MODE: only trying the first {TRY_STUDENTS} students (- set TRY_STUDENTS = 0 to do all)")
    print(f"Working with columns: {ad_col!r} + {fn_col!r}")
    if extra_cols:
        print(f"Extra columns carried to output: {extra_cols}")
    print("=" * 60)

    start_time = time.time()
    processed = 0
    skipped = 0
    network_fails = 0
    signal.signal(signal.SIGINT, _on_sigint)

    with SB(uc=True, test=True, incognito=True) as sb:
        sb.activate_cdp_mode(SITE_URL)
        sb.sleep(3)

        for idx, (_, student) in enumerate(df.iterrows()):
            admission = normalize_admission(student[ad_col])
            first_name = str(student[fn_col]).strip().upper()
            if not admission or not first_name:
                continue
            key = _row_key({"Admission No": admission, "First Name": first_name})
            if key in done_keys:
                skipped += 1
                continue

            if _stop_requested():
                print("    stop requested - finishing cleanly...", flush=True)
                break
            if os.path.exists(PAUSE_FILE):
                save_progress(records)
                if _wait_while_paused():
                    break
                print("    resumed.", flush=True)

            carry = {c: student[c] for c in extra_cols if c in student}
            row = {"Admission No": admission, "First Name": first_name, **carry}

            print(f"[{idx + 1}/{total}] {admission} | {first_name}", flush=True)
            print("    obtaining security token...", flush=True)

            token = harvest_token(sb, SITE_URL)
            while token is None and not _stop_requested():
                print("    token failed - waiting then retrying...")
                time.sleep(BACKOFF_SECONDS)
                token = harvest_token(sb, SITE_URL)
            if token is None:
                print("    stop requested while obtaining token - exiting.")
                break

            global _HARVESTED_TOKEN
            _HARVESTED_TOKEN = token

            result_row = None
            outcome = None
            retries = 0
            stop_now = False
            while retries < LOOKUP_RETRIES:
                try:
                    data = lookup_student(admission, first_name)
                except LookupError as exc:
                    msg = str(exc)
                    if msg.startswith("not_found"):
                        outcome = "Not Found"
                        row["Status"] = outcome
                        records.append(row)
                        network_fails = 0
                        print(f"    NOT FOUND - {msg}")
                        break
                    if "challenge" in msg:
                        print(f"    {msg} -> fresh token + backoff {BACKOFF_SECONDS}s")
                        token = harvest_token(sb, SITE_URL)
                        if token:
                            _HARVESTED_TOKEN = token
                        time.sleep(BACKOFF_SECONDS)
                        retries += 1
                        continue
                    if msg == "network":
                        network_fails += 1
                        if network_fails >= NETWORK_FAIL_LIMIT:
                            save_progress(records)
                            print(f"    {NETWORK_FAIL_LIMIT} consecutive network errors -> "
                                  "auto-pausing until internet returns...")
                            if not _wait_for_network():
                                stop_now = True
                                break
                            network_fails = 0
                        print(f"    network error -> retry in {BACKOFF_SECONDS}s")
                        time.sleep(BACKOFF_SECONDS)
                        retries += 1
                        continue
                    print(f"    error {msg} -> retry in {BACKOFF_SECONDS}s")
                    time.sleep(BACKOFF_SECONDS)
                    retries += 1
                    continue
                except Exception as exc:
                    _log_debug(
                        f"UNEXPECTED admission={admission!r} name={first_name!r} "
                        f"-> {exc!r}"
                    )
                    print(f"    unexpected error {exc!r} -> retry in {BACKOFF_SECONDS}s")
                    time.sleep(BACKOFF_SECONDS)
                    retries += 1
                    continue
                # success
                try:
                    if SAVE_RAW_SAMPLE and not os.path.exists(RAW_SAMPLE_FILE):
                        with open(RAW_SAMPLE_FILE, "w", encoding="utf-8") as fh:
                            json.dump(data, fh, ensure_ascii=False, indent=2)
                        print("    (raw response saved to raw_sample.json - check that all subjects are there)")
                    result_row = parse_result(data)
                except Exception as exc:
                    _log_debug(
                        f"PARSE FAIL admission={admission!r} name={first_name!r} "
                        f"-> {exc!r}"
                    )
                    print(f"    parse error {exc!r} -> retry in {BACKOFF_SECONDS}s")
                    time.sleep(BACKOFF_SECONDS)
                    retries += 1
                    continue
                outcome = "Found"
                network_fails = 0
                row.update(result_row)
                row["Status"] = outcome
                records.append(row)
                full = " ".join(p for p in (
                    row.get("First Name"), row.get("Middle Name"),
                    row.get("Last Name")) if p)
                tot = row.get("Total")
                print(f"    FOUND {full}  total={tot}")
                break
            else:
                row["Status"] = "Error"
                records.append(row)
                print("    ERROR after retries")

            if stop_now:
                print("    stopping - progress saved.")
                break

            processed += 1
            if processed % 5 == 0:
                save_progress(records)
            if TRY_STUDENTS and processed >= TRY_STUDENTS:
                print(f"    test limit reached ({TRY_STUDENTS}) - stopping.")
                break

            # throttling between students
            delay = random.uniform(MIN_DELAY, MAX_DELAY)
            if processed % PAUSE_EVERY == 0:
                extra = random.uniform(PAUSE_MIN, PAUSE_MAX)
                print(f"    pause {int(extra)}s to be kind to the server...")
                time.sleep(extra)
            else:
                time.sleep(delay)

    save_progress(records)

    merged_rows = list(load_progress_rows().values())
    out = pd.DataFrame(merged_rows) if merged_rows else pd.DataFrame(records)
    try:
        write_workbook(out, OUTPUT_FILE)
        final_file = OUTPUT_FILE
    except PermissionError:
        base, ext = os.path.splitext(OUTPUT_FILE)
        final_file = f"{base}_{time.strftime('%Y%m%d_%H%M%S')}{ext}"
        print()
        print("!" * 60)
        print("Could not overwrite the results file - it is probably open in")
        print("Excel / WPS Office. Please close it, or use the file saved below.")
        print("!" * 60)
        write_workbook(out, final_file)
    print("=" * 60)
    print(f"Done. {processed} checked, {skipped} already done (resume).")
    print(f"Saved: {final_file}  (cumulative across runs)")
    elapsed = time.time() - start_time
    print(f"Time used: {int(elapsed // 60)}m {int(elapsed % 60)}s")


if __name__ == "__main__":
    main()