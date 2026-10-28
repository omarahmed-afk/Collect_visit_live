# =====================================================================
# THIS SCRIPT'S CLINIC - the only thing that differs between the 32 files
# =====================================================================
CLINIC_NAME = "Greenpoint"
CLINIC_HAS_OT = False      # True = this clinic also has an OT column
CLINIC_SHEET_TAB = "Greenpoint"   # this clinic's tab in the spreadsheet (PT and OT together)

import os
import sys
import json
import signal
from pathlib import Path
import re
import time
import random
import logging
import pandas as pd

from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# =====================================================================
# SHEET WRITE GUARD - stops the 32 scripts colliding on the Google Sheet
# (lock + retries + count-only upsert + random 3-5 min gap after each run)
# =====================================================================
import os
import time
import uuid
import random
import logging
from contextlib import contextmanager


log = logging.getLogger(__name__)


# ---------------------------------------------------------------------
# SETTINGS
# ---------------------------------------------------------------------
# Every script that writes to the same spreadsheet MUST use the same path.
# To share the lock with the 32 existing scripts, set this to the same
# file they use (see the note in the README).
SHEET_LOCK_PATH = os.environ.get("SHEET_LOCK_PATH") or os.path.join(
    os.path.dirname(os.path.abspath(__file__)), ".locks", "google_sheet.lock")
USE_SHEET_LOCK = os.environ.get("USE_SHEET_LOCK", "1") != "0"
SHEET_LOCK_TIMEOUT = int(os.environ.get("SHEET_LOCK_TIMEOUT", "150"))
SHEET_LOCK_STALE = int(os.environ.get("SHEET_LOCK_STALE", "240"))
SCRIPT_LABEL = os.environ.get("SCRIPT_LABEL", "script")   # shown inside the lock file


API_ATTEMPTS = 5
_TRANSIENT = ("429", "quota", "rate limit", "ratelimitexceeded",
              "500", "502", "503", "backenderror", "internal error",
              "timed out", "timeout", "connection")


_my_token = {"value": None}     # set while THIS process holds the lock




# =====================================================================
# 1. CROSS-PROCESS LOCK
# =====================================================================


def acquire_sheet_lock(timeout=None):
    """Takes the shared lock. Always returns True eventually - after the
    timeout it gives up waiting and lets the write go ahead (the retries in
    _api absorb any quota error). Returns True whether or not it owns it;
    ownership is tracked internally so release() is safe."""
    _my_token["value"] = None
    if not USE_SHEET_LOCK:
        return True


    os.makedirs(os.path.dirname(SHEET_LOCK_PATH) or ".", exist_ok=True)
    timeout = SHEET_LOCK_TIMEOUT if timeout is None else timeout
    started = time.time()
    deadline = started + timeout
    token = f"{SCRIPT_LABEL} {os.getpid()} {uuid.uuid4().hex} {time.time()}"


    while True:
        try:
            fd = os.open(SHEET_LOCK_PATH, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, token.encode())
            os.close(fd)
            _my_token["value"] = token
            waited = time.time() - started
            if waited > 1:
                log.info(f"  -  Waited {waited:.1f}s for the sheet lock")
            return True
        except FileExistsError:
            pass
        except Exception as e:
            log.debug(f"lock error: {e}")
            return True                 # never block the run on the lock itself


        # a lock left behind by a crashed copy
        try:
            age = time.time() - os.path.getmtime(SHEET_LOCK_PATH)
            if age > SHEET_LOCK_STALE:
                log.warning(f"Removing a stale sheet lock ({age:.0f}s old)")
                os.remove(SHEET_LOCK_PATH)
                continue
        except FileNotFoundError:
            continue                    # released meanwhile - try again now
        except Exception:
            pass


        if time.time() > deadline:
            log.warning(f"Another copy has held the sheet lock for {timeout}s - "
                        f"writing anyway rather than delaying these rows")
            return True


        time.sleep(random.uniform(0.5, 2.0))




def release_sheet_lock():
    """Removes the lock only if this process is the one holding it."""
    token = _my_token["value"]
    _my_token["value"] = None
    if not token:
        return
    try:
        with open(SHEET_LOCK_PATH, "r") as f:
            if f.read() != token:
                return                  # someone else's lock - leave it
        os.remove(SHEET_LOCK_PATH)
    except Exception:
        pass




@contextmanager
def sheet_lock(timeout=None):
    """with sheet_lock():  ...read + write the sheet here...  """
    acquire_sheet_lock(timeout)
    try:
        yield
    finally:
        release_sheet_lock()




# =====================================================================
# 2. RETRY ON QUOTA / TRANSIENT ERRORS
# =====================================================================


def _api(fn, *args, **kwargs):
    """Runs one Google API call, retrying when the quota is hit.
    All copies share the same ~60 writes/minute allowance, so a burst
    returns 429. Retrying after a growing pause turns a lost write into a
    slow one."""
    attempts = kwargs.pop("_attempts", API_ATTEMPTS)
    for attempt in range(1, attempts + 1):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            msg = str(e).lower()
            transient = any(k in msg for k in _TRANSIENT)
            if not transient or attempt == attempts:
                raise
            pause = min(60, (2 ** attempt) + random.uniform(0, 1.5))
            log.warning(f"Google API busy ({type(e).__name__}) - retrying in "
                        f"{pause:.1f}s (attempt {attempt}/{attempts})")
            time.sleep(pause)
    return None




# =====================================================================
# 3. SAFE UPSERT (read fresh inside the lock, update in place, append new)
# =====================================================================


def _col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s




def _same_value(a, b):
    """1 == "1" == "1.0" == " 1 "; non-numbers compared as trimmed text."""
    sa, sb = str(a if a is not None else "").strip(), str(b if b is not None else "").strip()
    try:
        return float(sa.replace(",", "")) == float(sb.replace(",", ""))
    except ValueError:
        return sa == sb




def upsert_rows(ws, rows, key_fn, headers=None, lock=True, update_columns=None):
    """Writes `rows` (lists) into worksheet `ws`, one row per key.


    key_fn(row) -> hashable key or None. Rows whose key already exists in
    the sheet are updated in place; the rest are appended.


    update_columns: list of 0-based column indexes. When given, an existing
    row only gets THOSE cells rewritten (e.g. [4] = column E "Count"), and
    only if the value actually changed. When None, the whole row is
    overwritten.


    Returns the set of keys that were actually written. A failure of the
    append does not roll back the updates (same as the originals)."""
    if not rows:
        return set()


    def _do():
        existing = _api(ws.get_all_values) or []


        if not existing and headers:
            _api(ws.append_row, list(headers), value_input_option="RAW")
            existing = [list(headers)]


        index, current = {}, {}
        for line_no, row in enumerate(existing[1:], start=2):
            k = key_fn(row)
            if k:
                index[k] = line_no
                current[k] = row


        width = max(len(headers or []), max(len(r) for r in rows))
        last_col = _col_letter(width)
        next_free = len(existing) + 1


        updates, appends = [], []
        update_keys, append_keys = set(), set()
        for row in rows:
            k = key_fn(row)
            if k in index:
                n = index[k]
                if update_columns is None:
                    updates.append({"range": f"A{n}:{last_col}{n}", "values": [row]})
                    update_keys.add(k)
                else:
                    old = current.get(k, [])
                    for c in update_columns:
                        new_v = row[c] if c < len(row) else ""
                        old_v = old[c] if c < len(old) else ""
                        if not _same_value(new_v, old_v):
                            cell = f"{_col_letter(c + 1)}{n}"
                            updates.append({"range": cell, "values": [[new_v]]})
                            update_keys.add(k)
                    current[k] = row
            else:
                appends.append(row)
                append_keys.add(k)
                if k:
                    index[k] = next_free    # a repeat in this batch -> update
                next_free += 1


        done = set()
        if updates:
            try:
                try:
                    _api(ws.batch_update, updates, value_input_option="RAW")
                except TypeError:
                    _api(ws.batch_update, updates)      # older gspread
                done |= update_keys
                log.info(f"Updated {len(updates)} existing row(s) in place")
            except Exception as e:
                log.error(f"Updating existing rows failed: {e}")


        if appends:
            try:
                _api(ws.append_rows, appends, value_input_option="RAW")
                done |= append_keys
                log.info(f"Added {len(appends)} new row(s)")
            except Exception as e:
                log.error(f"Adding new rows failed: {e}")
        return done


    if not lock:
        return _do()
    with sheet_lock():
        return _do()




# =====================================================================
# 3b. VISIT ROWS - one row per Clinic + date_of_visit + Time + Appointment
#     Type. When an appointment changes, the SAME row is rewritten.
# =====================================================================

_INCOMPLETE_SCOPES = set()   # (clinic, date) whose last read did not add up


def _visit_key(r):
    if len(r) < 4 or str(r[0]).strip() in ("", "Clinic"):
        return None
    return tuple(normalize_key_part(x) for x in r[:4])


def sync_visit_rows(ws, rows, headers):
    """Makes the sheet match `rows` for the clinic + dates contained in `rows`.
    Call it while holding sheet_lock().

      same key, same Count       -> nothing is written
      same key, other Count      -> only column E of that row
      a row whose key is no longer on the page + a new key with the same
      clinic + date + TIME (type changed, e.g. Initial Examination ->
      Follow-Up)                 -> that SAME row is rewritten, no new row
      a time with nobody left    -> its row is deleted (a new time gets a
                                    new row)
      new key, no row to reuse   -> appended
      row no longer on the page and nothing replaces it -> row deleted,
                                    never left at 0
    Old duplicate rows of one key are reused / deleted the same way.
    A clinic + date whose last read did not add up is never reused or
    deleted from (new keys are appended, old rows are left as they are),
    so a bad read cannot wipe good rows.
    Returns False (nothing written) when this process does not own the lock:
    rows are only rewritten or deleted by row number under the lock."""
    if USE_SHEET_LOCK and not _my_token["value"]:
        log.warning("Sheet lock not owned (timeout) - skipping this write; "
                    "it is sent again at the next publish")
        return False

    existing = _api(ws.get_all_values) or []
    if not existing:
        _api(ws.append_row, list(headers), value_input_option="RAW")
        existing = [list(headers)]

    incoming = {}
    for r in rows:
        k = _visit_key(r)
        if k:
            incoming[k] = list(r[:5])
    scopes = {(k[0], k[1]) for k in incoming}

    first_row, spare = {}, {}
    for n, r in enumerate(existing[1:], start=2):
        k = _visit_key(r)
        if not k or (k[0], k[1]) not in scopes:
            continue                       # another clinic / another day
        if k in incoming and k not in first_row:
            first_row[k] = n
        else:
            spare.setdefault((k[0], k[1]), []).append(n)   # gone, or a duplicate

    updates, appends, deletes = [], [], []
    for k, r in incoming.items():
        if k in first_row:
            n = first_row[k]
            old = existing[n - 1]
            if not _same_value(r[4], old[4] if len(old) > 4 else ""):
                updates.append({"range": f"E{n}", "values": [[r[4]]]})
            continue
        scope = (k[0], k[1])
        free = [] if scope in _INCOMPLETE_SCOPES else spare.get(scope, [])
        same_time = [n for n in free if _visit_key(existing[n - 1])[2] == k[2]]
        if same_time:
            n = same_time[0]          # same clinic + date + TIME, type changed
            free.remove(n)
            log.info(f"  -  Row {n}: {existing[n - 1][2:5]} -> {r[2:5]} (same row)")
            updates.append({"range": f"A{n}:E{n}", "values": [r]})
        else:
            appends.append(r)

    for scope, free in spare.items():
        if scope in _INCOMPLETE_SCOPES:
            if free:
                log.info(f"  -  {scope}: last read did not add up - "
                         f"{len(free)} old row(s) left as they are")
            continue
        deletes.extend(free)

    ok_so_far = True
    if updates:
        try:
            try:
                _api(ws.batch_update, updates, value_input_option="RAW")
            except TypeError:
                _api(ws.batch_update, updates)
            log.info(f"Updated {len(updates)} cell range(s) in place")
        except Exception as e:
            ok_so_far = False
            log.error(f"Updating rows failed: {e}")

    # the layout as it is now, to find where new rows belong
    current = [list(r) for r in existing]
    for u in updates:
        m = re.match(r"A(\d+):E\1$", u["range"])
        if m:
            current[int(m.group(1)) - 1] = list(u["values"][0])

    if deletes and ok_so_far:
        order = sorted(set(deletes), reverse=True)     # bottom first
        for n in order:
            log.info(f"  -  Row {n}: {existing[n - 1][:5]} is no longer on the page - deleted")
        try:
            requests = [{"deleteDimension": {"range": {
                "sheetId": ws.id, "dimension": "ROWS",
                "startIndex": n - 1, "endIndex": n}}} for n in order]
            try:
                _api(ws.spreadsheet.batch_update, {"requests": requests})
            except AttributeError:
                for n in order:
                    _api(ws.delete_rows, n)
            for n in order:
                del current[n - 1]
        except Exception as e:
            log.error(f"Deleting rows failed: {e}")

    if appends and ok_so_far:
        _insert_in_order(ws, current, appends)
    elif appends:
        log.warning(f"{len(appends)} new row(s) not added because an earlier step failed "
                    f"- they are sent again at the next publish")
    if ok_so_far:
        try:
            _sort_tab_if_needed(ws, current)
        except Exception as e:
            log.error(f"Sorting the tab failed: {e}")
    return ok_so_far


def _sheet_value(v):
    """'3' -> 3 so a rewritten Count stays a number."""
    if isinstance(v, str) and re.fullmatch(r"-?\d+", v.strip()):
        return int(v.strip())
    return v


def _sort_tab_if_needed(ws, current):
    """`current` = the tab as it is now (header first). If its rows are not
    in date / time / clinic / type order, rewrites them in that order.
    Rows that cannot be read as a visit stay at the bottom."""
    data = current[1:]
    keyed = [(_visit_sort_key(r), i, r) for i, r in enumerate(data)]
    good = [x for x in keyed if x[0] is not None]
    bad = [x for x in keyed if x[0] is None]
    in_order = (all(good[i][0] <= good[i + 1][0] for i in range(len(good) - 1))
                and all(x[1] > good[-1][1] for x in bad) if good else True)
    if in_order or not data:
        return
    wanted = [r for _, _, r in sorted(good, key=lambda x: (x[0], x[1]))] + [r for _, _, r in bad]
    values = [[_sheet_value(v) for v in (list(r[:5]) + [""] * (5 - len(r[:5])))] for r in wanted]
    body = [{"range": f"A2:E{len(values) + 1}", "values": values}]
    try:
        _api(ws.batch_update, body, value_input_option="RAW")
    except TypeError:
        _api(ws.batch_update, body)
    current[1:] = wanted
    log.info(f"Tab re-sorted by date and time ({len(values)} row(s))")


def _visit_sort_key(r):
    """(date, time, clinic, type) - 'No appointments' (no time) sorts first."""
    k = _visit_key(r)
    if not k:
        return None
    clinic, date, tm, typ = k
    return (date, tm if re.match(r"^\d{2}:\d{2}$", tm) else "", clinic, typ)


def _insert_in_order(ws, current, new_rows):
    """Inserts each new row under its own day, right after the last time that
    comes before it (before the first row that sorts after it) - not at the
    bottom of the tab. Falls back to appending if inserting fails."""
    plan = []
    for r in sorted(new_rows, key=_visit_sort_key):
        sk = _visit_sort_key(r)
        idx = len(current)
        for i in range(1, len(current)):
            ck = _visit_sort_key(current[i])
            if ck is not None and ck > sk:
                idx = i
                break
        current.insert(idx, list(r))
        plan.append((idx, list(r)))
    try:
        requests = [{"insertDimension": {
            "range": {"sheetId": ws.id, "dimension": "ROWS",
                      "startIndex": idx, "endIndex": idx + 1},
            "inheritFromBefore": idx > 1}} for idx, _ in plan]
        try:
            _api(ws.spreadsheet.batch_update, {"requests": requests})
        except AttributeError:
            for idx, _ in plan:
                _api(ws.insert_rows, [[""] * 5], row=idx + 1)
        values = [{"range": f"A{idx + 1}:E{idx + 1}", "values": [r]} for idx, r in plan]
        try:
            _api(ws.batch_update, values, value_input_option="RAW")
        except TypeError:
            _api(ws.batch_update, values)
        for idx, r in plan:
            log.info(f"  -  New row {idx + 1}: {r}")
        log.info(f"Inserted {len(plan)} new row(s) in date/time order")
    except Exception as e:
        log.error(f"Inserting rows in order failed ({e}) - adding them at the bottom")
        _api(ws.append_rows, [r for _, r in plan], value_input_option="RAW")


# =====================================================================
# 4. RANDOMISED CYCLE GAP
# =====================================================================


def cycle_sleep_seconds(min_minutes=3, max_minutes=5):
    """Random gap to sleep AFTER a run has finished: any number of seconds
    between min_minutes and max_minutes (default 3:00 - 5:00). Because each
    script's run takes a different time and the gap is random, the scripts
    drift apart instead of all hitting the sheet in the same second."""
    low = min(min_minutes, max_minutes) * 60
    high = max(min_minutes, max_minutes) * 60
    return random.randint(int(low), int(high))




def normalize_key_part(value):
    """strip + lowercase + collapse spaces; dates -> YYYY-MM-DD and times ->
    HH:MM (24h) when they parse, so '2026-10-26' == '10/26/2026' and
    '10:00 AM' == '10:00'. Use it inside key_fn."""
    from datetime import datetime
    text = " ".join(str(value or "").split())
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    for fmt in ("%I:%M %p", "%I:%M%p", "%H:%M", "%H:%M:%S", "%I:%M:%S %p"):
        try:
            return datetime.strptime(text.upper(), fmt).strftime("%H:%M")
        except ValueError:
            pass
    return text.lower()

# =====================================================================
# LOGGING
# =====================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]   # the per-clinic file is added in _configure_instance
)
log = logging.getLogger(__name__)


def step(msg: str):
    log.info(f"{'─' * 60}")                          
    log.info(f"  ▶  {msg}")
    log.info(f"{'─' * 60}")


def ok(msg: str):
    log.info(f"  ✔  {msg}")


def warn(msg: str):
    log.warning(f"  ⚠  {msg}")


def fail(msg: str):
    log.error(f"  ✖  {msg}")


# =====================================================================
# HUMAN-LIKE DELAY
# =====================================================================

def human_delay(min_sec: float = 1.5, max_sec: float = 4.0):
    delay = random.uniform(min_sec, max_sec)
    log.info(f"  ·  Waiting {delay:.1f}s ...")
    time.sleep(delay)


def smart_wait(driver, wait, by, value, timeout=10):
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
        return el
    except Exception:
        return None


def wait_for_page_stable(driver, timeout=8):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
    except Exception:
        pass


# =====================================================================
# SETTINGS
# =====================================================================

_settings_path = Path(__file__).resolve().parent.parent / "local_settings.json"
_local_settings = json.loads(_settings_path.read_text(encoding="utf-8")) if _settings_path.exists() else {}

WEBPT_URL = "https://app.webpt.com"
USERNAME  = os.environ.get("WEBPT_USERNAME", _local_settings.get("WEBPT_USERNAME", ""))
PASSWORD  = os.environ.get("WEBPT_PASSWORD", _local_settings.get("WEBPT_PASSWORD", ""))
OUTPUT_PATH  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "results.xlsx")
WAIT_TIMEOUT = 20


def env(name, default=None):
    """Environment overrides shared local settings."""
    return os.environ.get(name, _local_settings.get(name, default))


AWS_REGION = env("AWS_REGION") or env("AWS_DEFAULT_REGION", "us-east-1")
ARTIFACTS_BUCKET = env("WEBPT_ARTIFACTS_BUCKET")
ARTIFACTS_PREFIX = env("WEBPT_ARTIFACTS_PREFIX", "scheduled-visits").strip("/")
CLOUDWATCH_NAMESPACE = env("WEBPT_METRICS_NAMESPACE", "WebPT/ScheduledVisits")
GOOGLE_SECRET_ID = env("GOOGLE_SECRET_ID", "webpt/google")
# Bundled secret: webpt_username, webpt_password, google_service_account.
WEBPT_SECRET_ID = env("WEBPT_SECRET_ID")
# Workbooks and screenshots may contain patient information.
LOCAL_WORK_DIR = env("WEBPT_LOCAL_WORK_DIR") or os.path.dirname(os.path.abspath(__file__))
WEBPT_CREDENTIALS_FILE = env("WEBPT_CREDENTIALS_FILE", "credentials.xlsx")
WEBPT_URL = env("WEBPT_URL", WEBPT_URL)
LOG_DIR = env("WEBPT_LOG_DIR") or LOCAL_WORK_DIR
# Metrics are only sent when the namespace is set explicitly.
METRICS_ENABLED = bool(os.environ.get("WEBPT_METRICS_NAMESPACE")
                       or _local_settings.get("WEBPT_METRICS_NAMESPACE"))

# Linux / AWS: headless Chrome with a fixed binary and driver when present.
CHROME_BINARY_PATH = env("CHROME_BINARY_PATH", "/usr/local/bin/chrome")
CHROMEDRIVER_PATH = env("CHROMEDRIVER_PATH", "/usr/local/bin/chromedriver")
_default_headless = "1" if sys.platform.startswith("linux") and not os.environ.get("DISPLAY") else "0"
HEADLESS = str(env("WEBPT_HEADLESS", _default_headless)).strip().lower() in ("1", "true", "yes")

# Per-clinic WebPT accounts. A tab with three columns:
#   CLINIC NAME | USERNAME | PASSWORD
# By default it is read from the same spreadsheet the counts go to.
CREDENTIALS_SHEET_URL = env("CREDENTIALS_SHEET_URL")
CREDENTIALS_SHEET_TAB = env("CREDENTIALS_SHEET_TAB", "Account Credentials")
CREDENTIALS_CACHE_SECONDS = int(env("CREDENTIALS_CACHE_SECONDS", "600"))
CURRENT_CLINIC = None           # set by _configure_instance

_secrets_client = _s3_client = _cloudwatch_client = None
_bundled_secret = None


def _get_aws_secret(secret_id):
    global _secrets_client
    if _secrets_client is None:
        import boto3
        _secrets_client = boto3.client("secretsmanager", region_name=AWS_REGION)
    response = _secrets_client.get_secret_value(SecretId=secret_id)
    raw = response.get("SecretString")
    if raw is None:
        raw = response["SecretBinary"].decode("utf-8")
    return json.loads(raw)


def _load_bundled_secret():
    global _bundled_secret
    if _bundled_secret is None:
        secret = _get_aws_secret(WEBPT_SECRET_ID)
        if not secret.get("google_service_account"):
            raise ValueError("Bundled WebPT secret requires google_service_account")
        _bundled_secret = secret
    return _bundled_secret


def _clinic_key(name):
    """'Hell's Kitchen' == 'Hells Kitchen' == 'hells  kitchen'."""
    return re.sub(r"[^a-z0-9]", "", str(name or "").casefold())


def _header_key(name):
    return re.sub(r"[^a-z]", "", str(name or "").casefold())


_account_cache = {"at": 0.0, "value": None}


def _accounts_from_rows(rows):
    """rows: list of lists, first row = headers. Returns {clinic_key: (name, user, pwd)}."""
    if not rows:
        return {}
    headers = [_header_key(h) for h in rows[0]]

    def col(*names):
        for n in names:
            if n in headers:
                return headers.index(n)
        return None

    c_clinic = col("clinicname", "clinic")
    c_user = col("username", "user", "webptusername")
    c_pass = col("password", "webptpassword")
    if c_clinic is None or c_user is None or c_pass is None:
        raise ValueError(f"'{CREDENTIALS_SHEET_TAB}' must have the columns "
                         f"CLINIC NAME | USERNAME | PASSWORD (found: {rows[0]})")
    accounts = {}
    for r in rows[1:]:
        cell = lambda i: str(r[i]).strip() if i < len(r) and r[i] is not None else ""
        name, user, pwd = cell(c_clinic), cell(c_user), cell(c_pass)
        if name and user and pwd:
            accounts[_clinic_key(name)] = (name, user, pwd)
    return accounts


def _read_account_rows():
    """Reads the Account Credentials tab from Google Sheets; falls back to a
    local Excel file (WEBPT_CREDENTIALS_FILE) with the same columns."""
    url = CREDENTIALS_SHEET_URL or _google_sheet_settings()[0]
    if url:
        try:
            client = _gs_client()
            sh = _api(client.open_by_url, url)
            ws = _api(sh.worksheet, CREDENTIALS_SHEET_TAB)
            return _api(ws.get_all_values) or []
        except Exception as exc:
            warn(f"Could not read the '{CREDENTIALS_SHEET_TAB}' tab: {type(exc).__name__}: {exc}")
    path = Path(WEBPT_CREDENTIALS_FILE)
    if not path.is_absolute():
        path = _settings_path.parent / path
    if path.exists():
        sheets = pd.read_excel(path, sheet_name=None, dtype=str)
        frame = next((f for n, f in sheets.items()
                      if n.strip().casefold() == CREDENTIALS_SHEET_TAB.casefold()),
                     next(iter(sheets.values())))
        frame = frame.fillna("")
        return [list(frame.columns)] + frame.values.tolist()
    return []


def _account_for_clinic(clinic, force=False):
    """(username, password) for this clinic from the Account Credentials tab,
    or None. Cached for CREDENTIALS_CACHE_SECONDS, so a password changed in
    the sheet is picked up at the next login without restarting."""
    fresh = time.time() - _account_cache["at"] < CREDENTIALS_CACHE_SECONDS
    if force or _account_cache["value"] is None or not fresh:
        try:
            accounts = _accounts_from_rows(_read_account_rows())
            if accounts or _account_cache["value"] is None:
                _account_cache["value"] = accounts
                _account_cache["at"] = time.time()
        except Exception as exc:
            warn(f"Account Credentials not usable: {exc}")
            if _account_cache["value"] is None:
                _account_cache["value"] = {}
    accounts = _account_cache["value"] or {}
    found = accounts.get(_clinic_key(clinic))
    if not found:
        return None
    _, user, pwd = found
    shared = sorted(n for n, u, _ in accounts.values()
                    if u.casefold() == user.casefold() and _clinic_key(n) != _clinic_key(clinic))
    if shared:
        warn(f"Username '{user}' is also used for {shared} - two sessions on one "
             f"account log each other out ('oust them')")
    return user, pwd


def _load_webpt_credentials(force=False):
    global USERNAME, PASSWORD
    # 1) this clinic's row in the Account Credentials tab
    if CURRENT_CLINIC:
        account = _account_for_clinic(CURRENT_CLINIC, force=force)
        if account:
            USERNAME, PASSWORD = account
            return
        warn(f"No row for '{CURRENT_CLINIC}' in '{CREDENTIALS_SHEET_TAB}' - "
             f"falling back to the other credential settings")
    # 2) the bundled AWS secret
    if WEBPT_SECRET_ID:
        secret = _load_bundled_secret()
        if secret.get("webpt_username") and secret.get("webpt_password"):
            USERNAME = secret["webpt_username"]
            PASSWORD = secret["webpt_password"]
            return
    # 3) env / local_settings / first row of the local credentials file
    if not (USERNAME and PASSWORD):
        path = Path(WEBPT_CREDENTIALS_FILE)
        if not path.is_absolute():
            path = _settings_path.parent / path
        if path.exists():
            row = pd.read_excel(path, dtype=str).fillna("").iloc[0]
            values = {str(key).strip().lower(): value for key, value in row.items()}
            USERNAME = USERNAME or values.get("webpt_username", values.get("username", ""))
            PASSWORD = PASSWORD or values.get("webpt_password", values.get("password", ""))


def _google_secret_credentials(credentials_path):
    if WEBPT_SECRET_ID:
        return _load_bundled_secret()["google_service_account"]
    # Preserve local Google credentials unless a secret is explicitly selected.
    if env("GOOGLE_SECRET_ID") or not credentials_path.exists():
        return _get_aws_secret(GOOGLE_SECRET_ID)
    return None




def _s3_upload(local_path, key_suffix):
    """Uploads a file to s3://WEBPT_ARTIFACTS_BUCKET/<prefix>/<suffix> with
    KMS encryption. Does nothing when no bucket is configured. Never raises."""
    global _s3_client
    if not ARTIFACTS_BUCKET:
        return False
    try:
        if _s3_client is None:
            import boto3
            _s3_client = boto3.client("s3", region_name=AWS_REGION)
        key = f"{ARTIFACTS_PREFIX}/{key_suffix}".lstrip("/")
        _s3_client.upload_file(str(local_path), ARTIFACTS_BUCKET, key,
                               ExtraArgs={"ServerSideEncryption": "aws:kms"})
        log.info(f"  -  Uploaded {Path(local_path).name} -> s3://{ARTIFACTS_BUCKET}/{key}")
        return True
    except Exception as exc:
        warn(f"S3 upload failed for {local_path}: {type(exc).__name__}: {exc}")
        return False


def put_metric(name, value, unit="Count"):
    """CloudWatch metric per clinic. Only when WEBPT_METRICS_NAMESPACE is set."""
    global _cloudwatch_client
    if not METRICS_ENABLED:
        return
    try:
        if _cloudwatch_client is None:
            import boto3
            _cloudwatch_client = boto3.client("cloudwatch", region_name=AWS_REGION)
        _cloudwatch_client.put_metric_data(
            Namespace=CLOUDWATCH_NAMESPACE,
            MetricData=[{"MetricName": name, "Value": float(value), "Unit": unit,
                         "Dimensions": [{"Name": "Clinic",
                                         "Value": CURRENT_CLINIC or "all"}]}])
    except Exception as exc:
        log.debug(f"CloudWatch put_metric_data failed for {name}: {exc}")


# =====================================================================
# NAME / DATE HELPERS
# =====================================================================

def get_first_name(full_name: str) -> str:
    full_name = full_name.strip()
    full_name = re.sub(r'[,،‚]', ',', full_name)
    if "," in full_name:
        return full_name.split(",", 1)[1].strip()
    parts = full_name.split()
    return parts[-1] if parts else full_name


def _get_last_name(full_name: str) -> str:
    full_name = full_name.strip()
    full_name = re.sub(r'[,،‚]', ',', full_name)
    if "," in full_name:
        return full_name.split(",", 1)[0].strip()
    parts = full_name.split()
    return " ".join(parts[:-1]) if len(parts) > 1 else full_name


def normalize_date(date_str: str) -> str:
    try:
        clean = re.sub(r'\(.*?\)', '', date_str.strip())
        clean = re.sub(r'(\d+)(st|nd|rd|th)', r'\1', clean).strip()
        return pd.to_datetime(clean).strftime("%m/%d/%Y")
    except Exception:
        return date_str.strip()


# =====================================================================
# DRIVER
# =====================================================================

def create_driver() -> webdriver.Chrome:
    step("Launching Chrome browser")
    options = webdriver.ChromeOptions()

    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    if HEADLESS:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
    else:
        options.add_argument("--start-maximized")
    if CHROME_BINARY_PATH and os.path.exists(CHROME_BINARY_PATH):
        options.binary_location = CHROME_BINARY_PATH

    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    user_agents = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    ]
    options.add_argument(f"--user-agent={random.choice(user_agents)}")

    if CHROMEDRIVER_PATH and os.path.exists(CHROMEDRIVER_PATH):
        service = Service(CHROMEDRIVER_PATH)
    else:
        service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    if HEADLESS:
        driver.set_window_size(1920, 1080)

    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
            Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3] });
            Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] });
        """
    })

    ok("Browser launched successfully")
    return driver


# =====================================================================
# LOGIN
# =====================================================================

def _visible_element(driver, by, value):
    try:
        for el in driver.find_elements(by, value):
            if el.is_displayed():
                return el
    except Exception:
        pass
    return None


def _dashboard_ready(driver) -> bool:
    return bool(_visible_element(driver, By.LINK_TEXT, "Advanced Search")
                or _visible_element(driver, By.ID, "ClinicChange")
                or driver.find_elements(By.ID, "ClinicChange"))


def _is_login_page(driver) -> bool:
    """True when WebPT sent us back to the sign-in page (session ended)."""
    try:
        if _dashboard_ready(driver):
            return False
        if (_visible_element(driver, By.ID, "username")
                or _visible_element(driver, By.ID, "password")):
            return True
        url = (driver.current_url or "").lower()
        return any(k in url for k in ("login", "signin", "sign-in", "/auth"))
    except Exception:
        return False


def _submit_login_step(driver, field):
    button = _visible_element(driver, By.CSS_SELECTOR, "button[type='submit']")
    if button is not None:
        driver.execute_script("arguments[0].click();", button)
    else:
        field.send_keys(Keys.ENTER)


def _complete_login(driver, timeout=90):
    """Types the username and password on whatever login step is showing,
    answers the 'oust them' prompt, and waits for the dashboard."""
    _load_webpt_credentials()
    if not USERNAME or not PASSWORD:
        raise RuntimeError("WEBPT_USERNAME / WEBPT_PASSWORD are empty")
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if _dashboard_ready(driver):
            return
        clicked_prompt = False
        for prompt_text in ["Yes, oust them!", "Yes"]:
            btn = _visible_element(driver, By.XPATH, f"//button[contains(text(),'{prompt_text}')]")
            if btn is not None:
                driver.execute_script("arguments[0].click();", btn)
                ok(f"Dismissed prompt: '{prompt_text}'")
                clicked_prompt = True
                break
        if clicked_prompt:
            time.sleep(2)
            continue
        password = _visible_element(driver, By.ID, "password")
        if password is not None:
            password.clear()
            password.send_keys(PASSWORD)
            _submit_login_step(driver, password)
            log.info("  →  Password entered")
            time.sleep(3)
            continue
        username = _visible_element(driver, By.ID, "username")
        if username is not None:
            username.clear()
            username.send_keys(USERNAME)
            _submit_login_step(driver, username)
            log.info("  →  Username entered")
            time.sleep(2)
            continue
        time.sleep(1)
    _account_cache["at"] = 0.0      # re-read Account Credentials on the next try
    raise RuntimeError("Login did not reach the dashboard in time")


def login(driver, wait):
    step("Logging in to WebPT")

    driver.get(WEBPT_URL)
    _complete_login(driver)
    ok("Login successful — dashboard ready")

    # ── Open View Schedule once, right after login ──────────────────
    # From this point on, the script stays on the View Schedule page
    # for the entire run. It does NOT return to the dashboard between
    # patients. The flow is:
    #   View Schedule → Clinic → Date → Search → Close popup
    #   → (next patient) Clinic (skipped if same) → Date → Search → ...
    step("Opening View Schedule (one-time)")
    if not _open_view_schedule(driver, wait):
        raise RuntimeError("Could not open View Schedule after login; collection stopped.")


def _open_view_schedule(driver, wait):
    """
    Opens the View Schedule page from the current page.
    Tries multiple strategies. Verifies success by checking for the
    ClinicChange dropdown.
    """
    for attempt in range(1, 4):
        log.info(f"  →  Opening View Schedule (attempt {attempt}/3) ...")

        clicked = False
        for locator in [
            (By.LINK_TEXT,          "View Schedule"),
            (By.PARTIAL_LINK_TEXT,  "View Schedule"),
            (By.XPATH,              "//a[contains(text(),'View Schedule')]"),
            (By.XPATH,              "//*[contains(text(),'View Schedule')]"),
            (By.CSS_SELECTOR,       "a[href*='scheduler']"),
        ]:
            try:
                el = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(locator))
                driver.execute_script("arguments[0].click();", el)
                wait_for_page_stable(driver, timeout=20)
                human_delay(1.0, 2.5)
                clicked = True
                break
            except Exception:
                continue

        # Check if View Schedule is now open
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ClinicChange"))
            )
            ok(f"View Schedule opened (attempt {attempt})")
            return True
        except Exception:
            pass

        # Fallback: direct URL
        try:
            log.info(f"  →  Trying direct URL ...")
            driver.get(WEBPT_URL + "/scheduler/")
            wait_for_page_stable(driver, timeout=20)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ClinicChange"))
            )
            ok(f"View Schedule opened via direct URL (attempt {attempt})")
            return True
        except Exception:
            pass

        if attempt < 3:
            log.info(f"  →  Will retry ...")
            time.sleep(3)

    warn("Could not open View Schedule after 3 attempts")
    return False


# =====================================================================
# CLINIC SELECTION
# =====================================================================

_current_clinic = None

# Tracks whether the No Show + Cancelled filters are currently active.
# When they are, we can reuse them for the next patient (no need to
# toggle them off and on again). They reset whenever the clinic changes.
_extra_filters_active = False


def _get_clinic_select_element(driver):
    try:
        return driver.find_element(By.ID, "ClinicChange")
    except Exception:
        try:
            return driver.find_element(By.NAME, "ClinicChange")
        except Exception:
            return None


def _read_current_clinic_from_page(driver) -> str:
    try:
        sel_el = _get_clinic_select_element(driver)
        if sel_el is None:
            return ""
        try:
            sel = Select(sel_el)
            first_selected = sel.first_selected_option
            text = first_selected.text.strip()
            if text:
                return text
        except Exception:
            pass
        try:
            text = driver.execute_script(
                "var s=arguments[0]; "
                "if(s.selectedIndex<0) return ''; "
                "return s.options[s.selectedIndex].text || '';",
                sel_el
            )
            return (text or "").strip()
        except Exception:
            pass
    except Exception:
        pass
    return ""


def select_clinic_in_schedule(driver, clinic_name: str):
    global _current_clinic, _extra_filters_active

    target_clinic = clinic_name.strip()
    clinic_lower  = target_clinic.lower()

    if _current_clinic and _current_clinic.lower() == clinic_lower:
        ok(f"Clinic '{target_clinic}' already selected (cached) — skipping")
        return

    current_on_page = _read_current_clinic_from_page(driver)
    log.info(f"  →  Currently selected clinic: '{current_on_page}'")
    if current_on_page and (
        current_on_page.lower() == clinic_lower
        or clinic_lower in current_on_page.lower()
        or current_on_page.lower() in clinic_lower
    ):
        ok(f"Clinic '{current_on_page}' already shown — skipping")
        _current_clinic = current_on_page
        return

    # Clinic is going to CHANGE — invalidate the extra-filters state
    # because the new clinic's view will reset all filter selections.
    if _extra_filters_active:
        log.info("  →  Clinic is changing — filters will reset")
        _extra_filters_active = False

    log.info(f"  →  Changing clinic to '{target_clinic}' via visual dropdown ...")

    trigger_clicked = False

    try:
        sel_el = _get_clinic_select_element(driver)
        candidates = driver.find_elements(
            By.XPATH,
            "//input[@type='text' or not(@type)]"
        )
        for cand in candidates:
            try:
                if not cand.is_displayed():
                    continue
                val = (cand.get_attribute("value") or "").strip()
                if current_on_page and val == current_on_page:
                    log.info(f"  →  Found visible clinic input (value='{val}') — clicking ...")
                    try:
                        cand.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", cand)
                    time.sleep(1)
                    trigger_clicked = True
                    break
            except Exception:
                continue
    except Exception as e:
        log.debug(f"Strategy A error: {e}")

    if not trigger_clicked:
        try:
            for sel_str in [
                "img[class*='trigger']",
                "div[class*='trigger']",
                "img[class*='arrow']",
                "div[class*='arrow-trigger']",
                "*[class*='x-form-trigger']",
            ]:
                triggers = driver.find_elements(By.CSS_SELECTOR, sel_str)
                for trg in triggers:
                    try:
                        if not trg.is_displayed():
                            continue
                        log.info(f"  →  Trying trigger '{sel_str}' ...")
                        try:
                            trg.click()
                        except Exception:
                            driver.execute_script("arguments[0].click();", trg)
                        time.sleep(1)
                        trigger_clicked = True
                        break
                    except Exception:
                        continue
                if trigger_clicked:
                    break
        except Exception as e:
            log.debug(f"Strategy B error: {e}")

    if not trigger_clicked:
        try:
            sel_el = _get_clinic_select_element(driver)
            if sel_el:
                driver.execute_script(
                    """
                    var s = arguments[0];
                    var p = s.parentElement;
                    if (p) {
                        var rect = p.getBoundingClientRect();
                        var ev = new MouseEvent('click', {
                            bubbles: true, cancelable: true, view: window,
                            clientX: rect.right - 10,
                            clientY: rect.top + rect.height / 2
                        });
                        p.dispatchEvent(ev);
                    }
                    """,
                    sel_el
                )
                time.sleep(1)
                trigger_clicked = True
                log.info("  →  Clicked dropdown via parent element")
        except Exception as e:
            log.debug(f"Strategy C error: {e}")

    if not trigger_clicked:
        warn("Could not click clinic dropdown trigger")
        return

    time.sleep(0.8)

    search_typed = False
    try:
        all_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='text'], input:not([type])")
        for inp in all_inputs:
            try:
                if not inp.is_displayed():
                    continue
                val = (inp.get_attribute("value") or "").strip()
                if val:
                    continue
                cls = (inp.get_attribute("class") or "").lower()
                iid  = (inp.get_attribute("id") or "").lower()
                name = (inp.get_attribute("name") or "").lower()
                if "patient" in iid or "patient" in name or "date" in iid:
                    continue
                log.info(f"  →  Typing '{target_clinic}' into search box (id='{iid}', class='{cls[:40]}')")
                inp.clear()
                time.sleep(0.3)
                inp.send_keys(target_clinic)
                time.sleep(1.2)
                search_typed = True
                break
            except Exception:
                continue
    except Exception as e:
        log.debug(f"Search type error: {e}")

    if not search_typed:
        log.info("  →  No search box found; will scan the visible list directly")

    matched = False
    try:
        item_selectors = [
            "div.x-combo-list-item",
            "div.x-boundlist-item",
            "li.x-boundlist-item",
            "div[class*='combo-list-item']",
            "div[class*='boundlist-item']",
            "div[class*='list-item']",
        ]
        all_items = []
        for isel in item_selectors:
            found = driver.find_elements(By.CSS_SELECTOR, isel)
            for f in found:
                if f.is_displayed():
                    all_items.append(f)

        log.info(f"  →  Visible list items: {len(all_items)}")

        item_texts = []
        for it in all_items[:25]:
            try:
                t = it.text.strip()
                if t:
                    item_texts.append(t)
            except Exception:
                pass
        log.info(f"  →  Visible items (first 25): {item_texts}")

        for it in all_items:
            try:
                t = it.text.strip()
                if t.lower() == clinic_lower:
                    log.info(f"  →  Clicking exact match: '{t}'")
                    driver.execute_script("arguments[0].click();", it)
                    time.sleep(2.5)
                    wait_for_page_stable(driver, timeout=10)
                    _current_clinic = t
                    ok(f"Clinic selected: '{t}'")
                    matched = True
                    return
            except Exception:
                continue

        if not matched:
            for it in all_items:
                try:
                    t = it.text.strip()
                    if not t:
                        continue
                    if clinic_lower in t.lower() or t.lower() in clinic_lower:
                        log.info(f"  →  Clicking substring match: '{t}'")
                        driver.execute_script("arguments[0].click();", it)
                        time.sleep(2.5)
                        wait_for_page_stable(driver, timeout=10)
                        _current_clinic = t
                        ok(f"Clinic selected: '{t}'")
                        matched = True
                        return
                except Exception:
                    continue
    except Exception as e:
        log.debug(f"Item click error: {e}")

    if not matched:
        log.info("  →  Visual dropdown approach didn't match; trying hidden <select> fallback")
        try:
            sel_el = _get_clinic_select_element(driver)
            if sel_el is None:
                warn("ClinicChange <select> not found")
                return

            js_options = driver.execute_script(
                "var s=arguments[0],r=[];for(var i=0;i<s.options.length;i++){var o=s.options[i];r.push({v:o.value,t:o.text});}return r;",
                sel_el
            ) or []

            log.info(f"  →  Hidden <select> has {len(js_options)} options")

            matched_value = None
            matched_text  = None
            for o in js_options:
                if (o.get('t') or '').strip().lower() == clinic_lower:
                    matched_value = o['v']; matched_text = o['t'].strip(); break
            if not matched_value:
                for o in js_options:
                    t = (o.get('t') or '').strip().lower()
                    if t and (clinic_lower in t or t in clinic_lower):
                        matched_value = o['v']; matched_text = o['t'].strip(); break

            if matched_value:
                driver.execute_script(
                    "var s=arguments[0],v=arguments[1];"
                    "for(var i=0;i<s.options.length;i++){if(s.options[i].value===v){s.selectedIndex=i;break;}}"
                    "s.value=v;"
                    "var ev;try{ev=new Event('change',{bubbles:true,cancelable:true});}"
                    "catch(e){ev=document.createEvent('HTMLEvents');ev.initEvent('change',true,true);}"
                    "s.dispatchEvent(ev);",
                    sel_el, matched_value
                )
                time.sleep(3)
                wait_for_page_stable(driver, timeout=10)
                _current_clinic = matched_text
                log.info(f"Clinic change requested via hidden <select>: '{matched_text}'")
                return
        except Exception as e:
            warn(f"Hidden <select> fallback failed: {e}")

    warn(f"Could not select clinic '{clinic_name}' — using current")


# =====================================================================
# CALENDAR NAVIGATION
# =====================================================================

def navigate_schedule_to_date(driver, target_date: str):
    try:
        target_normalized = normalize_date(target_date)
        target_dt         = pd.to_datetime(target_normalized)
        target_day        = str(target_dt.day)
        target_month_abbr = target_dt.strftime("%b")
        target_year       = target_dt.strftime("%Y")
        target_month_full = target_dt.strftime("%b %Y")

        log.info(f"  →  Target: {target_month_full} day {target_day}")

        current = _read_calendar_header(driver)
        log.info(f"  →  Current calendar header: '{current}'")

        if current and target_month_abbr.lower() in current.lower() and target_year in current:
            ok(f"Already on correct month: {current}")
        else:
            reached = _calendar_select_month_year(driver, target_month_abbr, target_year)
            if not reached:
                warn(f"Arrow navigation failed — cannot safely navigate to {target_month_full}")
                return

        time.sleep(1)
        clicked = _calendar_click_day(driver, target_day, target_month_full)

        if not clicked:
            warn(f"Could not click day {target_day} in {target_month_full}")
            return

        time.sleep(2)

        verify_header = _read_calendar_header(driver)
        if verify_header and (target_month_abbr.lower() not in verify_header.lower() or target_year not in verify_header):
            warn(f"After clicking day {target_day}, calendar shows '{verify_header}' — expected '{target_month_full}'")
            warn("Attempting recovery via arrow navigation ...")
            recovered = _calendar_select_month_year(driver, target_month_abbr, target_year)
            if recovered:
                time.sleep(1)
                _calendar_click_day(driver, target_day, target_month_full)
                time.sleep(2)
                ok(f"Recovery successful: navigated to {target_normalized}")
            else:
                warn("Recovery failed — proceeding with best effort")
        else:
            ok(f"Schedule navigated to: {target_normalized}")

    except Exception as e:
        warn(f"navigate_schedule_to_date error: {e}")


def _read_calendar_header(driver) -> str:
    pattern = re.compile(
        r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}',
        re.IGNORECASE
    )

    try:
        tables = driver.find_elements(By.TAG_NAME, "table")
        for tbl in tables:
            tbl_text = tbl.text
            if "Su" in tbl_text and "Mo" in tbl_text and "Sa" in tbl_text:
                rows = tbl.find_elements(By.TAG_NAME, "tr")
                if rows:
                    header_text = rows[0].text.strip()
                    m = pattern.search(header_text)
                    if m:
                        return m.group(0)
                for cell in tbl.find_elements(By.XPATH, ".//td | .//th"):
                    text = cell.text.strip()
                    if re.match(r'^\d{1,2}$', text):
                        continue
                    m = pattern.search(text)
                    if m:
                        return m.group(0)
    except Exception:
        pass

    for tag in ("td", "th"):
        try:
            elements = driver.find_elements(By.TAG_NAME, tag)
            for el in elements:
                text = el.text.strip()
                if re.match(r'^\d{1,2}$', text):
                    continue
                m = pattern.search(text)
                if m:
                    return m.group(0)
        except Exception:
            continue
    return ""


def _calendar_select_month_year(driver, target_month_str: str, target_year_str: str) -> bool:
    target_month_full = f"{target_month_str} {target_year_str}"
    log.info(f"  →  Arrow-navigating calendar to: {target_month_full}")

    try:
        target_dt = pd.to_datetime(f"01 {target_month_str} {target_year_str}")
    except Exception as e:
        warn(f"Cannot parse target date '{target_month_full}': {e}")
        return False

    for attempt in range(24):
        time.sleep(1)
        current = _read_calendar_header(driver)
        log.info(f"  →  Calendar header (attempt {attempt+1}): '{current}'")

        if not current:
            warn("Cannot read calendar header")
            time.sleep(1.5)
            continue

        if target_month_str.lower() in current.lower() and target_year_str in current:
            ok(f"Correct month reached: {current}")
            return True

        try:
            cur_dt     = pd.to_datetime("01 " + current)
            go_forward = target_dt > cur_dt
        except Exception:
            go_forward = True

        if go_forward:
            _click_calendar_arrow(driver, ">")
        else:
            _click_calendar_arrow(driver, "<")

    warn(f"Could not reach {target_month_full} after 24 attempts")
    return False


def _click_calendar_arrow(driver, symbol: str):
    try:
        tables = driver.find_elements(By.TAG_NAME, "table")
        cal_table = None
        for tbl in tables:
            txt = tbl.text
            if "Mo" in txt and "Su" in txt and "Sa" in txt:
                cal_table = tbl
                break

        if cal_table:
            for a in cal_table.find_elements(By.TAG_NAME, "a"):
                try:
                    t = a.text.strip()
                    if t == symbol and chr(10) not in a.text:
                        clicked_ok = False
                        try:
                            a.click()
                            clicked_ok = True
                            log.info(f"  →  Clicked '{symbol}' via native .click()")
                        except Exception:
                            pass
                        if not clicked_ok:
                            try:
                                from selenium.webdriver.common.action_chains import ActionChains
                                ActionChains(driver).move_to_element(a).click().perform()
                                clicked_ok = True
                                log.info(f"  →  Clicked '{symbol}' via ActionChains")
                            except Exception:
                                pass
                        if not clicked_ok:
                            try:
                                driver.execute_script("""
                                    var el = arguments[0];
                                    var e1 = new MouseEvent('mousedown', {bubbles:true, cancelable:true});
                                    var e2 = new MouseEvent('mouseup',   {bubbles:true, cancelable:true});
                                    var e3 = new MouseEvent('click',     {bubbles:true, cancelable:true});
                                    el.dispatchEvent(e1);
                                    el.dispatchEvent(e2);
                                    el.dispatchEvent(e3);
                                """, a)
                                clicked_ok = True
                                log.info(f"  →  Clicked '{symbol}' via MouseEvent dispatch")
                            except Exception:
                                pass
                        if clicked_ok:
                            return
                except Exception:
                    continue

            for td in cal_table.find_elements(By.TAG_NAME, "td"):
                try:
                    raw = td.text
                    stripped = raw.strip()
                    if stripped == symbol and chr(10) not in raw:
                        driver.execute_script("arguments[0].click();", td)
                        log.info(f"  →  Clicked '{symbol}' arrow via <td> in calendar table")
                        return
                except Exception:
                    continue

        for tag in ("a", "td", "span", "input", "button"):
            for el in driver.find_elements(By.TAG_NAME, tag):
                try:
                    raw = el.text
                    stripped = raw.strip()
                    val = (el.get_attribute("value") or "").strip()
                    if (stripped == symbol or val == symbol) and chr(10) not in raw:
                        driver.execute_script("arguments[0].click();", el)
                        log.info(f"  →  Clicked '{symbol}' arrow via <{tag}> (page scan)")
                        return
                except Exception:
                    continue

        warn(f"Arrow '{symbol}' not found in calendar")

        try:
            js = f"""
            var all = document.querySelectorAll('td, a, span, button, input');
            for (var i = 0; i < all.length; i++) {{
                var el = all[i];
                var t = (el.innerText || el.textContent || el.value || '').trim();
                if (t === '{symbol}' && !t.includes('\\n')) {{
                    el.click();
                    return true;
                }}
            }}
            return false;
            """
            result = driver.execute_script(js)
            if result:
                log.info(f"  →  Clicked '{symbol}' arrow via JavaScript")
            else:
                warn(f"  →  JavaScript arrow click also failed for '{symbol}'")
        except Exception as e2:
            warn(f"JavaScript arrow fallback failed: {e2}")

    except Exception as e:
        warn(f"_click_calendar_arrow('{symbol}') error: {e}")


def _calendar_click_day(driver, target_day: str, target_month_year: str = "") -> bool:
    try:
        cal_table = None
        tables = driver.find_elements(By.TAG_NAME, "table")
        for tbl in tables:
            tbl_text = tbl.text
            if "Su" in tbl_text and "Mo" in tbl_text and "Sa" in tbl_text:
                cal_table = tbl
                break

        if not cal_table:
            warn(f"Calendar table not found — cannot safely click day {target_day}")
            return False

        day_links = cal_table.find_elements(By.TAG_NAME, "a")
        log.info(f"  →  Calendar table has {len(day_links)} <a> links")

        candidates = []
        for link in day_links:
            try:
                link_text = link.text.strip()
                if link_text != target_day:
                    continue

                try:
                    parent_td = link.find_element(By.XPATH, "./ancestor::td[1]")
                    td_class = (parent_td.get_attribute("class") or "").lower()
                except Exception:
                    td_class = ""

                is_other_month = ("prevday" in td_class or "nextday" in td_class
                                   or "outside" in td_class or "other-month" in td_class
                                   or "x-datepicker-prevday" in td_class
                                   or "x-datepicker-nextday" in td_class)
                if is_other_month:
                    log.info(f"  →  Skipping day {target_day} from adjacent month (class='{td_class}')")
                    continue

                link_class = (link.get_attribute("class") or "").lower()
                if "prevday" in link_class or "nextday" in link_class:
                    log.info(f"  →  Skipping day {target_day} (link class='{link_class}')")
                    continue

                candidates.append((link, td_class, link_class))
            except Exception:
                continue

        log.info(f"  →  Found {len(candidates)} valid candidate(s) for day {target_day}")

        if not candidates:
            warn(f"Day {target_day} not found in current month inside the calendar table")
            for link in day_links:
                if link.text.strip() == target_day:
                    log.info(f"  →  Fallback: clicking first '{target_day}' link without class filter")
                    driver.execute_script("arguments[0].click();", link)
                    return True
            return False

        link, td_class, link_class = candidates[0]
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", link)
        ok(f"Clicked day {target_day} in calendar (active month)")
        return True

    except Exception as e:
        warn(f"_calendar_click_day error: {e}")
        return False


# =====================================================================
# PROCESS A SINGLE APPOINTMENT
# (Same flow as the original — but only extracts Status + Appt_Date)
# =====================================================================

def _process_single_appointment(driver, full_name: str, matched_el, icon_status: str):
    """
    Processes ONE matched appointment from the schedule:

    Logic:
      1. THE ICON IS THE PRIMARY SOURCE OF STATUS.
         The icon next to the patient's name (√, «, Φ, ?) decides the
         status BEFORE the popup is even opened.

      2. READ TREATING THERAPIST BEFORE CLICKING.
         The appointment's X-position tells us which therapist column
         it's under. This must be read before the popup opens (since
         the popup might obscure the schedule).

      3. THE POPUP LINK PROVIDES DATE/TIME.
         When the popup opens, we extract the date and time from the
         status link inside it.

      4. NO Scheduler History navigation.

    Returns (status, appt_date, therapist)
    """
    # The icon already gave us the status — that's our answer.
    final_status = icon_status if icon_status else "Other"
    appt_date    = ""
    therapist    = ""
    appt_type    = ""

    log.info(f"  →  Icon-detected status (primary): '{final_status}'")

    # ── Step 0a: Read the Treating Therapist BEFORE clicking ─────────
    # The popup might cover schedule columns once opened, so we must
    # read the therapist's name first while the appointment is still
    # visible in its column.
    try:
        therapist = _get_therapist_for_appointment(driver, matched_el)
        if therapist:
            ok(f"  →  Treating Therapist: '{therapist}'")
        else:
            log.info("  →  Treating Therapist: (could not determine)")
    except Exception as e:
        log.debug(f"Therapist read error: {e}")

    # ── Step 0b: Read the Appointment Type from tooltip (hover) ─────
    # Hover over the appointment to trigger its tooltip (e.g. "School INDV 9:30am"),
    # then take only the part before the time → "School INDV".
    # Must be done BEFORE clicking, since clicking opens the popup.
    try:
        appt_type = _get_tooltip_for_appointment(driver, matched_el)
        if appt_type:
            ok(f"  →  Appointment Type: '{appt_type}'")
        else:
            log.info("  →  Appointment Type: (could not determine)")
    except Exception as e:
        log.debug(f"Appointment type read error: {e}")

    # ── Step 1: Click the appointment to open the popup ─────────────
    try:
        driver.execute_script("arguments[0].scrollIntoView(true);", matched_el)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", matched_el)
        time.sleep(2)
    except Exception as e:
        log.debug(f"Click error: {e}")
        warn(f"Could not click appointment for '{full_name}'")
        return final_status, "", therapist, appt_type

    # ── Step 2: Wait briefly for popup to render ────────────────────
    time.sleep(1.2)

    # ── Step 3: Find the status link in the popup ───────────────────
    # The popup contains a link like:
    #   "Checked Out by Sandra Emad 04/29/2026 08:58 AM"
    #   "Checked In by Hagar Alaa 04/27/2026 09:49 AM"
    #   "No Show by Mike 04/22/2026 10:00 AM"
    #   "Cancelled by Sarah 04/22/2026 02:30 PM"
    #   "Edited by Elsayed Moussa 05/05/2026 02:13 PM"
    #   "Created by Hagar Alaa 05/09/2026 10:30 AM"
    popup_link_text = ""
    try:
        all_links = driver.find_elements(By.TAG_NAME, "a")
        for el in all_links:
            try:
                if not el.is_displayed():
                    continue
                link_href = (el.get_attribute("href") or "")
                # The status link points to scheduler/history with date=...&uid=...
                if ("scheduler/history" in link_href or "schedulerHistory" in link_href) \
                        and ("date=" in link_href or "uid=" in link_href):
                    popup_link_text = el.text.strip()
                    log.info(f"  →  Found popup link: '{popup_link_text[:150]}'")
                    break
            except Exception:
                continue

        # Fallback: any visible link whose text matches "<event> by <user> <date> <time>"
        if not popup_link_text:
            for el in all_links:
                try:
                    if not el.is_displayed():
                        continue
                    txt = el.text.strip()
                    if not txt:
                        continue
                    tl = txt.lower()
                    if (("checked out" in tl or "checked in" in tl or "no show" in tl
                         or "cancelled" in tl or "canceled" in tl
                         or "edited" in tl or "created" in tl)
                            and " by " in tl
                            and re.search(r'\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}', txt)):
                        popup_link_text = txt
                        log.info(f"  →  Found popup link (fallback): '{popup_link_text[:150]}'")
                        break
                except Exception:
                    continue
    except Exception as e:
        log.debug(f"Popup link search error: {e}")

    # ── Step 4: Extract Date + Time from the link (status comes from icon) ─
    # We do NOT change final_status based on the link's event word —
    # the icon is authoritative. We only take the date & time.
    if popup_link_text:
        try:
            # Try the strict pattern first:
            #   "<Event> by <User> MM/DD/YYYY HH:MM AM"
            m = re.match(
                r'^(Checked Out|Checked In|No Show|Cancelled|Edited|Created)\s+by\s+(.+?)\s+(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})\s+(\d{1,2}:\d{2}\s*[AP]M)\s*$',
                popup_link_text,
                re.IGNORECASE,
            )
            if m:
                event_word = m.group(1).strip()
                date_part  = m.group(3).strip()
                time_part  = m.group(4).strip()

                try:
                    appt_date = pd.to_datetime(f"{date_part} {time_part}").strftime("%Y-%m-%d %I:%M %p")
                except Exception:
                    appt_date = f"{date_part} {time_part}"

                log.info(f"  →  Link event word: '{event_word}' (ignored — icon decides status)")
                ok(f"  ✓ Date='{date_part}' | Time='{time_part}'")
            else:
                # Loose extraction: just find any date + time anywhere in the text
                log.info(f"  →  Strict pattern didn't match — trying loose extraction")
                dt_match = re.search(
                    r'(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})\s+(\d{1,2}:\d{2}\s*[AP]M)',
                    popup_link_text,
                    re.IGNORECASE
                )
                if dt_match:
                    date_part = dt_match.group(1).strip()
                    time_part = dt_match.group(2).strip()
                    try:
                        appt_date = pd.to_datetime(f"{date_part} {time_part}").strftime("%Y-%m-%d %I:%M %p")
                    except Exception:
                        appt_date = f"{date_part} {time_part}"
                    ok(f"  ✓ Loose extract: Date='{date_part}' | Time='{time_part}'")
                else:
                    log.info(f"  →  No date/time found in link — appt_date will be empty")
        except Exception as e:
            log.debug(f"  →  Link parse error: {e}")
    else:
        log.info("  →  No popup link found — appt_date will be empty")

    # ── Step 5: Close the popup and return to schedule ──────────────
    _ensure_on_schedule_view(driver)

    ok(f"FINAL: Status='{final_status}' | Date='{appt_date}' | Therapist='{therapist}' | Type='{appt_type}'")
    return final_status, appt_date, therapist, appt_type


def _close_popup_if_open(driver):
    from selenium.webdriver.common.keys import Keys

    try:
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys(Keys.ESCAPE)
        time.sleep(0.4)
    except Exception:
        pass

    close_selectors = [
        "div.x-tool-close",
        "img.x-tool-close",
        "div.x-tool-img.x-tool-close",
        "a.close-button",
        "div.close-button",
        "img.closeButton",
        "img[src*='close']",
        "*[class*='x-tool-close']",
        "*[title='Close']",
        "*[aria-label='Close']",
        "img[alt='close']",
        "img[alt='Close']",
    ]

    for sel in close_selectors:
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, sel)
            for el in elements:
                try:
                    if el.is_displayed():
                        driver.execute_script("arguments[0].click();", el)
                        time.sleep(0.5)
                        return
                except Exception:
                    continue
        except Exception:
            continue

    try:
        x_buttons = driver.find_elements(
            By.XPATH,
            "//*[normalize-space(text())='×' or normalize-space(text())='X' or normalize-space(text())='✕']"
        )
        for btn in x_buttons:
            try:
                if btn.is_displayed() and btn.size.get('width', 0) < 50:
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.5)
                    return
            except Exception:
                continue
    except Exception:
        pass


def _ensure_on_schedule_view(driver):
    try:
        if len(driver.window_handles) > 1:
            main_win = driver.window_handles[0]
            for handle in driver.window_handles[1:]:
                try:
                    driver.switch_to.window(handle)
                    driver.close()
                except Exception:
                    pass
            driver.switch_to.window(main_win)
            log.info("  →  Closed extra tabs, returned to schedule")
    except Exception as e:
        log.debug(f"Tab cleanup error: {e}")

    _close_popup_if_open(driver)
    time.sleep(0.5)


def _get_tooltip_for_appointment(driver, appt_element) -> str:
    """
    Hovers over the appointment to trigger its tooltip, then reads the
    tooltip text and returns the part BEFORE the time.

    WebPT tooltip HTML structure (confirmed):
      <div class="tooltiptext">
          <div>Teletherapy 6:30pm</div>     ← first child div = type + time
          <div>...patient name...</div>
          <div>(718)753-7116</div>
          ...
      </div>

    Strategy:
      1. Hover over the appointment
      2. Find the visible .tooltiptext element
      3. Read its FIRST child <div> only (e.g. "Teletherapy 6:30pm")
      4. Strip the time → returns "Teletherapy"

    Returns empty string if tooltip can't be read.
    """
    try:
        from selenium.webdriver.common.action_chains import ActionChains

        # ── Step 0: Clear any leftover tooltip from a previous appt ──
        try:
            body = driver.find_element(By.TAG_NAME, "body")
            ActionChains(driver).move_to_element_with_offset(body, 0, 0).perform()
            time.sleep(0.5)
        except Exception:
            pass

        # ── Step 1: Scroll appointment into view ────────────────────
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", appt_element)
            time.sleep(0.3)
        except Exception:
            pass

        # Get the appointment's position so we can match the NEAREST tooltip
        appt_box = driver.execute_script("""
            var r = arguments[0].getBoundingClientRect();
            return {left: r.left, right: r.right, top: r.top, bottom: r.bottom,
                    cx: (r.left+r.right)/2, cy: (r.top+r.bottom)/2};
        """, appt_element)
        appt_cx = appt_box.get('cx', 0) if appt_box else 0
        appt_cy = appt_box.get('cy', 0) if appt_box else 0

        # ── Step 2: Hover + read tooltip, with retries for "Loading" ─
        first_line = ""
        max_tries = 5
        for attempt in range(1, max_tries + 1):
            # Re-hover each attempt (keeps the tooltip alive / refreshes it)
            try:
                ActionChains(driver).move_to_element(appt_element).perform()
            except Exception:
                pass
            # Wait progressively longer; the tooltip loads its content async
            time.sleep(1.0 + attempt * 0.4)

            # Find the tooltip NEAREST to the appointment (by center distance).
            # This avoids reading a different appointment's tooltip.
            first_line = driver.execute_script("""
                var apptCx = arguments[0];
                var apptCy = arguments[1];
                var tips = document.querySelectorAll('.tooltiptext');
                var best = '';
                var bestDist = Infinity;
                for (var i = 0; i < tips.length; i++) {
                    var tip = tips[i];
                    var r = tip.getBoundingClientRect();
                    if (r.width < 10 || r.height < 5) continue;
                    var style = window.getComputedStyle(tip);
                    if (style.display === 'none' || style.visibility === 'hidden') continue;
                    if (parseFloat(style.opacity || '1') < 0.1) continue;

                    // Tooltip's first child div text
                    var txt = '';
                    var firstDiv = tip.querySelector('div');
                    if (firstDiv) {
                        txt = (firstDiv.innerText || firstDiv.textContent || '').trim();
                    }
                    if (!txt) {
                        var fullTxt = (tip.innerText || tip.textContent || '').trim();
                        if (fullTxt) txt = fullTxt.split('\\n')[0].trim();
                    }
                    if (!txt) continue;

                    // Distance from this tooltip's center to the appointment center
                    var tcx = (r.left + r.right) / 2;
                    var tcy = (r.top + r.bottom) / 2;
                    var dist = Math.sqrt(Math.pow(tcx-apptCx,2) + Math.pow(tcy-apptCy,2));

                    if (dist < bestDist) {
                        bestDist = dist;
                        best = txt;
                    }
                }
                return best;
            """, appt_cx, appt_cy)

            log.info(f"    [Tooltip] Attempt {attempt}: '{first_line}'")

            # If still loading or empty, retry
            if not first_line:
                continue
            if re.search(r'loading', first_line, re.IGNORECASE):
                log.info("    [Tooltip] Content still loading — retrying ...")
                first_line = ""
                continue

            # Got real content
            break

        if not first_line:
            log.info("    [Tooltip] No usable tooltip after retries")
            return ""

        log.info(f"    [Tooltip] First line: '{first_line}'")

        # ── Step 3: Strip the time portion off the end ──────────────
        # Time patterns: "6:30pm", "9:30 AM", "10:00pm", etc.
        cleaned = re.sub(
            r'\s*\d{1,2}:\d{2}\s*(am|pm|AM|PM)?\s*.*$',
            '',
            first_line
        ).strip()
        cleaned = re.sub(r'[\s\-:]+$', '', cleaned).strip()

        # Safety: never return "Loading..." as an appointment type
        if cleaned and re.search(r'loading', cleaned, re.IGNORECASE):
            log.info("    [Tooltip] Cleaned value is 'Loading' — discarding")
            return ""

        if cleaned:
            log.info(f"    [Tooltip] ✓ Appointment Type: '{cleaned}'")
            return cleaned
        else:
            log.info(f"    [Tooltip] After cleaning: empty (original was '{first_line}')")
            return ""

    except Exception as e:
        log.debug(f"_get_tooltip_for_appointment error: {e}")
        return ""


def _get_therapist_for_appointment(driver, appt_element) -> str:
    """
    Returns the name of the Treating Therapist column that the given
    appointment is positioned under in the WebPT schedule.

    HTML structure in WebPT (confirmed):
      <div class="... calendar-header-slider ...">
          <span>ADELLA CHENG (5)</span>
      </div>

    Each therapist column header has the CSS class "calendar-header-slider".
    We find all of them, take the one whose X-range contains the
    appointment's X-center, and return its name (without the "(N)" count).
    """
    try:
        # ── Step 0: Scroll appointment into view ────────────────────
        try:
            driver.execute_script("""
                arguments[0].scrollIntoView({block: 'center', inline: 'center'});
            """, appt_element)
            time.sleep(0.3)
        except Exception:
            pass

        # ── Step 1: Get the appointment's bounding box ──────────────
        appt_rect = driver.execute_script("""
            var r = arguments[0].getBoundingClientRect();
            return {left: r.left, right: r.right, top: r.top, center_x: (r.left + r.right) / 2};
        """, appt_element)

        if not appt_rect:
            log.info("    [Therapist Lookup] Could not read appointment bounding box")
            return ""

        appt_left     = appt_rect.get('left', 0)
        appt_right    = appt_rect.get('right', 0)
        appt_center_x = appt_rect.get('center_x', 0)
        appt_top      = appt_rect.get('top', 0)
        log.info(f"    [Therapist Lookup] Appointment X: {appt_left:.1f}-{appt_right:.1f} (center {appt_center_x:.1f}), Top: {appt_top:.1f}")

        # ── Step 2: Find all "calendar-header-slider" elements ──────
        # This is the exact class WebPT uses for therapist column headers.
        headers = driver.execute_script("""
            var pattern = /^(.+?)\\s*\\((\\d+)\\)\\s*$/;
            var results = [];
            var nodes = document.querySelectorAll('.calendar-header-slider');
            for (var i = 0; i < nodes.length; i++) {
                var el = nodes[i];
                var txt = (el.innerText || el.textContent || '').trim();
                if (!txt) continue;
                if (txt.indexOf('\\n') >= 0) {
                    // Multi-line: take first line
                    txt = txt.split('\\n')[0].trim();
                }
                var r = el.getBoundingClientRect();
                if (r.width < 5) continue;

                var m = txt.match(pattern);
                var name = m ? m[1].trim() : txt;
                var count = m ? m[2] : '';

                results.push({
                    name: name,
                    count: count,
                    left: r.left,
                    right: r.right,
                    top: r.top,
                    width: r.width,
                    fullText: txt
                });
            }
            return results;
        """)

        if not headers:
            log.info("    [Therapist Lookup] No '.calendar-header-slider' elements found")
            return ""

        log.info(f"    [Therapist Lookup] Found {len(headers)} therapist column header(s):")
        for h in headers[:20]:
            log.info(f"      • '{h['fullText']}' [X: {h['left']:.0f}-{h['right']:.0f}, Top: {h['top']:.0f}]")

        # ── Step 3: Match by the appointment's LEFT EDGE ────────────
        # In WebPT, an appointment block ALWAYS starts at the left edge
        # of its own column. If an appointment is wide (spanning visually
        # into the next column), its therapist is still the one whose
        # column the LEFT edge falls into — NOT the one with more overlap.
        #
        # So: find the header whose X-range contains the appointment's
        # LEFT edge (with a small tolerance for sub-pixel rounding).
        TOL = 5  # pixels of tolerance

        left_matches = []
        for h in headers:
            if (h['left'] - TOL) <= appt_left <= (h['right'] + TOL):
                left_matches.append(h)

        if left_matches:
            # If the left edge lands in more than one header's range
            # (overlapping ranges at the boundary), pick the one whose
            # left edge is closest to (but not greater than) appt_left —
            # i.e. the column that actually starts at/just-before the appt.
            left_matches.sort(key=lambda h: abs(h['left'] - appt_left))
            chosen = left_matches[0]
            log.info(f"    [Therapist Lookup] ✓ MATCH (by left edge {appt_left:.0f}): '{chosen['name']}' [col X: {chosen['left']:.0f}-{chosen['right']:.0f}]")
            return chosen['name']

        # ── Step 4: Fallback — header whose left edge is closest below appt_left ──
        # Find the column whose left edge is the largest value that is
        # still <= appt_left (i.e. the column the appointment starts in).
        below = [h for h in headers if h['left'] <= appt_left + TOL]
        if below:
            below.sort(key=lambda h: h['left'])
            chosen = below[-1]  # largest left <= appt_left
            log.info(f"    [Therapist Lookup] ✓ MATCH (left edge nearest column start): '{chosen['name']}'")
            return chosen['name']

        # ── Step 5: Fallback — maximum overlap ──────────────────────
        def _overlap(a_left, a_right, b_left, b_right):
            lo = max(a_left, b_left)
            hi = min(a_right, b_right)
            return max(0.0, hi - lo)

        best_header   = None
        best_overlap  = 0.0
        for h in headers:
            ov = _overlap(appt_left, appt_right, h['left'], h['right'])
            if ov > best_overlap:
                best_overlap = ov
                best_header  = h

        if best_header and best_overlap > 0:
            log.info(f"    [Therapist Lookup] ✓ MATCH (max overlap {best_overlap:.0f}px): '{best_header['name']}'")
            return best_header['name']

        # ── Step 6: Fallback — closest header by X distance ─────────
        log.info("    [Therapist Lookup] No edge/overlap match — using closest by X")
        closest = None
        closest_dist = float('inf')
        for h in headers:
            h_center = (h['left'] + h['right']) / 2
            dist = abs(h_center - appt_center_x)
            if dist < closest_dist:
                closest_dist = dist
                closest = h

        if closest:
            log.info(f"    [Therapist Lookup] ✓ Closest: '{closest['name']}' (X dist: {closest_dist:.0f})")
            return closest['name']

    except Exception as e:
        log.debug(f"_get_therapist_for_appointment error: {e}")
        import traceback as _tb
        for line in _tb.format_exc().splitlines():
            log.debug(f"  {line}")

    return ""


def find_patient_in_schedule(driver, full_name: str, active_filter: str = ""):
    """
    Returns a list of (status, appt_date, therapist) tuples — one per appointment.

    active_filter: if a status filter ("No Show" or "Cancelled") is currently
                   active, pass it here. When the patient is found while a
                   filter is active, that filter's name BECOMES the status —
                   even if no icon is visible. This is critical for No Show
                   patients whose '?' icon is sometimes rendered as an image
                   that doesn't appear in the visible text.

    Name matching is robust to middle name differences:
      Excel:    "FAGAINS, MICHAEL KEITH"   first=MICHAEL KEITH, last=FAGAINS
      Schedule: "FAGAINS, MICHAEL K. ..."  ← only initial!
      We match using the FIRST WORD of first_name (MICHAEL) + last_name,
      and strip dots so "K." becomes "K".
    """
    first_name  = get_first_name(full_name)
    last_name   = _get_last_name(full_name)
    first_name  = re.sub(r'[,،‚]', '', first_name).strip()
    last_name   = re.sub(r'[,،‚]', '', last_name).strip()

    # Use only the FIRST WORD of the first name, in case Excel has
    # the middle name written out but the schedule shows only an initial.
    # e.g. Excel "MICHAEL KEITH" vs schedule "MICHAEL K." → match "MICHAEL"
    first_name_primary = first_name.split()[0] if first_name.split() else first_name

    log.info(f"  →  Searching schedule: First='{first_name_primary}' Last='{last_name}'  (full first name: '{first_name}')")
    time.sleep(2)

    def _norm(s: str) -> str:
        if not s:
            return ""
        s = str(s)
        s = re.sub(r'[،‚]', ',', s)
        s = re.sub(r"['`’]", '', s)
        # Strip dots (so "K." matches "K")
        s = s.replace('.', '')
        s = re.sub(r'\s+', ' ', s).strip()
        return s.upper()

    first_norm = _norm(first_name_primary)
    last_norm  = _norm(last_name)

    try:
        content_div = driver.find_element(By.ID, "contentWindow")
        page_text   = content_div.get_attribute("innerText") or ""
    except Exception:
        page_text = driver.execute_script("return document.body.innerText || '';")

    page_text_norm = _norm(page_text)

    matched_in_page = False
    if first_norm and first_norm in page_text_norm and last_norm and last_norm in page_text_norm:
        matched_in_page = True
    else:
        first_alt = first_norm.replace('-', ' ')
        last_alt  = last_norm.replace('-', ' ')
        page_alt  = page_text_norm.replace('-', ' ')
        if first_alt in page_alt and last_alt in page_alt:
            matched_in_page = True
            log.info(f"  →  Matched via hyphen-normalization")

    if not matched_in_page:
        warn(f"'{full_name}' not found in schedule page → Not Found")
        log.info(f"  →  Searched for First='{first_norm}' Last='{last_norm}'")
        return [("Not Found", "", "", "")]

    ok(f"'{full_name}' confirmed in schedule page")

    appointment_specs = []

    try:
        title_divs = driver.find_elements(By.CSS_SELECTOR, "div.title")
        log.info(f"  →  Scanning {len(title_divs)} div.title elements ...")

        for idx, el in enumerate(title_divs):
            try:
                txt = el.text.strip()
                if not txt:
                    continue
                txt_norm = _norm(txt)

                matched_here = (first_norm in txt_norm and last_norm in txt_norm)

                if not matched_here:
                    txt_alt   = txt_norm.replace('-', ' ')
                    first_alt = first_norm.replace('-', ' ')
                    last_alt  = last_norm.replace('-', ' ')
                    matched_here = (first_alt in txt_alt and last_alt in txt_alt)

                if not matched_here:
                    continue

                log.info(f"  →  Found match #{len(appointment_specs)+1}: '{txt[:80]}'")

                inner_html = ""
                try:
                    inner_html = el.get_attribute("innerHTML") or ""
                except Exception:
                    pass

                icon_status = None

                for marker in ["√", "✓", "✔", "checked in", "check-in"]:
                    if marker in txt or marker.lower() in inner_html.lower():
                        icon_status = "Checked In"
                        log.info(f"  →  Icon '{marker}' detected → Checked In")
                        break

                if not icon_status:
                    for marker in ["«", "»", "✗", "checked out", "check-out"]:
                        if marker in txt or marker.lower() in inner_html.lower():
                            icon_status = "Checked Out"
                            log.info(f"  →  Icon '{marker}' detected → Checked Out")
                            break

                if not icon_status:
                    for marker in ["Φ", "φ", "ϕ", "cancelled", "canceled"]:
                        if marker in txt or marker.lower() in inner_html.lower():
                            icon_status = "Cancelled"
                            log.info(f"  →  Icon '{marker}' detected → Cancelled")
                            break

                if not icon_status:
                    # ── No Show detection ──────────────────────────────
                    # The '?' icon appears next to the patient name, often
                    # ATTACHED to it without a space, e.g. "?Floresluna, Margarita"
                    # We need to detect '?' that is NOT inside a question
                    # (i.e. not at end of sentence). The icon is always:
                    #   - At the START of the text/name, OR
                    #   - Right after a separator like "(", ":", or whitespace
                    #
                    # Strong signals (any of these = No Show):
                    #   - txt starts with '?'
                    #   - txt contains '? ' or ' ?'
                    #   - inner_html contains '?' near the start
                    #   - "no show" / "no-show" anywhere in text
                    is_no_show = False
                    txt_stripped = txt.strip()

                    if txt_stripped.startswith('?'):
                        is_no_show = True
                        log.info(f"  →  '?' at start of text → No Show")
                    elif re.search(r'(^|[\s\(])\?', txt):
                        is_no_show = True
                        log.info(f"  →  '?' icon detected → No Show")
                    elif "no show" in txt.lower() or "no-show" in txt.lower():
                        is_no_show = True
                        log.info(f"  →  'no show' text → No Show")
                    else:
                        # Check inner HTML for an early '?' that's not part
                        # of a longer string (DOB, ID, etc.)
                        # Look at first 200 chars of inner_html for ?
                        early_html = inner_html[:200] if inner_html else ""
                        if '?' in early_html:
                            # Make sure it's not just a URL parameter or attribute
                            # Look for ? as actual content (not href="?...")
                            # Strategy: find '?' positions and check if they're
                            # inside text content (between > and <)
                            content_q = re.search(r'>([^<]*?\?[^<]*?)<', early_html)
                            if content_q:
                                is_no_show = True
                                log.info(f"  →  '?' in inner HTML content → No Show")

                    if is_no_show:
                        icon_status = "No Show"

                # ── ACTIVE FILTER FALLBACK ─────────────────────────────
                # If we're currently viewing a "No Show" or "Cancelled"
                # filter and the patient appears in that filtered list,
                # then their status MUST be that filter's name. The icon
                # may not be visible (sometimes WebPT renders the '?' or
                # 'Φ' as an image that doesn't appear in the text).
                # The fact that they're listed under this filter is itself
                # the strongest evidence of their status.
                if not icon_status and active_filter:
                    if active_filter.lower() in ("no show", "no-show", "noshow"):
                        icon_status = "No Show"
                        log.info(f"  →  Active filter is 'No Show' → Status: No Show (icon not visible)")
                    elif active_filter.lower() in ("cancelled", "canceled"):
                        icon_status = "Cancelled"
                        log.info(f"  →  Active filter is 'Cancelled' → Status: Cancelled (icon not visible)")

                if not icon_status:
                    icon_status = "Other"
                    log.info(f"  →  No icon detected → Status will be 'Other'")

                appointment_specs.append({
                    "icon_status": icon_status,
                    "title_text":  txt,
                    "dom_index":   idx,
                })
            except Exception:
                continue
    except Exception as e:
        log.debug(f"div.title scan error: {e}")

    if not appointment_specs:
        warn(f"Could not find appointment div for '{full_name}' → Not Found")
        return [("Not Found", "", "", "")]

    log.info(f"  →  Total appointments to process for '{full_name}': {len(appointment_specs)}")

    results = []
    for spec_idx, spec in enumerate(appointment_specs):
        log.info(f"  ━━━ Processing appointment {spec_idx+1}/{len(appointment_specs)} ━━━")

        target_text = spec["title_text"]
        matched_el = None
        try:
            current_titles = driver.find_elements(By.CSS_SELECTOR, "div.title")
            matches_so_far = 0
            for el in current_titles:
                try:
                    if el.text.strip() == target_text:
                        if matches_so_far == spec_idx or len(appointment_specs) == 1:
                            matched_el = el
                            break
                        matches_so_far += 1
                except Exception:
                    continue
        except Exception as e:
            log.debug(f"Re-finding element error: {e}")

        if matched_el is None:
            warn(f"Could not re-find appointment {spec_idx+1} — skipping")
            continue

        result = _process_single_appointment(
            driver, full_name, matched_el, spec["icon_status"]
        )
        results.append(result)

        _ensure_on_schedule_view(driver)
        time.sleep(1)

    _ensure_on_schedule_view(driver)

    if not results:
        return [("Not Found", "", "", "")]

    return results


def _click_day_view(driver):
    log.info("  →  Clicking 'Day' view button ...")
    for xpath in [
        "//input[@value='Day']",
        "//button[normalize-space(text())='Day']",
        "//a[normalize-space(text())='Day']",
        "//td[normalize-space(text())='Day']",
        "//*[normalize-space(text())='Day']",
    ]:
        try:
            els = driver.find_elements(By.XPATH, xpath)
            for el in els:
                if el.is_displayed():
                    driver.execute_script("arguments[0].click();", el)
                    time.sleep(2)
                    ok("'Day' view selected")
                    return
        except Exception:
            continue
    warn("'Day' button not found — continuing with current view")


def _is_filter_active(driver, status: str) -> bool:
    """
    Returns True if the given filter button ('Checked Out', 'No Show',
    'Cancelled') currently appears to be active/pressed.

    WebPT styles active filters differently (usually pressed state, or
    'x-btn-pressed' class, or different background color). We check
    common Ext JS patterns.
    """
    try:
        # Find the button/element with this filter's text
        for xpath in [
            f"//input[@value='{status}']",
            f"//button[normalize-space(text())='{status}']",
            f"//a[normalize-space(text())='{status}']",
            f"//td[normalize-space(text())='{status}']",
            f"//*[normalize-space(text())='{status}']",
        ]:
            try:
                els = driver.find_elements(By.XPATH, xpath)
                for el in els:
                    if not el.is_displayed():
                        continue

                    # Check the element itself and its ancestors for
                    # pressed/active class indicators.
                    for target in [el] + el.find_elements(By.XPATH, "./ancestor::*[position()<=3]"):
                        try:
                            cls = (target.get_attribute("class") or "").lower()
                            if any(k in cls for k in [
                                "pressed", "active", "selected",
                                "x-btn-pressed", "x-pressed"
                            ]):
                                log.info(f"  →  Filter '{status}' is ACTIVE (class='{cls[:60]}')")
                                return True
                        except Exception:
                            continue

                    # If we found the button but no active class → not active
                    return False
            except Exception:
                continue
    except Exception:
        pass
    # Default: assume not active (safer to click and toggle if needed)
    return False


def _click_status_filter(driver, status: str):
    log.info(f"  →  Clicking '{status}' filter button ...")
    for xpath in [
        f"//input[@value='{status}']",
        f"//button[normalize-space(text())='{status}']",
        f"//a[normalize-space(text())='{status}']",
        f"//td[normalize-space(text())='{status}']",
        f"//*[normalize-space(text())='{status}']",
    ]:
        try:
            els = driver.find_elements(By.XPATH, xpath)
            for el in els:
                if el.is_displayed():
                    driver.execute_script("arguments[0].click();", el)
                    time.sleep(2)
                    ok(f"'{status}' filter selected")
                    return
        except Exception:
            continue
    warn(f"'{status}' button not found — continuing with current view")


# =====================================================================
# FULL SCHEDULE CHECK FLOW
# =====================================================================

def _is_on_schedule_view(driver) -> bool:
    """
    Returns True if the browser is currently on the View Schedule page.
    Detection is based on the presence of the ClinicChange dropdown element,
    which only exists on the schedule page.
    """
    try:
        el = driver.find_element(By.ID, "ClinicChange")
        return el is not None
    except Exception:
        return False


def check_schedule_status(driver, wait, full_name: str, target_date: str, clinic_name: str):
    """
    Returns a list of (status, appt_date) tuples — one per appointment.

    Flow per patient:
      1. Verify we're on View Schedule (open it once if not)
      2. Select Clinic (skipped if same as last)
      3. Navigate to Date in the calendar
      4. Click Day view
      5. Search for the patient
      6. Process found appointments (close popup at end)

    The script stays on View Schedule for the whole run — it does NOT
    return to the dashboard between patients.
    """
    step(f"Schedule Check | '{full_name}' | {target_date} | {clinic_name}")

    # ── Step 1: Make sure we're on View Schedule ──────────────────
    if _is_on_schedule_view(driver):
        ok("Already on View Schedule — staying here")
    else:
        warn("Not on View Schedule — re-opening it")
        opened = _open_view_schedule(driver, wait)
        if not opened:
            warn("Could not open View Schedule — returning Other")
            return [("Other", "", "", "")]

    # ── Step 2: Select clinic (skipped internally if same as last) ──
    select_clinic_in_schedule(driver, clinic_name)

    # ── Step 3: Navigate to the target date in the calendar ─────────
    navigate_schedule_to_date(driver, target_date)

    # ── Step 4: Make sure we're in Day view ─────────────────────────
    _click_day_view(driver)

    # ── Verify date actually changed ────────────────────────────────
    try:
        time.sleep(1)
        page_body = driver.find_element(By.TAG_NAME, "body").text[:2000]
        try:
            target_d = pd.to_datetime(target_date)
            target_day_str = str(target_d.day)
            target_month_abbr = target_d.strftime("%b")
            target_year_str = str(target_d.year)
            shows_target = (target_month_abbr in page_body and target_day_str in page_body and target_year_str in page_body)
            if not shows_target:
                log.info(f"  →  Day view doesn't seem to show target date — navigating again")
                navigate_schedule_to_date(driver, target_date)
                _click_day_view(driver)
                time.sleep(1)
        except Exception as e:
            log.debug(f"  →  Date verification skipped: {e}")
    except Exception as e:
        log.debug(f"  →  Day view verification error: {e}")

    # ── Step 5: Activate filters (only if not already active) ───────
    # Per user request: keep No Show + Cancelled active across patients
    # in the SAME clinic. When the clinic changes, _extra_filters_active
    # is reset by select_clinic_in_schedule, so we re-activate them here.
    global _extra_filters_active

    if not _extra_filters_active:
        log.info("  →  Activating filters for the first time in this clinic ...")
        if not _is_filter_active(driver, "Checked Out"):
            log.info("  →  Activating 'Checked Out' filter")
            _click_status_filter(driver, "Checked Out")
            time.sleep(0.8)

        log.info("  →  Activating 'No Show' filter")
        _click_status_filter(driver, "No Show")
        time.sleep(0.8)

        log.info("  →  Activating 'Cancelled' filter")
        _click_status_filter(driver, "Cancelled")
        time.sleep(1.2)

        _extra_filters_active = True
    else:
        log.info("  →  Filters (Checked Out + No Show + Cancelled) already active — reusing")

    # ── Step 6: Search for the patient (all 3 statuses visible) ─────
    # The icon next to the name determines the status.
    results = find_patient_in_schedule(driver, full_name, active_filter="")

    # ── Step 7: Keep filters active for the next patient ────────────
    # We do NOT toggle off No Show / Cancelled. They stay on so the next
    # patient (same clinic) can be searched immediately. They'll be reset
    # automatically when the clinic changes.

    return results


# =====================================================================
# PROCESS ONE PATIENT
# =====================================================================

def process_patient(driver, wait, patient_id: str, clinic_name: str,
                    target_date: str, full_name: str) -> list:
    step(f"Processing | ID:{patient_id} | {clinic_name} | {target_date} | {full_name}")

    appointments = check_schedule_status(driver, wait, full_name, target_date, clinic_name)

    out_records = []
    for idx, appt in enumerate(appointments):
        # Each appt is (status, appt_date, therapist, appt_type)
        status, appt_date, therapist, appt_type = appt
        if len(appointments) > 1:
            log.info(f"  →  Appointment {idx+1}/{len(appointments)} → Status: {status}")

        ok(f"STATUS for '{full_name}' (appt {idx+1}): [{status}] | Date: [{appt_date}] | Therapist: [{therapist}] | Type: [{appt_type}]")

        out_records.append({
            "EMR_ID":    patient_id,
            "Name":      full_name,
            "Clinic":    clinic_name,
            "Date":      target_date,
            "Status":    status,
            "Appt_Date": appt_date,
            "Therapist": therapist,
            "ApptType":  appt_type,
        })

    return out_records


# =====================================================================
# SAVE TO EXCEL — 7 COLUMNS
# =====================================================================

def save_styled_excel(results: list, output_path: str):
    """
    Header columns (9):
      A  Patient ID
      B  Patient Name
      C  Clinic Name
      D  Appointment Date
      E  Status
      F  Status Update Date
      G  Status Update Time
      H  Treating Therapist
      I  Appointment Type
    """
    headers = [
        "Patient ID",
        "Patient Name",
        "Clinic Name",
        "Appointment Date",
        "Status",
        "Status Update Date",
        "Status Update Time",
        "Treating Therapist",
        "Appointment Type",
    ]

    black_fill   = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
    white_bold   = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_side    = Side(style="thin", color="000000")
    thin_border  = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    data_font    = Font(name="Calibri", size=11)

    file_exists = os.path.exists(output_path)
    if file_exists:
        # ── Delete old file and start fresh ────────────────────────
        # User wants the results file to ALWAYS contain only the
        # current run's data — no old rows kept.
        try:
            os.remove(output_path)
            log.info(f"  →  Deleted old file: {output_path}")
        except PermissionError:
            warn(f"Could not delete old file (open in Excel?): {output_path}")
            # Will save with timestamp suffix at the end
        except Exception as e:
            warn(f"Could not delete old file: {e}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    next_row = 1
    need_headers = True
    log.info("  →  Creating fresh output file")

    if need_headers:
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.fill      = black_fill
            cell.font      = white_bold
            cell.alignment = center_align
            cell.border    = thin_border
        next_row = 2

    rows_added = 0
    for record in results:
        appt_full = record.get("Appt_Date", "")
        if appt_full and " " in appt_full:
            parts = appt_full.split(" ", 1)
            try:
                appt_date_part = pd.to_datetime(parts[0]).strftime("%m/%d/%Y")
            except Exception:
                appt_date_part = parts[0]
            appt_time_part = parts[1] if len(parts) > 1 else ""
        elif appt_full:
            try:
                appt_date_part = pd.to_datetime(appt_full).strftime("%m/%d/%Y")
            except Exception:
                appt_date_part = appt_full
            appt_time_part = ""
        else:
            appt_date_part = ""
            appt_time_part = ""

        raw_date = record.get("Date", "")
        try:
            appt_date_fmt = pd.to_datetime(raw_date).strftime("%m/%d/%Y")
        except Exception:
            appt_date_fmt = raw_date

        emr_id_raw = record.get("EMR_ID", "")
        try:
            emr_id_val = int(emr_id_raw) if str(emr_id_raw).strip() else ""
        except (ValueError, TypeError):
            emr_id_val = emr_id_raw

        values = [
            emr_id_val,                       # A: Patient ID
            record.get("Name",      ""),      # B: Patient Name
            record.get("Clinic",    ""),      # C: Clinic Name
            appt_date_fmt,                    # D: Appointment Date
            record.get("Status",    ""),      # E: Status
            appt_date_part,                   # F: Status Update Date
            appt_time_part,                   # G: Status Update Time
            record.get("Therapist", ""),      # H: Treating Therapist
            record.get("ApptType",  ""),      # I: Appointment Type
        ]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=next_row, column=col_idx, value=value)
            cell.font      = data_font
            cell.alignment = center_align
            cell.border    = thin_border
        next_row += 1
        rows_added += 1

    if need_headers:
        col_widths = {
            "A": 12, "B": 25, "C": 18, "D": 16,
            "E": 15, "F": 18, "G": 18, "H": 22, "I": 20,
        }
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

    for row_num in range(next_row - rows_added, next_row):
        ws.row_dimensions[row_num].height = 18

    try:
        wb.save(output_path)
        ok(f"Saved {rows_added} row(s) to: {output_path}")
    except PermissionError:
        ts_path = output_path.replace('.xlsx', f'_{int(time.time())}.xlsx')
        wb.save(ts_path)
        warn(f"File was locked. Saved as: {ts_path}")


# =====================================================================
# MAIN
# =====================================================================

CLINIC_SLUG = re.sub(r"[^A-Za-z0-9_-]+", "_", CLINIC_NAME).strip("_")
CLINIC_COUNTS_OUTPUT_PATH = os.path.join(
    LOCAL_WORK_DIR, f"clinic_hourly_counts_{CLINIC_SLUG}.xlsx"
)


def _read_selected_calendar_date(driver) -> str:
    selection = driver.execute_script(r"""
        const calendar = document.getElementById('sideMenuCalendar');
        const components = window.Ext && Ext.ComponentMgr && Ext.ComponentMgr.all;
        if (calendar && components && components.items) {
            for (const component of components.items) {
                const node = component.el && component.el.dom;
                if (!node || !calendar.contains(node) || typeof component.getValue !== 'function') continue;
                const value = component.getValue();
                if (value instanceof Date && !isNaN(value.getTime())) {
                    return {month: ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][value.getMonth()],
                            year: value.getFullYear(), day: value.getDate()};
                }
            }
        }
        const visible = el => el.getClientRects().length &&
            getComputedStyle(el).visibility !== 'hidden';
        const monthPattern = /\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})\b/i;
        const tables = Array.from(document.querySelectorAll('table'))
            .filter(el => visible(el) && monthPattern.test(el.innerText) &&
                /\bMo\b/.test(el.innerText) && /\bSu\b/.test(el.innerText))
            .sort((a, b) => a.querySelectorAll('*').length - b.querySelectorAll('*').length);
        for (const table of tables) {
            const header = table.innerText.match(monthPattern);
            const days = new Set();
            for (const el of table.querySelectorAll('td, a, span, button')) {
                const text = (el.innerText || '').trim();
                if (!visible(el) || !/^\d{1,2}$/.test(text)) continue;
                const cell = el.closest('td');
                const classes = (el.className + ' ' + (cell ? cell.className : '')).toLowerCase();
                if (/prevday|nextday|outside|other-month/.test(classes)) continue;
                const marked = /(^|[\s_-])(selected|active|chosen)([\s_-]|$)/.test(classes) ||
                    el.getAttribute('aria-selected') === 'true' ||
                    (cell && cell.getAttribute('aria-selected') === 'true');
                // Some legacy calendars mark the selected day only in yellow.
                const yellow = [el, cell].filter(Boolean).some(node =>
                    getComputedStyle(node).backgroundColor === 'rgb(255, 255, 0)');
                if (marked || yellow) days.add(Number(text));
            }
            if (days.size === 1) return {month: header[1], year: header[2], day: [...days][0]};
        }
        return null;
    """)
    if not selection:
        return ""
    try:
        return pd.to_datetime(
            f"{selection['month']} {selection['day']} {selection['year']}",
            format="%b %d %Y",
        ).strftime("%Y-%m-%d")
    except (KeyError, ValueError, TypeError):
        return ""


OT_SUFFIX = "OT"


def _letters(text):
    return re.sub(r"[^a-z]", "", str(text or "").casefold())


def _ot_reason(entry, clinic_name):
    """Return the reason this individual schedule column is OT, or an empty string."""
    title = str(entry.get("title") or "").strip()
    label_name = str(entry.get("type") or "").strip()
    if re.search(r"\bOT\b", f"{title} {entry.get('header', '')}", re.IGNORECASE):
        return "OT text"
    if not CLINIC_HAS_OT:
        return ""
    if title:
        return ""
    clinic, plain = _letters(clinic_name), _letters(label_name)
    if not (len(plain) >= 4 and (plain == clinic or clinic.startswith(plain)
                               or plain.startswith(clinic))):
        return ""
    first = next((ch for ch in label_name if ch.isalpha()), "")
    return "single clinic-name column" if first.isupper() else ""


def _clinic_variants():
    return [CLINIC_NAME, f"{CLINIC_NAME} {OT_SUFFIX}"]


def _clinic_sort_key(name):
    name = str(name).strip()
    is_ot = name.endswith(f" {OT_SUFFIX}")
    base = name[:-len(OT_SUFFIX) - 1] if is_ot else name
    return base.casefold(), int(is_ot)


def save_clinic_counts_excel(records):
    """Writes two sheets:
       Hourly Counts : Clinic | date_of_visit | Hour | Appointment Type | Count
       Appointments  : one row per appointment (for checking the numbers)"""
    details = pd.DataFrame(records)
    for column in ["Clinic", "date_of_visit", "Hour", "Appointment Type",
                   "Therapist", "Patient", "Tooltip", "Status"]:
        if column not in details:
            details[column] = ""
    if "Hour24" not in details:
        details["Hour24"] = 9999
    details = details.fillna({"Hour": "", "Appointment Type": "", "Therapist": "",
                              "Patient": "", "Tooltip": "", "Status": "",
                              "date_of_visit": "", "Hour24": 9999})

    keys = ["Clinic", "date_of_visit", "Hour24", "Hour", "Appointment Type"]
    ok_rows = details[details["Status"] == "OK"]
    hourly = ok_rows.groupby(keys, sort=False).size().reset_index(name="Count")

    # A day with no appointments still gets one row, so the sheet shows it
    # (and older numbers for that day are replaced).
    empty = details[details["Status"] == "EMPTY"][["Clinic", "date_of_visit"]].drop_duplicates()
    empty = empty.assign(**{"Hour24": -1, "Hour": "", "Appointment Type": "No appointments", "Count": 0})

    hourly = pd.concat([hourly, empty], ignore_index=True)
    order = {name: i for i, name in enumerate(_clinic_variants())}
    hourly["_order"] = hourly["Clinic"].map(order).fillna(len(order))
    hourly = hourly.sort_values(["_order", "date_of_visit", "Hour24", "Appointment Type"])
    hourly = hourly[["Clinic", "date_of_visit", "Hour", "Appointment Type", "Count"]]
    hourly = hourly.rename(columns={"Hour": "Time"})

    appointments = details[details["Status"] != "EMPTY"][
        ["Clinic", "date_of_visit", "Therapist", "Patient", "Hour",
         "Appointment Type", "Tooltip", "Status"]
    ].rename(columns={"Hour": "Time"})
    Path(CLINIC_COUNTS_OUTPUT_PATH).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(CLINIC_COUNTS_OUTPUT_PATH, engine="openpyxl") as writer:
        hourly.to_excel(writer, sheet_name="Hourly Counts", index=False)
        appointments.to_excel(writer, sheet_name="Appointments", index=False)
    _s3_upload(CLINIC_COUNTS_OUTPUT_PATH, f"exports/{Path(CLINIC_COUNTS_OUTPUT_PATH).name}")


def _schedule_heading_matches_date(heading, target_date):
    match = re.search(r"\b(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?\b", heading or "")
    if not match:
        return False
    target = pd.Timestamp(target_date)
    month, day = int(match[1]), int(match[2])
    year = int(match[3]) if match[3] else target.year
    if year < 100:
        year += 2000
    return (year, month, day) == (target.year, target.month, target.day)


def _save_schedule_failure(driver, clinic_name, target_date, attempt):
    """Logs where the browser was when reading failed. No screenshot and no
    file is saved (screenshots would hold patient information)."""
    try:
        log.info("Schedule failure | %s | %s | attempt %s | URL: %s | title: %s",
                 clinic_name, target_date, attempt, driver.current_url, driver.title)
    except Exception as exc:
        warn(f"Could not read the browser state: {type(exc).__name__}")


def _relogin(driver, wait):
    """Session ended → log in again and reopen View Schedule, so the run
    continues from the same clinic/date it stopped at."""
    global _current_clinic, _extra_filters_active
    warn("Back on the LOGIN page — logging in again and continuing")
    _current_clinic = None
    _extra_filters_active = False
    driver.switch_to.default_content()
    if _is_login_page(driver):
        _complete_login(driver)
        ok("Logged in again")
        if not _open_view_schedule(driver, wait):
            raise RuntimeError("Could not open View Schedule after re-login")
    else:
        login(driver, wait)


def _ensure_logged_in(driver, wait):
    try:
        driver.switch_to.default_content()
    except Exception:
        pass
    if _is_login_page(driver):
        _relogin(driver, wait)
        return True
    return False


def _restore_count_schedule(driver, wait):
    global _current_clinic, _extra_filters_active
    _current_clinic = None
    _extra_filters_active = False
    driver.switch_to.default_content()
    if _ensure_logged_in(driver, wait):
        wait.until(lambda d: _get_clinic_select_element(d) is not None)
        return
    if not _open_view_schedule(driver, wait):
        if not _ensure_logged_in(driver, wait):
            raise RuntimeError("Schedule recovery failed; collection stopped.")
    wait.until(lambda d: _get_clinic_select_element(d) is not None)


# ---------------------------------------------------------------------
# 1) Select "Day" view by its exact id and verify the "selected" class
# ---------------------------------------------------------------------
def _select_day_view(driver, wait):
    def day_selected(d):
        try:
            btn = d.find_element(By.ID, "app-calendar-tb-day")
            return "selected" in (btn.get_attribute("class") or "").split()
        except Exception:
            return False

    if day_selected(driver):
        ok("Day view already selected")
        return

    btn = wait.until(EC.presence_of_element_located((By.ID, "app-calendar-tb-day")))
    driver.execute_script("arguments[0].click();", btn)
    wait.until(day_selected, "Day view button did not become selected")
    wait_for_page_stable(driver, timeout=WAIT_TIMEOUT)
    ok("Day view selected")


# ---------------------------------------------------------------------
# 2a) Read schedule column headers "NAME (N)" (was nested read_pt_label)
# ---------------------------------------------------------------------
def _read_pt_label(d):
    return d.execute_script(r"""
        const visible = el => el.getClientRects().length &&
            getComputedStyle(el).visibility !== 'hidden';
        const countPattern = /^([^()\r\n]+?)\s*\(\s*(\d+)\s*\)$/;
        const labelsIn = el => (el.innerText || '').split(/\r?\n/)
            .map(text => text.trim()).filter(text => countPattern.test(text));
        const headers = Array.from(document.querySelectorAll('.calendar-header-slider'))
            .filter(visible);
        const headerNodes = [...new Set(headers.flatMap(header =>
            [header, ...header.querySelectorAll('*')]))];
        const countElements = headerNodes
            .filter(el => visible(el) && !['SCRIPT', 'STYLE', 'OPTION'].includes(el.tagName))
            .flatMap(el => labelsIn(el).filter(label =>
                !Array.from(el.children).some(child => visible(child) &&
                    labelsIn(child).includes(label)))
                .map(label => ({el, label})));
        const datePattern = /\b(?:Mon(?:day)?|Tue(?:sday)?|Wed(?:nesday)?|Thu(?:rsday)?|Fri(?:day)?|Sat(?:urday)?|Sun(?:day)?)\s+\d{1,2}\/\d{1,2}(?:\/\d{2,4})?\b/i;
        const dates = Array.from(document.querySelectorAll('div,span,th,td'))
            .filter(el => visible(el) && datePattern.test(el.innerText || '') &&
                !Array.from(el.children).some(child => datePattern.test(child.innerText || '')))
            .map(el => {
                let container = el;
                for (let parent = el.parentElement; parent; parent = parent.parentElement) {
                    const headings = (parent.innerText || '').match(
                        new RegExp(datePattern.source, 'gi')) || [];
                    if (headings.length !== 1) break;
                    container = parent;
                }
                return {text: el.innerText.match(datePattern)[0],
                        rect: container.getBoundingClientRect(),
                        headingRect: el.getBoundingClientRect()};
            });
        return countElements.map(({el, label}) => {
                const rect = el.getBoundingClientRect();
                const column = el.closest('.calendar-header-slider');
                const header = (column.innerText || '').trim();
                const title = header.split(/\r?\n/).map(text => text.trim())
                    .filter(text => text && !countPattern.test(text))
                    .join(' ');
                const center = rect.left + rect.width / 2;
                const candidates = dates.filter(date => {
                    const r = date.rect;
                    return date.headingRect.top <= rect.top &&
                        r.left <= center && r.right >= center;
                }).sort((a, b) => (a.rect.width - b.rect.width) ||
                    b.headingRect.top - a.headingRect.top);
                const date = candidates.length ? candidates[0].text : '';
                const match = label.match(countPattern);
                return {title, header, date, label, type: match[1].trim(),
                        count: Number(match[2]), x: rect.left, y: rect.top};
            }).sort((a, b) => a.y - b.y || a.x - b.x)
            .map(({x, y, ...entry}) => entry);
    """) or []


def _filter_entries_for_date(items, target_date):
    """In Day view only one day is shown. Keep headers that match the
    target date; if the page shows no date heading at all, the calendar
    check already proved the date, so accept the headers as that day."""
    matched = [i for i in items if _schedule_heading_matches_date(i["date"], target_date)]
    if matched:
        return matched
    if items and all(not i["date"] for i in items):
        label = pd.Timestamp(target_date).strftime("%a %m/%d/%Y")
        return [{**i, "date": label} for i in items]
    return []


# ---------------------------------------------------------------------
# 2a-bis) Read each appointment block: type + time (from its tooltip)
# ---------------------------------------------------------------------
# Every appointment's tooltip starts with "<Type> <time>", e.g.
# "Follow-Up 7:00am". We read that line for every appointment and
# count how many of each type fall in each hour.

_READ_APPOINTMENTS_JS = r"""
    const visible = el => el.getClientRects().length &&
        getComputedStyle(el).visibility !== 'hidden';
    const countPattern = /^([^()\r\n]+?)\s*\(\s*(\d+)\s*\)$/;
    const readHeaders = () => Array.from(document.querySelectorAll('.calendar-header-slider'))
        .filter(visible).map(h => {
            const r = h.getBoundingClientRect();
            const header = (h.innerText || '').trim();
            const lines = header.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
            const countLine = lines.find(s => countPattern.test(s)) || '';
            const m = countLine.match(countPattern);
            const title = lines.filter(s => !countPattern.test(s)).join(' ');
            const first = lines[0] || '';
            const firstMatch = first.match(countPattern);
            return {name: firstMatch ? firstMatch[1].trim() : first,
                    title, header, type: m ? m[1].trim() : '',
                    left: r.left, right: r.right};
        }).filter(h => h.right - h.left >= 5);
    const firstLine = tip => {
        let lines = Array.from(tip.querySelectorAll('div'))
            .filter(d => !d.querySelector('div'))
            .map(d => (d.textContent || '').trim());
        if (!lines.length)
            lines = (tip.textContent || '').trim().split(/\r?\n/).map(s => s.trim());
        return lines.filter(Boolean).slice(0, 4).join('\n');
    };
    const out = [];
    for (const t of document.querySelectorAll('div.title')) {
        if (!visible(t) || t.closest('.tooltiptext')) continue;
        let tip = null;
        for (let p = t; p && p !== document.body; p = p.parentElement) {
            const found = p.querySelector('.tooltiptext');
            if (found) { tip = found; break; }
            if (!p.parentElement ||
                p.parentElement.querySelectorAll('div.title').length > 1) break;
        }
        t.scrollIntoView({block: 'nearest', inline: 'nearest'});
        const r = t.getBoundingClientRect();
        const cols = readHeaders().filter(h => h.left - 5 <= r.left && r.left <= h.right + 5)
            .sort((a, b) => Math.abs(a.left - r.left) - Math.abs(b.left - r.left));
        out.push({el: t, title: (t.innerText || '').trim(),
                  therapist: cols.length ? cols[0].name : '',
                  column: cols.length ? cols[0] : null,
                  tip: tip ? firstLine(tip) : '',
                  tipText: tip ? (Array.from(tip.querySelectorAll('div'))
                      .filter(d => !d.querySelector('div'))
                      .map(d => (d.textContent || '').trim()).filter(Boolean).join('\n')
                      || (tip.textContent || '')) : ''});
    }
    return out;
"""

_HOVER_TOOLTIP_JS = r"""
    const el = arguments[0], hint = arguments[1], relaxed = arguments[3] === true;
    const norm = s => (s || '').toUpperCase().replace(/[^A-Z]/g, '');
    const firstLine = tip => {
        let lines = Array.from(tip.querySelectorAll('div'))
            .filter(d => !d.querySelector('div'))
            .map(d => (d.textContent || '').trim());
        if (!lines.length)
            lines = (tip.textContent || '').trim().split(/\r?\n/).map(s => s.trim());
        return lines.filter(Boolean).slice(0, 4).join('\n');
    };
    const fullText = tip => {
        const lines = Array.from(tip.querySelectorAll('div'))
            .filter(d => !d.querySelector('div'))
            .map(d => (d.textContent || '').trim()).filter(Boolean);
        return lines.length ? lines.join('\n') : (tip.textContent || '');
    };
    const namePattern = /[A-Za-z][A-Za-z' \-]*,\s*[A-Za-z]/;
    const clean = t => (t || '').split(/\r?\n/)
        .filter(l => !/last\s*change|edited\s*by|created\s*by/i.test(l)).join('\n');
    const own = t => Boolean(hint) && norm(t).includes(hint);
    // a tooltip showing ANOTHER patient's name is never accepted
    // an empty slot's tooltip: "<Type> <time>" then a line with only a number
    const emptySlot = t => /^\s*\d+\s*$/.test((t || '').split('\n')[1] || '');
    const belongs = t => !hint || own(t) || !namePattern.test(clean(t));
    // 1) Tooltip that belongs to this appointment block
    for (let p = el; p && p !== document.body; p = p.parentElement) {
        const tip = p.querySelector('.tooltiptext');
        if (tip) {
            const line = firstLine(tip), text = fullText(tip);
            if (line && !/loading/i.test(line) && belongs(text))
                return {line, text, own: own(text)};
            break;
        }
        if (!p.parentElement ||
            p.parentElement.querySelectorAll('div.title').length > 1) break;
    }
    // 2) Otherwise a visible tooltip near it: one that shows this patient,
    //    or one that just appeared after hovering, close to the block, and
    //    showing no other patient (some tooltips show "Progress Note Needed"
    //    instead of the patient name).

    const TIP_SELECTOR = '.tooltiptext, .x-tip, .tooltip, .qtip, [role="tooltip"]';
    const typeTime = /^[A-Za-z][^\n]*?\d{1,2}:\d{2}\s*(a\.?m\.?|p\.?m\.?)?/i;
    const tipCandidates = () => {
        const out = [];
        for (const t of document.querySelectorAll(TIP_SELECTOR))
            if (!out.some(o => o.contains(t))) out.push(t);    // outermost only
        return out;
    };
    const genericLines = t => (t.innerText || t.textContent || '')
        .split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    const before = new Set(arguments[2] || []);
    const r = el.getBoundingClientRect();
    const cx = (r.left + r.right) / 2, cy = (r.top + r.bottom) / 2;
    let best = null, bestRank = Infinity, bestDist = Infinity;
    for (const tip of tipCandidates()) {
        const tr = tip.getBoundingClientRect();
        if (tr.width < 10 || tr.height < 5) continue;
        const st = getComputedStyle(tip);
        if (st.display === 'none' || st.visibility === 'hidden' ||
            parseFloat(st.opacity || '1') < 0.1) continue;
        const generic = !tip.classList.contains('tooltiptext');
        const gl = generic ? genericLines(tip) : null;
        if (generic && !typeTime.test(gl[0] || '')) continue;   // not an appointment tooltip
        const line = generic ? gl.slice(0, 4).join('\n') : firstLine(tip);
        const text = generic ? gl.join('\n') : fullText(tip);
        if (!line || /loading/i.test(line)) continue;
        const sig = line + '@' + Math.round(tr.left) + ',' + Math.round(tr.top);
        const dist = Math.hypot((tr.left + tr.right) / 2 - cx, (tr.top + tr.bottom) / 2 - cy);
        let rank;
        if (before.has(sig)) continue;                       // was showing before this hover
        if (own(text)) rank = 0;
        else if (!belongs(text)) continue;                    // another patient
        else if (!before.has(sig) && (relaxed || dist <= 450))
            rank = emptySlot(text) ? 2 : 1;   // new, next to it (2 = the group's tooltip)
        else continue;
        if (rank < bestRank || (rank === bestRank && dist < bestDist)) {
            bestRank = rank; bestDist = dist; best = {line, text, own: rank === 0};
        }
    }
    return best;
"""


_NOTE_TITLE = re.compile(
    r"^\W*(opens?\b|close[sd]?\b|do not book|blocked?\b|all appointments)", re.IGNORECASE)


_NAME_WITH_DOB = re.compile(r"[A-Za-z].*,\s*[A-Za-z].*\b\d{1,2}/\d{1,2}/\d{2,4}\b")
_CALENDAR_MARKER = re.compile(r"^\W*calendar\s*(start|end)\b", re.IGNORECASE)


def _exclusion_reason(title, appt_type, patient_line, tip_lines=None):
    """Returns why an entry is NOT a real patient appointment ('' = count it).
       - blocked schedule entries        ('blocked schedule 9:00am' / 'Closed')
       - slots with no patient           ('Follow-Up 11:30am' / '1')
       - green schedule notes / blocks   ('OPEN AT 9AM', 'CLOSE AT 7PM',
                                          'DO NOT BOOK LAST 30 MINS', ...)"""
    title = re.sub(r"[،‚]", ",", title or "").strip()
    first_title_line = title.split("\n")[0].strip()
    patient_line = (patient_line or "").strip()
    if (_CALENDAR_MARKER.search(appt_type or "")
            or _CALENDAR_MARKER.search(first_title_line)):
        return "Calendar Start / Calendar End"
    if re.search(r"block", appt_type or "", re.IGNORECASE):
        return "blocked schedule"
    if patient_line.casefold() in ("closed", "close"):
        return "closed"
    title_has_name = any(re.search(r"[A-Za-z].*,\s*[A-Za-z]", l) for l in title.split("\n"))
    if patient_line and not re.search(r"[A-Za-z]", patient_line) and not title_has_name:
        return f"no patient ('{patient_line}')"
    if _NOTE_TITLE.search(first_title_line):
        return f"schedule note ('{first_title_line[:40]}')"
    # Real appointments show the patient as "LAST, FIRST". A warning such as
    # "Primary Authorization Expired" can come first, so any line holding
    # "LAST, FIRST - MM/DD/YYYY" (title or tooltip) also counts.
    other_lines = [l for l in title.split("\n")[1:] + list(tip_lines or [])[1:]
                   if not re.search(r"last\s*change|edited\s*by|created\s*by", str(l), re.IGNORECASE)]
    if not (re.search(r"[A-Za-z].*,\s*[A-Za-z]", first_title_line)
            or any(_NAME_WITH_DOB.search(str(l)) for l in other_lines)):
        return f"no patient name ('{first_title_line[:40]}')"
    return ""


def _usable_tooltip(line):
    return bool(line) and not re.search(r"loading", line, re.IGNORECASE)


def _name_hint(title):
    """'?Floresluna, Margarita' -> 'FLORESLUNA' (used to check the tooltip
    really belongs to this patient). Looks at every line, so a warning line
    such as 'Primary Authorization Expired' before the name is skipped."""
    for line in str(title or "").split("\n"):
        if "," in line:
            last, rest = line.split(",", 1)
            # "POC Gay, Bertha" -> GAY: badges (POC, PN, ...) in front of the
            # name are dropped; "De La Cruz" -> CRUZ still matches the tooltip
            words = re.findall(r"[A-Za-z'\-]+", last)
            last = re.sub(r"[^A-Z]", "", words[-1].upper()) if words else ""
            if len(last) < 2:
                continue
            # first name = first word after the comma ("Covington, ehab" ->
            # COVINGTONEHAB), so two Covingtons side by side are told apart
            first = re.match(r"\s*([A-Za-z][A-Za-z'\-]*)", rest)
            return last + (re.sub(r"[^A-Z]", "", first.group(1).upper()) if first else "")
    return ""


_OTHER_NAME = re.compile(r"[A-Za-z][A-Za-z' \-]*,\s*[A-Za-z]")


def _tip_belongs(text, hint):
    """True when a tooltip can be this patient's: it shows the patient's
    last name, or it shows no patient name at all. A tooltip showing ANOTHER
    patient's name is rejected. 'Last change: Edited by ...' is ignored."""
    if not hint:
        return True
    text = "\n".join(l for l in str(text or "").split("\n")
                     if not re.search(r"last\s*change|edited\s*by|created\s*by", l, re.I))
    if hint in re.sub(r"[^A-Z]", "", text.upper()):
        return True
    return not _OTHER_NAME.search(text)


_SECTION_START = re.compile(
    r"^[A-Za-z][A-Za-z0-9 &/'()+.\-]*?\s+\d{1,2}:\d{2}\s*(?:a\.?m\.?|p\.?m\.?)?\s*$",
    re.IGNORECASE)


def _tooltip_section(text, hint, nth=0):
    """A tooltip may list several appointments of the same slot, e.g.
         Follow-Up 5:00pm / Smith, John ... / OneOnOne 5:00pm / Doe, Jane ...
    Every "<Type> <time>" line starts a new appointment. Returns
    (lines of THIS patient's appointment, how):
      'single'    - only one appointment in the tooltip
      'matched'   - several; the one holding this patient's name
      'ambiguous' - several and none (or more than one) holds the name"""
    lines = [l.strip() for l in str(text or "").split("\n") if l.strip()]
    starts = [i for i, l in enumerate(lines) if _SECTION_START.match(l)
              and not re.search(r"\b(opens?|close[sd]?)\b", l, re.IGNORECASE)]
    if len(starts) <= 1:
        return lines, "single"
    sections = [lines[s:e] for s, e in zip(starts, starts[1:] + [len(lines)])]
    if hint:
        own = [sec for sec in sections
               if hint in re.sub(r"[^A-Z]", "", " ".join(sec).upper())]
        if len(own) == 1:
            return own[0], "matched"
        if len(own) > 1:
            # the same patient booked twice in this slot: the n-th block of
            # that name takes the n-th appointment that names him/her
            return own[min(nth, len(own) - 1)], "matched"
    return [s[0] for s in sections], "ambiguous"


def _empty_slot_tooltip(text):
    """'Follow-Up 10:30am' / '1' / ... - the tooltip of a slot with no patient."""
    lines = [l.strip() for l in str(text or "").split("\n") if l.strip()]
    return len(lines) > 1 and bool(re.fullmatch(r"\d+", lines[1]))


_BADGES = re.compile(r"\bPOC\b[\s:\-]*", re.IGNORECASE)


def _strip_badges(text):
    """Removes the POC badge from a block's text or a tooltip:
    'POC Gay, Bertha' -> 'Gay, Bertha', 'POC Expiring Plan of Care' ->
    'Expiring Plan of Care'; a line that was only 'POC' disappears."""
    lines = []
    for line in str(text or "").split("\n"):
        line = re.sub(r"\s{2,}", " ", _BADGES.sub("", line)).strip()
        if line:
            lines.append(line)
    return "\n".join(lines)


def _tip_is_own(text, hint):
    return bool(hint) and hint in re.sub(r"[^A-Z]", "", str(text or "").upper())


_VISIBLE_TOOLTIPS_JS = r"""
    const out = [];
    const TIP_SELECTOR = '.tooltiptext, .x-tip, .tooltip, .qtip, [role="tooltip"]';
    const typeTime = /^[A-Za-z][^\n]*?\d{1,2}:\d{2}\s*(a\.?m\.?|p\.?m\.?)?/i;
    const tipCandidates = () => {
        const found = [];
        for (const t of document.querySelectorAll(TIP_SELECTOR))
            if (!found.some(o => o.contains(t))) found.push(t);    // outermost only
        return found;
    };
    const genericLines = t => (t.innerText || t.textContent || '')
        .split(/\r?\n/).map(s => s.trim()).filter(Boolean);

    const firstLine = tip => {
        let lines = Array.from(tip.querySelectorAll('div'))
            .filter(d => !d.querySelector('div'))
            .map(d => (d.textContent || '').trim());
        if (!lines.length)
            lines = (tip.textContent || '').trim().split(/\r?\n/).map(s => s.trim());
        return lines.filter(Boolean).slice(0, 4).join('\n');
    };
    for (const tip of tipCandidates()) {
        const tr = tip.getBoundingClientRect();
        if (tr.width < 10 || tr.height < 5) continue;
        const st = getComputedStyle(tip);
        if (st.display === 'none' || st.visibility === 'hidden' ||
            parseFloat(st.opacity || '1') < 0.1) continue;
        const generic = !tip.classList.contains('tooltiptext');
        const line = generic ? genericLines(tip).slice(0, 4).join('\n') : firstLine(tip);
        if (line) out.push(line + '@' + Math.round(tr.left) + ',' + Math.round(tr.top));
    }
    return out;
"""


def _move_mouse_away(driver):
    """Hover a column header so the previous tooltip closes."""
    from selenium.webdriver.common.action_chains import ActionChains
    try:
        for header in driver.find_elements(By.CSS_SELECTOR, ".calendar-header-slider"):
            if header.is_displayed():
                ActionChains(driver).move_to_element(header).perform()
                time.sleep(0.4)
                return
    except Exception:
        pass


def _hover_tooltip_line(driver, element, hint, timeout=6.0):
    """Hovers the appointment and waits until ITS tooltip has real content.
    Tries twice; the second time it first moves the mouse away so the
    old tooltip closes. Returns (line, full_text, shows_this_patient)."""
    from selenium.webdriver.common.action_chains import ActionChains
    # round 3 = relaxed: any tooltip that opened after the hover, at any
    # distance, as long as it shows no other patient
    for round_number, wait_for in ((1, timeout), (2, timeout + 4), (3, 5.0)):
        if round_number == 2:
            _move_mouse_away(driver)
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center', inline:'nearest'});", element)
        before = driver.execute_script(_VISIBLE_TOOLTIPS_JS) or []
        deadline = time.monotonic() + wait_for
        last_hover = 0.0
        fallback = None
        while time.monotonic() < deadline:
            if time.monotonic() - last_hover > 1.5:
                ActionChains(driver).move_to_element(element).perform()
                last_hover = time.monotonic()
            found = driver.execute_script(_HOVER_TOOLTIP_JS, element, hint, before,
                                          round_number == 3) or {}
            line = (found.get("line") or "") if isinstance(found, dict) else ""
            if _usable_tooltip(line):
                result = (line.strip(), found.get("text") or "", bool(found.get("own")))
                if result[2] or not hint:
                    return result
                # a tooltip WITHOUT this patient's name: keep it as a fallback,
                # but give the patient's own tooltip 2 more seconds to appear
                if fallback is None:
                    fallback = result
                    deadline = min(deadline, time.monotonic() + 2.0)
            time.sleep(0.25)
        if fallback is not None:
            if round_number == 3:
                log.info("  ·  Tooltip taken in the relaxed 3rd round (opened away from the block)")
            return fallback
    try:
        r = driver.execute_script(
            "var r=arguments[0].getBoundingClientRect();return [Math.round(r.left),Math.round(r.top)];",
            element)
        shown = [s.split("\n")[0] + " @" + s.rsplit("@", 1)[-1]
                 for s in (driver.execute_script(_VISIBLE_TOOLTIPS_JS) or [])]
        log.info(f"  ·  No usable tooltip - block at {r}; tooltips on screen: {shown or 'none'}")
    except Exception:
        pass
    return "", "", False


def _parse_tooltip_line(line):
    """'Follow-Up 3:30pm' -> ('Follow-Up', 930, '3:30 PM')
       The middle value is minutes after midnight (used only for sorting).
       no time found       -> (<text>, 9999, 'Unknown')"""
    line = re.sub(r"\s+", " ", line or "").strip()
    if not line:
        return "Unknown", 9999, "Unknown"
    m = re.search(r"(\d{1,2}):(\d{2})\s*(a\.?m\.?|p\.?m\.?)?", line, re.IGNORECASE)
    if not m:
        return line, 9999, "Unknown"
    hour, minute = int(m.group(1)), int(m.group(2))
    suffix = (m.group(3) or "").lower()
    if suffix.startswith("p") and hour != 12:
        hour += 12
    elif suffix.startswith("a") and hour == 12:
        hour = 0
    hour %= 24
    appt_type = line[:m.start()].strip(" -:,") or "Unknown"
    label = f"{(hour % 12) or 12}:{minute:02d} {'AM' if hour < 12 else 'PM'}"
    return appt_type, hour * 60 + minute, label


# ---------------------------------------------------------------------
# 2b) One clinic + one day: select clinic, date, Day view, read appointments
# ---------------------------------------------------------------------
_DAY_VERIFIED = {}     # (clinic, date) -> True when types and counts were confirmed
_EMPTY_STREAK = {}     # date -> how many reads in a row found the day empty


def _verify_day(headers_seen, appointments):
    """Checks one read of a clinic + day. Returns a list of problems ([] = OK).
      - every counted appointment has a readable type and time
      - no two patients got the very same tooltip unless each tooltip shows
        that patient's own name (a repeated tooltip = the wrong one was read)
      - per schedule column, the number counted matches the header count
        "(N)"; entries that are not counted (empty slots, notes, Calendar
        Start/End) may explain a header that is higher."""
    problems = []
    found = list(appointments.values())
    counted = [a for a in found if not a.get("Excluded")]

    unknown = [a for a in counted if a["Hour"] == "Unknown"
               or a["Appointment Type"] in ("", "Unknown")]
    if unknown:
        problems.append(f"{len(unknown)} appointment(s) with no readable type/time")

    by_text = {}
    for a in counted:
        if a.get("_hint") and not a.get("_own") and a.get("_tiptext"):
            by_text.setdefault(a["_tiptext"], set()).add(a["_hint"])
    repeated = sum(len(h) for h in by_text.values() if len(h) > 1)
    if repeated:
        problems.append(f"{repeated} patient(s) got the same tooltip as another patient "
                        f"- their type may be wrong")

    def col_key(a):
        c = a.get("_column")
        return (c.get("title", ""), c.get("type", "")) if c else ("?", "?")

    expected = {}
    for item in headers_seen.values():
        expected[(item["title"], item["type"])] = item["count"]
    got, skipped = {}, {}
    for a in found:
        bucket = skipped if a.get("Excluded") else got
        bucket[col_key(a)] = bucket.get(col_key(a), 0) + 1
    for k in sorted(set(expected) | set(got)):
        e, g, x = expected.get(k, 0), got.get(k, 0), skipped.get(k, 0)
        if not (g <= e <= g + x):
            problems.append(f"column '{' '.join(p for p in k if p)}': header says {e}, "
                            f"counted {g}" + (f" (+{x} not counted)" if x else ""))
    return problems


def _collect_clinic_day(driver, wait, clinic_name, target_date):
    global _current_clinic

    select_element = wait.until(lambda d: _get_clinic_select_element(d))
    options = driver.execute_script(
        "return Array.from(arguments[0].options)"
        ".filter(o => !o.disabled && o.value)"
        ".map(o => ({name: o.text.trim(), value: o.value}));",
        select_element,
    )
    matches = [o for o in options if o["name"].casefold() == clinic_name.casefold()]
    if len(matches) != 1:
        raise RuntimeError(f"Clinic '{clinic_name}' was not uniquely found in the dropdown.")
    target = matches[0]

    def selected_clinic(d):
        element = _get_clinic_select_element(d)
        return element is not None and d.execute_script(
            "return arguments[0].value;", element
        ) == target["value"]

    if not selected_clinic(driver):
        _current_clinic = None
        select_clinic_in_schedule(driver, target["name"])
    wait.until(selected_clinic, "Selected clinic did not match the requested clinic")
    wait_for_page_stable(driver, timeout=WAIT_TIMEOUT)

    slow = WebDriverWait(driver, DAY_LOAD_TIMEOUT, poll_frequency=0.5)
    navigate_schedule_to_date(driver, target_date)
    slow.until(lambda d: _read_selected_calendar_date(d) == target_date,
               f"Calendar did not move to {target_date} within {DAY_LOAD_TIMEOUT:.0f}s")
    _select_day_view(driver, wait)
    slow.until(lambda d: _read_selected_calendar_date(d) == target_date,
               "Calendar date changed after selecting Day view")

    last_label = None
    stable_since = None

    def loaded_pt_label(d):
        nonlocal last_label, stable_since
        busy = d.execute_script("""
            var extBusy = window.Ext && Ext.Ajax &&
                typeof Ext.Ajax.isLoading === 'function' && Ext.Ajax.isLoading();
            var masks = document.querySelectorAll('.x-mask-loading, .x-loading-mask, .x-loadmask');
            return Boolean(extBusy) || Array.from(masks).some(function(el) {
                return el.getClientRects().length && getComputedStyle(el).visibility !== 'hidden';
            });
        """)
        if busy or not selected_clinic(d) or _read_selected_calendar_date(d) != target_date:
            last_label, stable_since = None, None
            return False
        result = _filter_entries_for_date(_read_pt_label(d), target_date)
        if not result:
            last_label, stable_since = None, None
            return False
        # the headers AND the number of appointment blocks must stop changing
        blocks = d.execute_script(
            "return Array.from(document.querySelectorAll('div.title')).filter("
            "t => t.getClientRects().length && !t.closest('.tooltiptext')).length;")
        signature = (result, blocks)
        if signature != last_label:
            last_label, stable_since = signature, time.monotonic()
            return False
        quiet = 15.0 if all(item["count"] == 0 for item in result) else 4.0
        return result if time.monotonic() - stable_since >= quiet else False

    headers_seen = {}
    appointments = {}

    def collect_headers(items):
        for item in items:
            headers_seen[(item["title"], item["type"])] = item

    def collect_appointments():
        """Reads every appointment visible right now. Tooltips that are not
        loaded yet are opened by hovering the appointment."""
        for _ in range(3):
            pending = 0
            occurrences = {}
            for appt in driver.execute_script(_READ_APPOINTMENTS_JS) or []:
                base_key = ((appt.get("column") or {}).get("header", ""),
                            appt["therapist"], appt["title"])
                occurrence = occurrences.get(base_key, 0)
                occurrences[base_key] = occurrence + 1
                key = base_key + (occurrence,)
                if key in appointments:
                    continue
                appt["title"] = _strip_badges(appt["title"])     # POC is ignored
                appt["tip"] = _strip_badges(appt["tip"])
                appt["tipText"] = _strip_badges(appt.get("tipText"))
                hint = _name_hint(appt["title"])
                tip_text = appt.get("tipText") or ""
                line = appt["tip"] if (_usable_tooltip(appt["tip"])
                                       and _tip_belongs(tip_text, hint)) else ""
                is_note = (not hint and (not appt["title"].strip() or bool(_NOTE_TITLE.search(
                    appt["title"].split("\n")[0].strip()))))
                if not line and not is_note:
                    try:
                        line, tip_text, _ = _hover_tooltip_line(driver, appt["el"], hint)
                    except StaleElementReferenceException:
                        pending += 1
                        continue
                if not line and not is_note:
                    warn(f"Tooltip not readable for '{appt['title'][:40]}' "
                         f"({appt['therapist']}) — counted as Unknown")
                line, tip_text = _strip_badges(line), _strip_badges(tip_text)
                tip_lines = [part.strip() for part in line.split("\n")]
                section, how = _tooltip_section(tip_text or line, hint, occurrence)
                if how == "matched":
                    tip_lines = section
                    log.info(f"  ·  Tooltip lists {len(_tooltip_section(tip_text, '')[0])} "
                             f"appointments in this slot - took this patient's: {section[0]}")
                elif how == "ambiguous" and hint:
                    warn(f"Tooltip lists several appointments ({' / '.join(section)}) and "
                         f"none clearly belongs to '{appt['title'][:40]}' - counted as Unknown")
                    tip_lines = ["Unknown"]
                elif how == "ambiguous":
                    tip_lines = [part.strip() for part in (tip_text or line).split("\n") if part.strip()]
                elif section:
                    tip_lines = section
                appt_type, hour24, hour_label = _parse_tooltip_line(tip_lines[0])
                reason = _exclusion_reason(appt["title"], appt_type,
                                           tip_lines[1] if len(tip_lines) > 1 else "",
                                           tip_lines)
                if reason:
                    log.info(f"  ·  Not counted: {reason} | {' / '.join(tip_lines)}")
                column = appt.get("column")
                ot_reason = _ot_reason(column, clinic_name) if column else ""
                appointment_clinic = f"{clinic_name} {OT_SUFFIX}" if ot_reason else clinic_name
                if column is None:
                    warn(f"[{clinic_name}] {target_date}: Column not found for "
                         f"'{appt['title'][:40]}'; keeping normal clinic name")
                if ot_reason and not reason:
                    log.info("[%s] %s %s | Column: %s (OT via %s)",
                             appointment_clinic, hour_label, appt_type,
                             column["header"].replace("\n", " / "), ot_reason)
                appointments[key] = {
                    "Clinic": appointment_clinic,
                    "Excluded": reason,
                    "Therapist": appt["therapist"],
                    "Patient": appt["title"].split("\n")[0],
                    "Tooltip": " | ".join(tip_lines),
                    "Appointment Type": appt_type,
                    "Hour24": hour24,
                    "Hour": hour_label,
                    "_column": column,
                    "_own": _tip_is_own(tip_text, hint),
                    "_hint": hint,
                    "_tiptext": re.sub(r"\s+", " ", tip_text).strip(),
                }
            if not pending:
                return

    def read_everything():
        nonlocal last_label, stable_since
        collect_headers(slow.until(loaded_pt_label, "The day did not finish loading"))
        collect_appointments()

        # Columns that only render when scrolled into view
        panes = driver.execute_script("""
            return Array.from(document.querySelectorAll('div')).filter(el => {
                var style = getComputedStyle(el);
                return el.querySelector('.calendar-header-slider') &&
                    !el.querySelector('#ClinicChange, [name="ClinicChange"], #sideMenuCalendar') &&
                    el.clientWidth > 0 && el.scrollWidth > el.clientWidth + 20 &&
                    /auto|scroll/.test(style.overflowX) && el.getClientRects().length;
            });
        """)
        for pane in panes:
            try:
                original, maximum, width = driver.execute_script(
                    "return [arguments[0].scrollLeft, "
                    "arguments[0].scrollWidth - arguments[0].clientWidth, "
                    "arguments[0].clientWidth];", pane)
                try:
                    position = 0
                    while True:
                        driver.execute_script("arguments[0].scrollLeft = arguments[1];", pane, position)
                        last_label, stable_since = None, None
                        time.sleep(0.5)
                        collect_headers(slow.until(loaded_pt_label, "The day did not finish loading"))
                        collect_appointments()
                        if position >= maximum:
                            break
                        position = min(maximum, position + max(1, int(width * 0.8)))
                finally:
                    driver.execute_script("arguments[0].scrollLeft = arguments[1];", pane, original)
            except StaleElementReferenceException:
                raise RuntimeError("Schedule pane was replaced while scrolling; retrying")

    problems = []
    for read_round in range(1, 4):
        if read_round > 1:
            headers_seen.clear()
            appointments.clear()
            last_label, stable_since = None, None
            _move_mouse_away(driver)
            warn(f"[{clinic_name}] {target_date}: re-reading every appointment "
                 f"(read {read_round}/3)")
        read_everything()
        problems = _verify_day(headers_seen, appointments)
        if not problems:
            if read_round > 1:
                ok(f"[{clinic_name}] {target_date}: types and counts confirmed on read {read_round}")
            break
        for problem in problems:
            warn(f"[{clinic_name}] {target_date} (read {read_round}/3): {problem}")
    only_unreadable = bool(problems) and all("no readable type/time" in p for p in problems)
    if only_unreadable:
        # everything else checked out: write the day without the unreadable
        # appointment(s) rather than dropping the whole day; this day's rows
        # are not reused or deleted this pass (see _INCOMPLETE_SCOPES)
        for a in appointments.values():
            if not a.get("Excluded") and (a["Hour"] == "Unknown"
                                          or a["Appointment Type"] in ("", "Unknown")):
                a["Excluded"] = "tooltip not readable"
                warn(f"[{clinic_name}] {target_date}: '{a['Patient'][:25]}' left out - its "
                     f"tooltip could not be read in 3 reads; the rest of the day is written "
                     f"and no row of this day is deleted this pass")
    _DAY_VERIFIED[(clinic_name, target_date)] = (not problems) or only_unreadable

    wait.until(selected_clinic, "Clinic changed during collection")
    if _read_selected_calendar_date(driver) != target_date:
        raise RuntimeError("Calendar date changed during collection")
    if not headers_seen:
        raise RuntimeError(f"No schedule headers found for {target_date}")

    if (CLINIC_HAS_OT
            and not any(_ot_reason(item, clinic_name) for item in headers_seen.values())):
        warn(f"[{clinic_name}] {target_date}: No OT column found")
    expected = sum(item["count"] for item in headers_seen.values())
    found = [{k: v for k, v in a.items() if not k.startswith("_")}
             for a in appointments.values()]
    counted = sum(1 for a in found if not a.get("Excluded"))
    unknown = sum(1 for a in found if not a.get("Excluded") and a["Hour"] == "Unknown")
    if counted != expected:
        warn(f"Header counts add up to {expected} but {counted} appointment(s) were counted")
    if unknown:
        warn(f"{unknown} appointment(s) still have an unreadable tooltip (Unknown)")
    scope_keys = {(normalize_key_part(v), normalize_key_part(target_date))
                  for v in (clinic_name, f"{clinic_name} {OT_SUFFIX}")}
    if problems:
        _INCOMPLETE_SCOPES.update(scope_keys)
    else:
        _INCOMPLETE_SCOPES.difference_update(scope_keys)
    return found



# ---------------------------------------------------------------------
# 2c) Clinic first, then each day; counts per hour and appointment type
# ---------------------------------------------------------------------
def scrape_clinic_pt_counts(driver, wait, target_dates, records=None, stop_at=None):
    records = records if records is not None else []
    skipped_days = {}          # date -> why it was not written this pass

    for clinic_name in [CLINIC_NAME]:
        if stop_at is not None and _now() >= stop_at:
            warn(f"Shift ended — stopping this pass before {clinic_name}")
            break
        for target_date in target_dates:
            step(f"Reading appointments | {clinic_name} | {target_date}")
            base = {"Clinic": clinic_name, "date_of_visit": target_date}

            attempt = 0
            relogins = 0
            while attempt < 3:
                attempt += 1
                try:
                    if _ensure_logged_in(driver, wait):
                        relogins += 1
                    if attempt > 1 or _get_clinic_select_element(driver) is None:
                        _restore_count_schedule(driver, wait)
                    appointments = _collect_clinic_day(driver, wait, clinic_name, target_date)
                    if not _DAY_VERIFIED.get((clinic_name, target_date), True):
                        for appt in appointments:
                            records.append({**base, **appt, "Status": "UNVERIFIED"})
                        warn(f"{clinic_name} | {target_date}: types/counts could not be "
                             f"confirmed after 3 reads - this day is NOT written to the "
                             f"sheet this pass; its rows stay as they are")
                        skipped_days[target_date] = "types/counts not confirmed"
                        break
                    # nothing at all on the page = maybe not loaded yet; notes,
                    # CLOSED or empty slots on it = it has loaded
                    if not appointments:
                        _EMPTY_STREAK[target_date] = _EMPTY_STREAK.get(target_date, 0) + 1
                        if _EMPTY_STREAK[target_date] < 2:
                            warn(f"{clinic_name} | {target_date}: nothing at all on the page - "
                                 f"not written until the next pass finds it empty too (a slow "
                                 f"page can look empty)")
                            skipped_days[target_date] = "looked empty - waiting for a second empty read"
                            break
                    else:
                        _EMPTY_STREAK.pop(target_date, None)

                    for appt in appointments:
                        appt = dict(appt)
                        reason = appt.pop("Excluded", "")
                        records.append({**base, **appt,
                                        "Status": f"EXCLUDED: {reason}" if reason else "OK"})
                    excluded_count = sum(1 for a in appointments if a.get("Excluded"))
                    unreadable = any(a.get("Excluded") == "tooltip not readable"
                                     for a in appointments)
                    appointments = [a for a in appointments if not a.get("Excluded")]
                    for variant in (clinic_name, f"{clinic_name} {OT_SUFFIX}"):
                        if (variant == clinic_name or CLINIC_HAS_OT) and not unreadable and not any(
                                a["Clinic"] == variant for a in appointments):
                            records.append({**base, "Clinic": variant, "Status": "EMPTY"})

                    summary = {}
                    for appt in appointments:
                        k = (appt["Clinic"], appt["Hour24"], appt["Hour"], appt["Appointment Type"])
                        summary[k] = summary.get(k, 0) + 1
                    for (label, _, hour, appt_type), count in sorted(summary.items()):
                        log.info("  →  %s | %s | %s: %d", f"{label} | {target_date}", hour, appt_type, count)
                    ok(f"{clinic_name} | {target_date}: {len(appointments)} appointment(s) counted, "
                       f"{excluded_count} block(s)/note(s) skipped")
                    break
                except Exception as exc:
                    # Logged out in the middle? Log in again and redo this
                    # same clinic/date without using up an attempt.
                    if relogins < 5 and _is_login_page(driver):
                        relogins += 1
                        attempt -= 1
                        try:
                            _relogin(driver, wait)
                        except Exception as login_exc:
                            fail(f"Re-login failed: {type(login_exc).__name__}: {login_exc}")
                        continue
                    message = str(exc).splitlines()
                    status = f"Error: {type(exc).__name__}: {message[0] if message else 'Unknown error'}"
                    fail(f"{clinic_name} | {target_date} (attempt {attempt}/3): {status}")
                    _save_schedule_failure(driver, clinic_name, target_date, attempt)
                    if attempt < 3:
                        warn(f"Restoring Schedule and retrying {clinic_name} for {target_date}")
                        continue
                    records.append({**base, "Status": status})
                    skipped_days[target_date] = status
                    fail(f"{clinic_name} | {target_date}: not read after 3 attempts - "
                         f"skipped this pass (its rows stay as they are), moving on "
                         f"to the next day")
                    break

            # Publish each completed day before collection moves to the next date.
            save_clinic_counts_excel(records)
            try:
                publish_google_sheet()
            except Exception as exc:
                fail(f"GOOGLE SHEETS UPDATE FAILED after {clinic_name} | {target_date}: "
                     f"{type(exc).__name__}: {exc}")

    ok(f"Clinic hourly counts saved to: {CLINIC_COUNTS_OUTPUT_PATH}")
    written = [d for d in target_dates if d not in skipped_days]
    log.info(f"  -  Days written this pass: {', '.join(written) or 'none'}")
    for day, why in skipped_days.items():
        warn(f"Day NOT written this pass: {day} - {why}")
    failed = [d for d, why in skipped_days.items() if why.startswith("Error")]
    if len(failed) >= 2:
        # several days could not be read at all: the browser session is
        # probably broken - end the pass so the next one starts on a fresh
        # Chrome + login (the days that worked are already written)
        raise RuntimeError(f"{len(failed)} day(s) could not be read: {', '.join(failed)}")
    return records


_gsheet_client = None  # reused across clinics and runs


def _google_sheet_settings():
    # Prefer hourly-specific settings, falling back to the existing Google settings.
    def setting(*names, default=""):
        for name in names:
            for value in (os.environ.get(name), _local_settings.get(name)):
                if value is not None and str(value).strip():
                    return str(value).strip()
        return default

    sheet_url = setting("HOURLY_SHEET_URL", "GOOGLE_SHEET_URL")
    credentials_path = Path(os.environ.get(
        "GOOGLE_SERVICE_ACCOUNT_FILE",
        _local_settings.get("GOOGLE_SERVICE_ACCOUNT_FILE", "google_service_account.json"),
    ))
    if not credentials_path.is_absolute():
        credentials_path = Path(__file__).resolve().parent.parent / credentials_path
    tab_name = CLINIC_SHEET_TAB      # one tab per clinic (set at the top of the file)
    return sheet_url, credentials_path, tab_name


def _gs_client():
    """Authorised gspread client (service account from the AWS secret or the
    local JSON file), built once and reused."""
    global _gsheet_client
    if _gsheet_client is None:
        import gspread
        _, credentials_path, _ = _google_sheet_settings()
        secret_credentials = _google_secret_credentials(credentials_path)
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        if secret_credentials is not None:
            _gsheet_client = gspread.service_account_from_dict(secret_credentials, scopes=scopes)
        else:
            _gsheet_client = gspread.service_account(filename=str(credentials_path), scopes=scopes)
    return _gsheet_client


def _open_google_spreadsheet():
    """Opens the spreadsheet, raising a clear error for each common problem."""
    global _gsheet_client
    sheet_url, credentials_path, _ = _google_sheet_settings()
    if not sheet_url:
        raise RuntimeError(
            f"HOURLY_SHEET_URL and GOOGLE_SHEET_URL are empty. Put the sheet link in {_settings_path} "
            '(e.g. "HOURLY_SHEET_URL": "https://docs.google.com/spreadsheets/d/...") '
            "or set it as an environment variable."
        )
    secret_credentials = _google_secret_credentials(credentials_path)
    try:
        import gspread
    except ImportError:
        raise RuntimeError("gspread is not installed. Run: pip install gspread")

    if _gsheet_client is None:
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        if secret_credentials is not None:
            _gsheet_client = gspread.service_account_from_dict(secret_credentials, scopes=scopes)
        else:
            _gsheet_client = gspread.service_account(filename=str(credentials_path), scopes=scopes)
    try:
        return _api(_gsheet_client.open_by_url, sheet_url)
    except gspread.exceptions.SpreadsheetNotFound:
        email = (secret_credentials if secret_credentials is not None else
                 json.loads(credentials_path.read_text(encoding="utf-8"))).get("client_email", "?")
        raise RuntimeError(
            f"Cannot open the sheet. Share it with {email} as Editor, "
            f"and check the URL: {sheet_url}"
        )


def check_google_sheet():
    """Run once at startup so a Sheets problem shows immediately, not after the first clinic."""
    step("Checking Google Sheets connection")
    try:
        with sheet_lock():
            spreadsheet = _open_google_spreadsheet()
            _, _, tab_name = _google_sheet_settings()
            tabs = [ws.title for ws in _api(spreadsheet.worksheets)]
        ok(f"Google Sheets OK: target tab '{tab_name}' | tabs: {tabs}")
        return True
    except Exception as exc:
        fail(f"GOOGLE SHEETS NOT WORKING: {type(exc).__name__}: {exc}")
        return False


SHEET_COLUMNS = ["Clinic", "date_of_visit", "Time", "Appointment Type", "Count"]


def publish_google_sheet():
    frame = pd.read_excel(CLINIC_COUNTS_OUTPUT_PATH, sheet_name="Hourly Counts").fillna("")
    rows = [[_sheet_cell(v) for v in row] for row in frame[SHEET_COLUMNS].values.tolist()]
    for row in rows:
        row[1] = _norm_sheet_date(row[1])
    today = _now().date().isoformat()
    clinics = {name.strip().casefold() for name in _clinic_variants()}
    rows = [r for r in rows if not (str(r[0]).strip().casefold() in clinics and r[1] == today)]
    # Deduplicate incoming keys before the supplied guard sees the batch.
    key_fn = lambda r: tuple(normalize_key_part(x) for x in r[:4]) if len(r) >= 4 and r[0] != "Clinic" else None
    rows = list({key_fn(row): row for row in rows}.values())
    with sheet_lock():
        spreadsheet = _open_google_spreadsheet()
        _, _, tab_name = _google_sheet_settings()
        worksheets = _api(spreadsheet.worksheets)
        ws = next((tab for tab in worksheets if tab.title.casefold() == tab_name.casefold()), None)
        if ws is None:
            ws = _api(spreadsheet.add_worksheet, title=tab_name, rows=max(100, len(rows) + 1), cols=5)
    # read fresh + rewrite / append / delete, all under one lock
    with sheet_lock():
        sync_visit_rows(ws, rows, SHEET_COLUMNS)


def _norm_sheet_date(value):
    """Normalize a date cell to YYYY-MM-DD so old and new rows compare equal."""
    text = str(value).strip()
    if not text:
        return ""
    try:
        return pd.to_datetime(text).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return text


def _sheet_cell(value):
    """Plain Python values for the sheet (12.0 -> 12)."""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    if hasattr(value, "item"):
        value = value.item()
        if isinstance(value, float) and value.is_integer():
            return int(value)
    return value


def _target_dates(now):
    """The next seven days starting tomorrow, using Cairo time."""
    today = now.date()
    return [(today + pd.Timedelta(days=n)).isoformat()
            for n in range(1, 8)]


def _remove_today_from_saved_counts():
    """Remove today's local rows, then publish under the existing sheet lock."""
    path = Path(CLINIC_COUNTS_OUTPUT_PATH)
    if not path.exists():
        save_clinic_counts_excel([])
    else:
        sheets = pd.read_excel(path, sheet_name=None)
        today = _now().date().isoformat()
        clinics = {name.strip().casefold() for name in _clinic_variants()}
        for name, frame in sheets.items():
            if {"Clinic", "date_of_visit"}.issubset(frame.columns):
                remove = (frame["Clinic"].astype(str).str.strip().str.casefold().isin(clinics)
                          & frame["date_of_visit"].map(_norm_sheet_date).eq(today))
                sheets[name] = frame.loc[~remove]
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for name, frame in sheets.items():
                frame.to_excel(writer, sheet_name=name, index=False)
    publish_google_sheet()


def _open_live_session():
    """Opens Chrome, logs in and opens View Schedule. Used once per shift
    (and again only if the session breaks)."""
    global _current_clinic, _extra_filters_active
    _current_clinic = None
    _extra_filters_active = False
    driver = create_driver()
    wait = WebDriverWait(driver, WAIT_TIMEOUT)
    try:
        login(driver, wait)
    except Exception:
        _close_live_session(driver)
        raise
    return driver, wait


def _close_live_session(driver):
    if driver is None:
        return
    try:
        driver.quit()
        ok("Browser closed")
    except Exception:
        pass


def collect_pass(driver, wait, pass_number, stop_at):
    """One full reading pass over all clinics, in the already-open browser."""
    target_dates = _target_dates(_now())
    _remove_today_from_saved_counts()
    if not target_dates:
        ok("No remaining dates this week after today")
        return
    step(f"PASS #{pass_number} | dates {target_dates[0]} → {target_dates[-1]}")
    started = time.monotonic()
    # Each clinic is pushed to Google Sheets as soon as it finishes
    scrape_clinic_pt_counts(driver, wait, target_dates, stop_at=stop_at)
    minutes = (time.monotonic() - started) / 60
    ok(f"PASS #{pass_number} finished in {minutes:.1f} min")
    put_metric("PassSeconds", minutes * 60, unit="Seconds")
    put_metric("PassFailed", 0)


# =====================================================================
# SCHEDULER: LIVE from 2 PM to 4 AM Cairo time.
#   - 2 PM: open Chrome + log in once, read everything (first pass)
#   - after each pass finishes: wait 3–5 minutes, then read again
#   - 4 AM: close Chrome and sleep until 2 PM
# =====================================================================

SCHEDULE_TIMEZONE = os.environ.get("VISIT_TIMEZONE", "Africa/Cairo")
WINDOW_START_HOUR = 14   # 2 PM
WINDOW_END_HOUR   = 4    # 4 AM (next day)
GAP_MIN_MINUTES   = 3    # wait after each pass: random between
GAP_MAX_MINUTES   = 5    # 3 and 5 minutes
# The MOST to wait for a clicked day to finish loading (slow pages). As soon
# as the day has loaded, the script carries on - it never waits the full time.
DAY_LOAD_TIMEOUT = float(os.environ.get(
    "WEBPT_DAY_LOAD_TIMEOUT", _local_settings.get("WEBPT_DAY_LOAD_TIMEOUT", 120)))
# Chrome is closed and reopened (fresh login) once it is this old - checked
# only between passes, never while a clinic/day is being read.
CHROME_RECYCLE_SECONDS = float(os.environ.get(
    "WEBPT_CHROME_RECYCLE_HOURS", _local_settings.get("WEBPT_CHROME_RECYCLE_HOURS", 4))) * 3600


def _now():
    return pd.Timestamp.now(tz=SCHEDULE_TIMEZONE)


def _in_window(t) -> bool:
    return t.hour >= WINDOW_START_HOUR or t.hour < WINDOW_END_HOUR


def _shift_date(t):
    """The date the current 2 PM - 4 AM shift started on."""
    d = t.normalize()
    if t.hour < WINDOW_END_HOUR:
        d -= pd.Timedelta(days=1)
    return d.date()


def _window_end(t):
    """4 AM that closes the shift containing t."""
    return pd.Timestamp(_shift_date(t), tz=SCHEDULE_TIMEZONE) + pd.Timedelta(days=1, hours=WINDOW_END_HOUR)


def _next_window_start(t):
    start = t.normalize() + pd.Timedelta(hours=WINDOW_START_HOUR)
    return start if start > t else start + pd.Timedelta(days=1)


def _sleep_until(target, reason):
    log.info(f"  ⏸  {reason} — next at {target:%Y-%m-%d %I:%M:%S %p} (Cairo)")
    while True:
        remaining = (target - _now()).total_seconds()
        if remaining <= 0:
            return
        time.sleep(min(remaining, 30))


def purge_past_rows():
    """Removes every row dated today or earlier (Cairo date) from this
    clinic's tab. Runs at the start of every shift (2 PM)."""
    today = _now().date().isoformat()
    step(f"Removing rows dated {today} or earlier from the sheet")
    with sheet_lock():
        if USE_SHEET_LOCK and not _my_token["value"]:
            warn("Sheet lock not owned - past rows are removed at the next shift start")
            return
        spreadsheet = _open_google_spreadsheet()
        _, _, tab_name = _google_sheet_settings()
        ws = next((t for t in _api(spreadsheet.worksheets)
                   if t.title.casefold() == tab_name.casefold()), None)
        if ws is None:
            return
        values = _api(ws.get_all_values) or []
        old = []
        for n, r in enumerate(values[1:], start=2):
            d = normalize_key_part(r[1]) if len(r) > 1 else ""
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", d) and d <= today:
                old.append(n)
        if not old:
            ok(f"No rows dated {today} or earlier in '{tab_name}'")
            return
        ranges = []                       # contiguous blocks, bottom first
        for n in sorted(old, reverse=True):
            if ranges and ranges[-1][0] == n + 1:
                ranges[-1][0] = n
            else:
                ranges.append([n, n])
        requests = [{"deleteDimension": {"range": {
            "sheetId": ws.id, "dimension": "ROWS",
            "startIndex": a - 1, "endIndex": b}}} for a, b in ranges]
        _api(ws.spreadsheet.batch_update, {"requests": requests})
        ok(f"Removed {len(old)} row(s) dated {today} or earlier from '{tab_name}'")


def run_live_shift():
    """Runs one shift: browser stays open, passes repeat until 4 AM."""
    shift_end = _window_end(_now())
    step(f"SHIFT STARTED | LIVE until {shift_end:%Y-%m-%d %I:%M %p} (Cairo)")
    try:
        purge_past_rows()
    except Exception as exc:
        fail(f"Could not remove past rows from the sheet: {type(exc).__name__}: {exc}")
    driver = wait = None
    driver_started_at = None
    pass_number = 0
    failures_in_a_row = 0
    try:
        while _now() < shift_end:
            # ---- soft Chrome recycle: only here, between passes ----------
            if (driver is not None and driver_started_at is not None
                    and CHROME_RECYCLE_SECONDS > 0
                    and time.monotonic() - driver_started_at >= CHROME_RECYCLE_SECONDS):
                age_h = (time.monotonic() - driver_started_at) / 3600
                step(f"Chrome session is {age_h:.1f}h old - closing it now that the "
                     f"previous pass has finished, then logging in again")
                _close_live_session(driver)
                driver = wait = None
                driver_started_at = None

            if driver is None:
                try:
                    driver, wait = _open_live_session()
                    driver_started_at = time.monotonic()
                except Exception:
                    log.exception("Could not open/log in to WebPT")
                    failures_in_a_row += 1
                    retry_at = min(_now() + pd.Timedelta(minutes=min(5 * failures_in_a_row, 15)), shift_end)
                    _sleep_until(retry_at, "Login failed; retrying")
                    continue

            pass_number += 1
            try:
                collect_pass(driver, wait, pass_number, shift_end)
                failures_in_a_row = 0
            except Exception:
                log.exception(f"PASS #{pass_number} failed — reopening the browser for the next pass")
                put_metric("PassFailed", 1)
                failures_in_a_row += 1
                _close_live_session(driver)
                driver = wait = None

            if _now() >= shift_end:
                break
            time.sleep(cycle_sleep_seconds(3, 5))
    finally:
        step("SHIFT ENDED (4 AM) — closing browser")
        _close_live_session(driver)


def run_scheduler():
    check_google_sheet()
    step(f"SCHEDULER STARTED | LIVE {WINDOW_START_HOUR}:00 → {WINDOW_END_HOUR}:00 Cairo | "
         f"{GAP_MIN_MINUTES}–{GAP_MAX_MINUTES} min between passes")
    while True:
        now = _now()
        if not _in_window(now):
            _sleep_until(_next_window_start(now), "Outside 2 PM – 4 AM window")
            continue
        try:
            run_live_shift()
        except Exception:
            # A broken shift must not stop the scheduler
            log.exception("Shift crashed — restarting")
            time.sleep(60)


def _configure_instance():
    """Sets up this script's clinic (CLINIC_NAME at the top of the file):
    its own log file, Excel file and WebPT account."""
    global USERNAME, PASSWORD, CURRENT_CLINIC
    clinic = CLINIC_NAME
    CURRENT_CLINIC = clinic
    slug = CLINIC_SLUG
    Path(LOCAL_WORK_DIR).mkdir(parents=True, exist_ok=True)

    accounts = _local_settings.get("CLINIC_ACCOUNTS", {})
    account = next((v for k, v in accounts.items() if k.casefold() == clinic.casefold()), None)
    if account and not WEBPT_SECRET_ID:
        USERNAME = account.get("WEBPT_USERNAME", USERNAME)
        PASSWORD = account.get("WEBPT_PASSWORD", PASSWORD)

    root = logging.getLogger()
    for handler in list(root.handlers):
        root.removeHandler(handler)
        handler.close()
    formatter = logging.Formatter(f"%(asctime)s [{clinic}] [%(levelname)s] %(message)s")
    log_folder = Path(LOG_DIR)
    log_folder.mkdir(parents=True, exist_ok=True)
    for handler in (logging.FileHandler(log_folder / f"webpt_{slug}.log", encoding="utf-8"),
                    logging.StreamHandler()):
        handler.setFormatter(formatter)
        root.addHandler(handler)
    log.info(f"Instance for clinic '{clinic}' | Excel: {CLINIC_COUNTS_OUTPUT_PATH} | "
             f"credentials from '{CREDENTIALS_SHEET_TAB}' | headless={HEADLESS}")


def _handle_sigterm(signum, frame):
    log.info(f"Signal {signum} received - closing the browser and stopping")
    raise SystemExit(0)      # runs the finally blocks: Chrome quits, lock released


def main():
    signal.signal(signal.SIGTERM, _handle_sigterm)
    try:
        run_scheduler()
    except KeyboardInterrupt:
        log.info("Scheduler stopped")


if __name__ == "__main__":
    os.environ.setdefault("SCRIPT_LABEL", Path(__file__).stem)
    SCRIPT_LABEL = os.environ["SCRIPT_LABEL"]
    _configure_instance()
    main()
