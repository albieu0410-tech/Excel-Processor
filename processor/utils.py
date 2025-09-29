from typing import Optional, Dict, Any, List, Callable
from datetime import datetime
import locale
import os
import re
import json
import time
import requests

LOG_DIR = "logs"

# Ensure log folder exists
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)


def log_message(widget, message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted = f"[{timestamp}] {message}\n"

    # Display in GUI
    widget.configure(state="normal")
    widget.insert("end", formatted)
    widget.configure(state="disabled")
    widget.see("end")

    # Write to log file
    append_to_log_file(formatted)


def append_to_log_file(message):
    now = datetime.now().strftime("%Y-%m-%d_%H%M")
    log_file = os.path.join(LOG_DIR, f"run_{now}.log")
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(message)


def normalize_str(s):
    """Normalize a string by lowercasing and removing non-alphanumeric characters."""
    if not isinstance(s, str):
        return ""
    return re.sub(r"\W+", "", s.lower())


MONTH_3 = {
    "jan": "Jan",
    "feb": "Feb",
    "mar": "Mar",
    "apr": "Apr",
    "may": "May",
    "jun": "Jun",
    "jul": "Jul",
    "aug": "Aug",
    "sep": "Sep",
    "oct": "Oct",
    "nov": "Nov",
    "dec": "Dec",
}


def safe_float(x, default=0.0):
    """Convert strings like '€ 1,234.56' or '1.234,56' to float safely."""
    if x is None:
        return default
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if not s:
        return default
    s = (
        s.replace("€", "")
        .replace("$", "")
        .replace("£", "")
        .replace("\u00a0", " ")
        .strip()
    )
    # handle EU decimals: '1.234,56' -> '1234.56'
    if "," in s and s.rfind(",") > s.rfind("."):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except Exception:
        return default


def normalize_currency(value, default="EUR"):
    """Normalize currency strings; default to EUR if missing."""
    if value is None:
        return default
    s = str(value).strip().upper()
    fixes = {
        "EURO": "EUR",
        "EUROS": "EUR",
        "EUE": "EUR",
        "USD$": "USD",
        "US$": "USD",
        "$": "USD",
        "POUND": "GBP",
        "£": "GBP",
    }
    return fixes.get(s, s or default)


def log_verbose(log, message):
    """Wrapper to allow future control over debug-level logs."""
    log(message)


def format_summary_sheet(ws):
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Alignment, Font

    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    for cell in ws[1]:
        cell.font = Font(bold=True)


# --- SharePoint config & Deal ID helpers ---
CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".excel_processor_config.json")


def load_config():
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_config(cfg: dict):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        # Non-fatal: ignore persistence errors
        pass


def normalize_deal_id(value) -> str:
    """Return a clean Deal ID string suitable for matching folder prefixes."""
    if value is None:
        return ""
    s = str(value).strip()
    # keep digits, letters, underscore, dash, slash
    s = re.sub(r"[^\w\-\/]", "", s)
    return s


# --- Date parsing helpers ---
MONTH_ABBRS = {
    "jan",
    "feb",
    "mar",
    "apr",
    "may",
    "jun",
    "jul",
    "aug",
    "sep",
    "oct",
    "nov",
    "dec",
}
MONTH_MAP = {
    # handle English short/long month names
    "jan": "01",
    "january": "01",
    "feb": "02",
    "february": "02",
    "mar": "03",
    "march": "03",
    "apr": "04",
    "april": "04",
    "may": "05",
    "jun": "06",
    "june": "06",
    "jul": "07",
    "july": "07",
    "aug": "08",
    "august": "08",
    "sep": "09",
    "sept": "09",
    "september": "09",
    "oct": "10",
    "october": "10",
    "nov": "11",
    "november": "11",
    "dec": "12",
    "december": "12",
}


def try_parse_datetime(s: str) -> Optional[datetime]:
    """
    Accepts many variants, including '16/Jan/2015 21:35:30' and '16-01-2015 21:35'.
    Returns a datetime or None.
    """
    if not s:
        return None
    s = str(s).strip()

    # Normalize month abbreviations
    parts = s.split(" ")
    if parts and "/" in parts[0]:
        dparts = parts[0].split("/")
        if len(dparts) >= 2 and dparts[1].lower() in MONTH_ABBRS:
            dparts[1] = dparts[1].title()
            parts[0] = "/".join(dparts)
            s = " ".join(parts)

    # Try numeric month first
    fmts = [
        # Existing formats
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%d/%b/%Y %H:%M:%S",
        "%d/%b/%Y %H:%M",
        "%d-%b-%Y %H:%M:%S",
        "%d-%b-%Y %H:%M",
        "%d/%b/%Y",
        "%d-%b-%Y",
        # Add new formats
        "%d.%m.%Y %H:%M:%S",  # German style
        "%d.%m.%Y",  # German style (date only)
        "%Y/%m/%d %H:%M:%S",  # Asian style
        # US-style month/day
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y",
    ]

    for fmt in fmts:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return None


parse_any_datetime = try_parse_datetime  # Alias for compatibility


# --- Status checking helpers ---
EXCLUDED_STATUSES = {
    "failed",
    "declined",
    "rejected",
    "canceled",
    "cancelled",
    "pending",
}
SUCCESS_STATUSES = {
    "dep_settled",
    "pay_closed",
    "rdm_closed",
    "closed",
    "settled",
    "completed",
    "success",
    "paid",
}


def is_failed_status(s: Optional[str]) -> bool:
    """Check if a status indicates failure/rejection/pending."""
    s = (s or "").strip().lower()
    return any(token in s for token in EXCLUDED_STATUSES)


def is_success_status(s: Optional[str]) -> bool:
    """Return True for positive/settled statuses only. Empty/None => False (caller may accept 'auto')."""
    if not s:
        return False
    s = str(s).strip().lower()
    # Quick negatives first
    bad_tokens = (
        "reject",
        "cancel",
        "decline",
        "fail",
        "pending",
        "hold",
        "on hold",
        "onhold",
    )
    if any(tok in s for tok in bad_tokens):
        return False
    # Expanded positives across providers (Unibet, Bwin, PokerStars)
    good_tokens = (
        "dep_settled",
        "settled",
        "completed",
        "complete",
        "success",
        "approved",
        "ok",
        "pay_closed",
        "rdm_closed",
        "closed",
        "done",
        "paid",
        "posted",  # PokerStars / some ledgers
        "processed",  # PokerStars
    )
    return any(tok in s for tok in good_tokens)


def excel_fmt_date(cell):
    """Apply German-style date format DD.MM.YYYY to a cell that already holds a date/datetime."""
    cell.number_format = "DD.MM.YYYY"


def excel_fmt_number(cell):
    """Apply #,##0.00 format with thousand separators."""
    cell.number_format = "#,##0.00"


def beautify_with_xlwings(file_path: str, log=print):
    """
    If xlwings & Excel are available, apply native Excel styling
    to 'Payment Summary' and 'Summary'. Safe no-op if xlwings missing.
    """
    try:
        import xlwings as xw
    except Exception as e:
        log(f"ℹ️ xlwings not available, skipping Excel styling: {e}")
        return

    app = xw.App(visible=False)
    try:
        wb = app.books.open(file_path)

        def style_sheet(name, header_fill=0x4472C4):
            if name not in [s.name for s in wb.sheets]:
                return
            sht = wb.sheets[name]

            # detect used range
            used = sht.used_range
            if used.count == 1 and (used.value is None or used.value == ""):
                return

            # header row
            hdr = sht.range("A1").expand("right")
            hdr.api.Font.Bold = True
            hdr.api.Interior.Color = header_fill
            hdr.api.Font.Color = 0xFFFFFF

            # formats
            headers = [str(v).strip().lower() for v in hdr.value]
            if "date" in headers:
                idx = headers.index("date") + 1
                sht.range((2, idx), (sht.cells.last_cell.row, idx)).number_format = (
                    "dd.mm.yyyy"
                )
            if "amount" in headers:
                idx = headers.index("amount") + 1
                sht.range((2, idx), (sht.cells.last_cell.row, idx)).number_format = (
                    "#,##0.00"
                )

            # freeze top row
            sht.api.Rows("2:2").Select()
            app.api.ActiveWindow.FreezePanes = True

            # autofit & make it pretty
            sht.autofit()

        style_sheet("Payment Summary", header_fill=0x5B9BD5)
        style_sheet("Summary", header_fill=0x4472C4)

        wb.save()
    finally:
        try:
            wb.close()
        except Exception:
            pass
        app.quit()


# =============================================================================
#                         API HELPERS (AC + IT)
# Can either return messages (for GUI to iterate) or stream to a logger.
# =============================================================================


def _env(name: str, default: str = "") -> str:
    v = os.getenv(name)
    return v if (v is not None and v != "") else default


def _build_url(base: str, *parts: str) -> str:
    base = (base or "").rstrip("/")
    tail = "/".join(p.strip("/") for p in parts if p is not None and str(p) != "")
    return f"{base}/{tail}" if tail else base


def _short(txt: str, n: int = 220) -> str:
    return (txt or "")[:n].replace("\r", " ").replace("\n", " ")


# ---------- IT API ----------
def _it_headers() -> Dict[str, str]:
    """
    Determine IT auth header.
    Token envs (either works):
      - IT_AUTH_TOKEN  (preferred)
      - IT_API_KEY
    Header name:
      - IT_API_AUTH_HEADER (default: "Authorization")
    If header is "Authorization" and token doesn't start with "Bearer ", we add it.
    """
    token = _env("IT_AUTH_TOKEN") or _env("IT_API_KEY")
    if not token:
        return {}
    header_name = _env("IT_API_AUTH_HEADER", "Authorization")
    if header_name.lower() == "authorization" and not token.lower().startswith(
        "bearer "
    ):
        token = f"Bearer {token}"
    return {header_name: token}


def it_get(
    path: str = "",
    params: Optional[Dict[str, Any]] = None,
    timeout_s: Optional[int] = None,
) -> requests.Response:
    """
    GET https://<IT_API_BASE_URL><IT_API_BASE_PATH>/<path> with long read timeout.
    Respects IT_API_TIMEOUT (default 90s).
    Retries 3x on transient request errors.
    """
    base = _env("IT_API_BASE_URL").rstrip("/")
    if not base:
        raise RuntimeError("IT_API_BASE_URL is not set.")
    base_path = _env("IT_API_BASE_PATH", "").strip()
    if base_path and not base_path.startswith("/"):
        base_path = "/" + base_path
    url = _build_url(base, base_path, path or "")

    read_to = int(timeout_s or int(_env("IT_API_TIMEOUT", "90")))  # default 90s
    sess = requests.Session()

    last_error: Optional[Exception] = None
    for attempt in range(3):  # small retry loop for transient gateway issues
        try:
            return sess.get(
                url, headers=_it_headers(), params=params, timeout=(15, read_to)
            )
        except requests.RequestException as e:
            last_error = e
            if attempt < 2:
                time.sleep(1.2 * (attempt + 1))
    # If we got here, all retries failed
    raise last_error if last_error else RuntimeError("IT request failed")


def probe_it_base_messages() -> List[str]:
    """
    Calls the IT base endpoint with a long timeout and returns concise messages.
    """
    msgs: List[str] = []
    base = _env("IT_API_BASE_URL").rstrip("/")
    bp = _env("IT_API_BASE_PATH", "").strip()
    if bp and not bp.startswith("/"):
        bp = "/" + bp
    shown_base = base + bp
    msgs.append(f"--- IT ({shown_base}) ---")
    msgs.append(
        f"🔧 IT probe → {shown_base}  (read timeout {_env('IT_API_TIMEOUT','90')}s)"
    )

    try:
        r = it_get("")  # base path (e.g., https://api.golden-tech.de/glts_ac)
        code = r.status_code
        js: Any = {}
        try:
            js = r.json()
        except Exception:
            js = {}

        keys = list(js.keys()) if isinstance(js, dict) else []
        data_items = (
            len(js.get("data", []))
            if isinstance(js, dict) and isinstance(js.get("data"), list)
            else 0
        )

        if 200 <= code < 300:
            msgs.append(
                f"✅ IT base [{code}] keys={keys or '(n/a)'} data_items={data_items}"
            )
        else:
            msgs.append(
                f"⚠️ IT base [{code}] keys={keys or '(n/a)'} data_items={data_items}"
            )
            if isinstance(js, dict) and "detail" in js:
                msgs.append(
                    "ℹ️ IT response contains 'detail' — check Authorization header or required filters."
                )

        # peek first item if present
        if isinstance(js, dict) and isinstance(js.get("data"), list) and js["data"]:
            item = js["data"][0]
            cid = str(item.get("id", ""))
            title = str(item.get("title", ""))[:60]
            notes = item.get("notes") or []
            msgs.append(f"ℹ️ IT sample — id: {cid}, title: {title}, notes: {len(notes)}")
            if notes:
                n0 = notes[0]
                ntxt = str(n0.get("note", ""))[:80]
                msgs.append(f"   └─ first note: {ntxt}")

    except requests.Timeout as e:
        msgs.append(
            f"❌ IT base request timed out ({e}). Raise IT_API_TIMEOUT (now {_env('IT_API_TIMEOUT','90')}s)."
        )
    except Exception as e:
        msgs.append(f"❌ IT base request failed: {e}")

    return msgs


# ---------- ActiveCampaign ----------
def _ac_headers() -> Dict[str, str]:
    key = _env("AC_API_KEY")
    headers = {"Accept": "application/json", "Content-Type": "application/json"}
    if key:
        headers["Api-Token"] = key
    return headers


def ac_root() -> str:
    base = _env("AC_API_BASE_URL").rstrip("/")
    if not base:
        base = _env("AC_API_URL").rstrip("/")
    if not base:
        raise RuntimeError("AC_API_BASE_URL / AC_API_URL is not set.")
    if not base.endswith("/api/3"):
        base = f"{base}/api/3"
    return base


def ac_get(
    path: str, params: Optional[Dict[str, Any]] = None, timeout_s: int = 30
) -> requests.Response:
    url = _build_url(ac_root(), path.strip("/"))
    sess = requests.Session()
    return sess.get(url, headers=_ac_headers(), params=params, timeout=(10, timeout_s))


def ac_list_deal_groups_and_stages_messages() -> List[str]:
    msgs: List[str] = []
    try:
        g = ac_get("dealGroups", params={"limit": 100, "offset": 0})
        if g.ok:
            js = g.json()
            groups = len(js.get("dealGroups", []))
            stages = len(js.get("dealStages", []))
            msgs.append(f"✅ AC /dealGroups OK — groups: {groups}, stages: {stages}")
        else:
            msgs.append(f"⚠️ AC /dealGroups {g.status_code}: {_short(g.text)}")
    except Exception as e:
        msgs.append(f"❌ AC /dealGroups error: {e}")

    try:
        s = ac_get("dealStages", params={"limit": 200, "offset": 0})
        if s.ok:
            js = s.json()
            msgs.append(
                f"✅ AC /dealStages OK — {len(js.get('dealStages', []))} stages"
            )
        else:
            msgs.append(f"⚠️ AC /dealStages {s.status_code}: {_short(s.text)}")
    except Exception as e:
        msgs.append(f"❌ AC /dealStages error: {e}")

    try:
        d = ac_get("deals", params={"limit": 5, "offset": 0})
        if d.ok:
            js = d.json()
            msgs.append(f"✅ AC /deals OK — sample size: {len(js.get('deals', []))}")
        else:
            msgs.append(f"⚠️ AC /deals {d.status_code}: {_short(d.text)}")
    except Exception as e:
        msgs.append(f"❌ AC /deals error: {e}")

    return msgs


def test_apis(log: Optional[Callable[[str], Any]] = None) -> Optional[List[str]]:
    """
    One-click probe used by the GUI.

    Usage:
      - Streaming: test_apis(self.log)         # logs as it goes, returns None
      - Collecting: msgs = test_apis()         # returns List[str] for you to iterate

    Env respected:
      AC_API_BASE_URL or AC_API_URL, AC_API_KEY
      IT_API_BASE_URL, IT_API_BASE_PATH, IT_AUTH_TOKEN or IT_API_KEY
      IT_API_TIMEOUT (default 90), IT_API_AUTH_HEADER (default Authorization)
    """
    messages: List[str] = []

    # --- AC ---
    ac_base = _env("AC_API_BASE_URL") or _env("AC_API_URL")
    if ac_base:
        messages.append(f"--- ActiveCampaign ({ac_root()}) ---")
        messages.extend(ac_list_deal_groups_and_stages_messages())
    else:
        messages.append("ℹ️ AC base not provided; skipped.")

    # --- IT ---
    it_base = _env("IT_API_BASE_URL")
    it_path = _env("IT_API_BASE_PATH", "")
    if it_base:
        messages.extend(probe_it_base_messages())
    else:
        messages.append("ℹ️ IT base not provided; skipped.")

    # Stream or return
    if log is not None:
        for m in messages:
            try:
                log(m)
            except Exception:
                print(m)
        return None
    return messages
