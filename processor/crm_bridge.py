# ac_field_test.py
# ─────────────────────────────────────────────────────────────────────────────
# (1) Your crm_bridge.py content FIRST (with fixes: added `import re`, fixed a loop typo)
# (2) Field + notes tester CLI (list fields, list/add/delete deal notes, set field)
# ─────────────────────────────────────────────────────────────────────────────

# crm_bridge.py
# Unifies access to your internal IT API (notes) + ActiveCampaign (deals/stages)
# Console logging only (no GUI changes). Lists ALL deals in the selected stage.

from __future__ import annotations

import os
import time
import json
import logging
from dataclasses import dataclass
from typing import Iterable, Set, List, Dict, Optional, Tuple, Any
import re  # needed by _looks_checked_text

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# --- NEW imports requested ---
import locale
from datetime import datetime

try:
    from zoneinfo import ZoneInfo  # Python 3.9+
except Exception:
    ZoneInfo = None
# --- /NEW ---

# Reuse your existing config + AC helpers (same ones your selector/tests use)
load_config = None
ac_get = None
ac_root = None
_ac_headers = None
try:
    from utils import (
        load_config as _lc,
        ac_get as _ag,
        ac_root as _ar,
        _ac_headers as _ah,
    )

    load_config, ac_get, ac_root, _ac_headers = _lc, _ag, _ar, _ah
except Exception:
    try:
        from processor.utils import (
            load_config as _lc,
            ac_get as _ag,
            ac_root as _ar,
            _ac_headers as _ah,
        )

        load_config, ac_get, ac_root, _ac_headers = _lc, _ag, _ar, _ah
    except Exception:
        pass

CHECKED_PREFIX = "Checked_"
DEFAULT_NOTE_KEYS = (
    "Note",
    "Notes",
    "Kommentar",
    "Comment",
    "Status",
    "Checked",
    "Checked Note",
    "CheckedNote",
)


def _normalize_str(s: Optional[str]) -> str:
    return (s or "").strip().lower()


def _looks_checked_text(text: str) -> bool:
    """
    Heuristic to decide if a text means 'checked'.
    - Case-insensitive 'checked'
    - Emojis/ticks like '✅', '✔', '✓'
    - German variants 'geprüft', 'kontrolliert'
    """
    t = _normalize_str(text)
    if not t:
        return False
    patterns = [
        r"\bchecked\b",
        r"\bgepr(?:ue|ü)ft\b",
        r"\bkontrolliert\b",
        r"[✅✔✓]",
    ]
    return any(re.search(p, t) for p in patterns)


def _deal_has_checked_note(
    deal: Dict[str, Any],
    note_keys: Iterable[str] = DEFAULT_NOTE_KEYS,
) -> bool:
    """
    Returns True if the given deal dict contains a 'checked' note
    in any of the typical note/comment/status fields.
    """
    for key in note_keys:
        if key in deal and _looks_checked_text(str(deal.get(key, ""))):
            return True
    # Fallback: scan all string fields
    for k, v in deal.items():
        if isinstance(v, str) and _looks_checked_text(v):
            return True
    return False


def _pick_deal_identifier(deal: Dict[str, Any]) -> Optional[str]:
    """
    Choose the best identifier for filename matching.
    Adjust order/keys to match your pipeline.
    """
    candidate_keys = [
        "AZ Datev",
        "AZ",
        "AZ_Long",
        "AZ_Short",
        "Aktenzeichen",
        "Aktenzeichen Datev",
        "AZDatev",
        "dealTitle",
        "title",
        "Title",
        "name",
        "Name",
        "Filename",
        "PDF",
        "Id",
        "ID",
        "dealId",
    ]
    for key in candidate_keys:
        val = str(deal.get(key, "")).strip()
        if val:
            return val
    return None


def _extract_identifiers_from_deals(deals: List[Dict[str, Any]]) -> Set[str]:
    ids: Set[str] = set()
    for d in deals:
        ident = _pick_deal_identifier(d)
        if ident:
            ids.add(ident.lower())
    return ids


# --- Optional SharePoint access via Microsoft Graph ---


class _SharePointClient:
    """
    Minimal Microsoft Graph client to list files in a SharePoint folder.

    Config (via utils.load_config):
      - TENANT_ID
      - CLIENT_ID
      - CLIENT_SECRET
      - SP_SITE_ID  (or SP_SITE_HOSTNAME + SP_SITE_PATH)
      - SP_DRIVE_ID (optional; defaults to site's default drive)
      - SP_FOLDER_PATH
    """

    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        site_id: Optional[str] = None,
        site_hostname: Optional[str] = None,
        site_path: Optional[str] = None,
        drive_id: Optional[str] = None,
        folder_path: Optional[str] = None,
        timeout: int = 30,
    ) -> None:
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.site_id = site_id
        self.site_hostname = site_hostname
        self.site_path = site_path
        self.drive_id = drive_id
        self.folder_path = folder_path
        self.timeout = timeout
        self._token: Optional[str] = None

    def _get_token(self) -> str:
        import requests

        if self._token:
            return self._token
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        }
        resp = requests.post(url, data=data, timeout=self.timeout)
        resp.raise_for_status()
        self._token = resp.json()["access_token"]
        return self._token

    @property
    def headers(self) -> Dict[str, str]:
        return {"Authorization": f"Bearer {self._get_token()}"}

    def _ensure_site_id(self) -> str:
        import requests

        if self.site_id:
            return self.site_id
        if not (self.site_hostname and self.site_path):
            raise ValueError(
                "Must provide either SP_SITE_ID OR (SP_SITE_HOSTNAME + SP_SITE_PATH)."
            )
        url = f"https://graph.microsoft.com/v1.0/sites/{self.site_hostname}:/sites/{self.site_path}"
        r = requests.get(url, headers=self.headers, timeout=self.timeout)
        r.raise_for_status()
        self.site_id = r.json()["id"]
        return self.site_id

    def list_files_in_folder(self) -> List[Dict[str, Any]]:
        import requests

        site_id = self._ensure_site_id()
        if not self.drive_id:
            drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"
            r = requests.get(drive_url, headers=self.headers, timeout=self.timeout)
            r.raise_for_status()
            self.drive_id = r.json()["id"]

        if not self.folder_path:
            url = (
                f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root/children"
            )
        else:
            url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{self.folder_path}:/children"

        items: List[Dict[str, Any]] = []
        next_url = url
        while next_url:
            r = requests.get(next_url, headers=self.headers, timeout=self.timeout)
            r.raise_for_status()
            data = r.json()
            items.extend(data.get("value", []))
            next_url = data.get("@odata.nextLink")
        return items


def _find_checked_prefix_matches(
    sp_items: List[Dict[str, Any]],
    candidate_ids: Set[str],
    prefix: str = CHECKED_PREFIX,
) -> Set[str]:
    """
    Returns the subset of candidate_ids that appear in any filename beginning with prefix.
    Matching is case-insensitive and allows the identifier to appear anywhere after the prefix.
    Example filename: 'Checked_000760-2023_foo.pdf' → matches id '000760-2023'
    """
    matched: Set[str] = set()
    pl = prefix.lower()
    for it in sp_items:
        name = _normalize_str(it.get("name"))
        if not name.startswith(pl):
            continue
        for cid in candidate_ids:
            if cid in name:
                matched.add(cid)
    return matched


def filter_unchecked_deals(
    deals: List[Dict[str, Any]],
    use_sharepoint: bool = True,
    checked_prefix: str = CHECKED_PREFIX,
    note_keys: Iterable[str] = DEFAULT_NOTE_KEYS,
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    Filters out deals that:
      1) contain a 'Checked' note
      2) OR have a file in SharePoint starting with 'Checked_' that contains the deal's identifier.

    Returns (filtered_deals, stats)
    """
    # (1) Drop deals with checked note
    without_checked_note: List[Dict[str, Any]] = []
    dropped_by_note: List[Dict[str, Any]] = []
    for d in deals:
        if _deal_has_checked_note(d, note_keys):
            dropped_by_note.append(d)
        else:
            without_checked_note.append(d)

    dropped_by_sp: List[Dict[str, Any]] = []
    if use_sharepoint:
        # Safe config load
        sp_items: List[Dict[str, Any]] = []
        try:
            # Reuse your existing config loader
            from utils import load_config  # uses your project's config

            cfg = load_config()
            tenant = cfg.get("TENANT_ID")
            client_id = cfg.get("CLIENT_ID")
            client_secret = cfg.get("CLIENT_SECRET")
            site_id = cfg.get("SP_SITE_ID")
            site_hostname = cfg.get("SP_SITE_HOSTNAME")
            site_path = cfg.get("SP_SITE_PATH")
            drive_id = cfg.get("SP_DRIVE_ID")
            folder_path = cfg.get("SP_FOLDER_PATH")

            if (
                tenant
                and client_id
                and client_secret
                and (site_id or (site_hostname and site_path))
            ):
                sp = _SharePointClient(
                    tenant_id=tenant,
                    client_id=client_id,
                    client_secret=client_secret,
                    site_id=site_id,
                    site_hostname=site_hostname,
                    site_path=site_path,
                    drive_id=drive_id,
                    folder_path=folder_path,
                )
                # List once per click
                sp_items = sp.list_files_in_folder()
            else:
                # Missing SP config → skip SP filtering silently
                sp_items = []
        except Exception:
            # Any SP error → skip SP filtering silently (don't break the GUI click)
            sp_items = []

        if sp_items:
            candidate_ids = _extract_identifiers_from_deals(without_checked_note)
            matched_ids = _find_checked_prefix_matches(
                sp_items, candidate_ids, checked_prefix
            )
            keep: List[Dict[str, Any]] = []
            for d in without_checked_note:
                cid = _pick_deal_identifier(d)
                if cid and cid.lower() in matched_ids:
                    dropped_by_sp.append(d)
                else:
                    keep.append(d)
            final_deals = keep
        else:
            final_deals = without_checked_note
    else:
        final_deals = without_checked_note

    stats = {
        "input_count": len(deals),
        "dropped_by_note_count": len(dropped_by_note),
        "dropped_by_sharepoint_count": len(dropped_by_sp),
        "output_count": len(final_deals),
        "checked_prefix": checked_prefix,
    }
    return final_deals, stats


# =========================
# Console logging helpers
# =========================


def _env(name: str, default: str = "") -> str:
    v = os.getenv(name)
    return v if (v is not None and v != "") else default


def _ensure_console_logger(name: str) -> logging.Logger:
    """
    Create a console logger if not already configured.
    Level controlled by CRM_BRIDGE_LOGLEVEL (default INFO).
    """
    log = logging.getLogger(name)
    if not any(isinstance(h, logging.StreamHandler) for h in log.handlers):
        handler = logging.StreamHandler()
        fmt = logging.Formatter("%(asctime)s %(levelname)s [%(name)s] %(message)s")
        handler.setFormatter(fmt)
        log.addHandler(handler)
    level = _env("CRM_BRIDGE_LOGLEVEL", "INFO").upper()
    try:
        log.setLevel(getattr(logging, level))
    except Exception:
        log.setLevel(logging.INFO)
    log.propagate = False
    return log


ROOT_LOG = _ensure_console_logger("CRMBridge")
IT_LOG = _ensure_console_logger("CRMBridge.IT")
AC_LOG = _ensure_console_logger("CRMBridge.AC")

# Thresholds for slow-call warnings
IT_WARN_S = float(_env("IT_API_WARN_S", "3.0"))
AC_WARN_S = float(_env("AC_API_WARN_S", "2.0"))


# =========================
# Configuration
# =========================


@dataclass(frozen=True)
class ITApiConfig:
    base_url: str = _env("IT_API_BASE_URL", "").rstrip("/")
    api_key: str = _env("IT_API_KEY", "") or _env("IT_AUTH_TOKEN", "")


@dataclass(frozen=True)
class ACConfig:
    # kept for completeness; actual requests go through utils.ac_get/ac_root
    base_url: str = (_env("AC_API_BASE_URL") or _env("AC_API_URL") or "").rstrip("/")
    api_key: str = _env("AC_API_KEY", "")


# =========================
# HTTP Client (shared)
# =========================


def _build_session(
    timeout: int = 20, total_retries: int = 3, backoff: float = 0.5
) -> requests.Session:
    """
    Shared session with sensible retries for idempotent requests.
    """
    session = requests.Session()
    retries = Retry(
        total=total_retries,
        backoff_factor=backoff,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET", "POST", "PUT", "PATCH"]),
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    # store a default timeout on the session via attribute
    session.request = _wrap_with_timeout(session.request, timeout=timeout)  # type: ignore
    return session


def _wrap_with_timeout(fn, timeout: int):
    def wrapper(method, url, **kwargs):
        if "timeout" not in kwargs:
            kwargs["timeout"] = timeout
        return fn(method, url, **kwargs)

    return wrapper


class APIError(RuntimeError):
    def __init__(
        self,
        message: str,
        status: Optional[int] = None,
        payload: Optional[Dict[str, Any]] = None,
    ):
        super().__init__(message)
        self.status = status
        self.payload = payload or {}


# =========================
# IT Team API (Leads + notes via BASE endpoint)
# =========================


class ITApiClient:
    """
    There is NO /notes endpoint. Notes live on the lead/case objects returned by the BASE endpoint.

    IT guidance:
      - No 'limit' and no 'filters[id]'.
      - If you want filters by Lead ID, use '?id=<ID>'.
      - You can fetch ALL leads (no params).

    Allowed query params:
      id, title, owner, FirstName, LastName, BirthDate, Title, Country, State,
      Email, Opponent, GamblingProvider
    """

    ALLOWED_FILTERS = {
        "id",
        "title",
        "owner",
        "FirstName",
        "LastName",
        "BirthDate",
        "Title",
        "Country",
        "State",
        "Email",
        "Opponent",
        "GamblingProvider",
    }

    def __init__(
        self,
        cfg: Optional[ITApiConfig] = None,
        session: Optional[requests.Session] = None,
    ):
        self.cfg = cfg or ITApiConfig()
        if not self.cfg.base_url or not self.cfg.api_key:
            raise ValueError(
                "IT_API_BASE_URL and IT_API_KEY/IT_AUTH_TOKEN must be set."
            )
        it_timeout = int(_env("IT_API_TIMEOUT", "90"))
        self.session = session or _build_session(timeout=it_timeout)
        self.log = IT_LOG

    def _headers(self) -> Dict[str, str]:
        header_name = _env("IT_API_AUTH_HEADER", "Authorization")
        token = self.cfg.api_key
        if (
            header_name.lower() == "authorization"
            and token
            and not token.lower().startswith("bearer ")
        ):
            token = f"Bearer {token}"
        return {
            header_name: token,
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

    def _log_http(
        self,
        label: str,
        url: str,
        params: Optional[Dict[str, Any]],
        resp: Optional[requests.Response],
        t0: float,
    ):
        dt = time.perf_counter() - t0
        status = resp.status_code if resp is not None else "ERR"
        size = (
            len(resp.content) if (resp is not None and resp.content is not None) else 0
        )
        msg = (
            f"{label} {url} params={params or {}} -> {status} in {dt:.3f}s body={size}B"
        )
        if dt >= IT_WARN_S:
            self.log.warning(msg)
        else:
            self.log.info(msg)

    def list_leads(self, **filters: str) -> List[Dict[str, Any]]:
        """
        Fetch the full list or a filtered list using ONLY allowed IT filters.
        """
        base = self.cfg.base_url.rstrip("/")
        params = {
            k: v for k, v in (filters or {}).items() if k in self.ALLOWED_FILTERS and v
        }
        t0 = time.perf_counter()
        r = self.session.get(base, headers=self._headers(), params=params or None)
        self._log_http("IT GET", base, params or None, r, t0)
        r.raise_for_status()
        js = r.json() or {}
        data = js.get("data", js)
        if isinstance(data, list):
            return data
        if isinstance(data, dict):
            return [data]
        return []

    def list_notes(self, case_id: str) -> List[Dict[str, Any]]:
        """
        Prefer a single filtered call (?id=) then fall back to full list and detail.
        """
        base = self.cfg.base_url.rstrip("/")

        # 1) Filtered by id
        try:
            t0 = time.perf_counter()
            r = self.session.get(
                base, headers=self._headers(), params={"id": str(case_id)}
            )
            self._log_http("IT GET", base, {"id": str(case_id)}, r, t0)
            if r.ok:
                js = r.json() or {}
                data = js.get("data", js)
                if isinstance(data, list):
                    for item in data:
                        if str(item.get("id")) == str(case_id):
                            return item.get("notes") or []
                if isinstance(data, dict) and str(data.get("id")) == str(case_id):
                    return data.get("notes") or []
        except Exception as e:
            self.log.error("IT base?id call failed: %s", e)

        # 2) Full list
        try:
            t1 = time.perf_counter()
            r2 = self.session.get(base, headers=self._headers())
            self._log_http("IT GET", base, None, r2, t1)
            if r2.ok:
                js = r2.json() or {}
                data = js.get("data", js)
                if isinstance(data, list):
                    for item in data:
                        if str(item.get("id")) == str(case_id):
                            return item.get("notes") or []
                if isinstance(data, dict) and str(data.get("id")) == str(case_id):
                    return data.get("notes") or []
        except Exception as e:
            self.log.error("IT full-list call failed: %s", e)

        # 3) Detail fallback
        url_detail = f"{base}/{case_id}"
        try:
            t2 = time.perf_counter()
            r3 = self.session.get(url_detail, headers=self._headers())
            self._log_http("IT GET", url_detail, None, r3, t2)
            if r3.ok:
                dj = r3.json() or {}
                if isinstance(dj, dict):
                    if "data" in dj and isinstance(dj["data"], dict):
                        return dj["data"].get("notes") or []
                    return dj.get("notes") or []
        except Exception as e:
            self.log.error("IT detail call failed: %s", e)

        return []

    def add_note(self, *args, **kwargs):
        raise NotImplementedError(
            "Your IT API has no /notes writer endpoint; implement if available."
        )


# =========================
# ActiveCampaign API (Deals & Stages via utils.ac_get)
# =========================


class ActiveCampaignClient:
    """
    Uses utils.ac_get()/ac_root() so '/api/3' is appended exactly once (same as your selector).
    """

    def __init__(
        self,
        cfg: Optional[ACConfig] = None,
        session: Optional[requests.Session] = None,
        rate_limit_s: float = 0.0,
    ):
        self.cfg = cfg or ACConfig()
        if not (self.cfg.base_url or os.getenv("AC_API_URL")) or not self.cfg.api_key:
            raise ValueError("AC_API_BASE_URL/AC_API_URL and AC_API_KEY must be set.")
        ac_timeout = int(_env("AC_API_TIMEOUT", "20"))
        self.session = session or _build_session(timeout=ac_timeout)
        self.rate_limit_s = rate_limit_s
        self.log = AC_LOG

    def _headers(self) -> Dict[str, str]:
        if _ac_headers:
            return _ac_headers()
        return {
            "Api-Token": self.cfg.api_key,
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

    def _rl(self):
        if self.rate_limit_s > 0:
            time.sleep(self.rate_limit_s)

    def _log_http(
        self,
        label: str,
        path: str,
        params: Optional[Dict[str, Any]],
        resp: Optional[requests.Response],
        t0: float,
    ):
        dt = time.perf_counter() - t0
        status = resp.status_code if resp is not None else "ERR"
        size = (
            len(resp.content) if (resp is not None and resp.content is not None) else 0
        )
        msg = f"{label} {path} params={params or {}} -> {status} in {dt:.3f}s body={size}B"
        if dt >= AC_WARN_S:
            self.log.warning(msg)
        else:
            self.log.info(msg)

    # --- NEW helper to normalize root for write ops ---
    def _root(self) -> str:
        return ac_root() if ac_root else (self.cfg.base_url.rstrip("/") + "/api/3")

    # ---------------- Read helpers ----------------

    def list_stages(self, pipeline_id: Optional[str] = None) -> List[Dict[str, Any]]:
        params = {"limit": 200, "offset": 0}
        if pipeline_id:
            params["filters[dealGroupid]"] = str(pipeline_id)
        if ac_get:
            self._rl()
            t0 = time.perf_counter()
            r = ac_get("dealStages", params=params)
            self._log_http("AC GET", "dealStages", params, r, t0)
            if not r.ok:
                raise APIError(f"AC list_stages failed: {r.text}", status=r.status_code)
            return r.json().get("dealStages", []) or []
        # Fallback
        root = self._root()
        self._rl()
        t0 = time.perf_counter()
        resp = self.session.get(
            f"{root}/dealStages", headers=self._headers(), params=params
        )
        self._log_http("AC GET", f"{root}/dealStages", params, resp, t0)
        if not resp.ok:
            raise APIError(
                f"AC list_stages failed: {resp.text}", status=resp.status_code
            )
        return resp.json().get("dealStages", []) or []

    def stage_id_by_name(
        self, name: str, pipeline_id: Optional[str] = None
    ) -> Optional[str]:
        name_norm = (name or "").strip().lower()
        for st in self.list_stages(pipeline_id=pipeline_id):
            if str(st.get("title", "")).strip().lower() == name_norm:
                return str(st.get("id"))
        return None

    def list_deals(self, params: Dict[str, Any]) -> Dict[str, Any]:
        """
        Low-level passthrough for GET /deals with params already shaped.
        """
        if ac_get:
            self._rl()
            t0 = time.perf_counter()
            r = ac_get("deals", params=params)
            self._log_http("AC GET", "deals", params, r, t0)
            if not r.ok:
                raise APIError(f"AC list_deals failed: {r.text}", status=r.status_code)
            return r.json()

        root = self._root()
        url = f"{root}/deals"
        self._rl()
        t0 = time.perf_counter()
        resp = self.session.get(url, headers=self._headers(), params=params)
        self._log_http("AC GET", url, params, resp, t0)
        if not resp.ok:
            raise APIError(
                f"AC list_deals failed: {resp.text}", status=resp.status_code
            )
        return resp.json()

    # ---------------- Write helpers (NEW) ----------------

    def move_deal_to_stage(self, deal_id: str, stage_id: str) -> dict:
        """
        PATCH /deals/{id} with {"deal":{"stage": <stage_id>}}
        """
        url = f"{self._root()}/deals/{deal_id}"
        payload = {"deal": {"stage": str(stage_id)}}
        resp = self.session.patch(
            url, headers=self._headers(), data=json.dumps(payload), timeout=(10, 30)
        )
        if not resp.ok:
            raise APIError(
                f"move_deal_to_stage failed: {resp.text}", status=resp.status_code
            )
        try:
            return resp.json()
        except Exception:
            return {"ok": True, "status_code": resp.status_code, "text": resp.text}

    def create_deal_note(self, deal_id: str, text: str) -> dict:
        """
        Prefer POST /deals/{id}/notes; fallback to /notes with reltype=Deal.
        """
        url = f"{self._root()}/deals/{deal_id}/notes"
        payload = {"note": {"note": text}}
        r = self.session.post(
            url, headers=self._headers(), data=json.dumps(payload), timeout=(10, 30)
        )
        if r.ok:
            try:
                return r.json()
            except Exception:
                return {"ok": True, "status_code": r.status_code, "text": r.text}
        # fallback
        url2 = f"{self._root()}/notes"
        payload2 = {"note": {"note": text, "relid": int(deal_id), "reltype": "Deal"}}
        r2 = self.session.post(
            url2, headers=self._headers(), data=json.dumps(payload2), timeout=(10, 30)
        )
        if not r2.ok:
            raise APIError(f"create_deal_note failed: {r2.text}", status=r2.status_code)
        try:
            return r2.json()
        except Exception:
            return {"ok": True, "status_code": r2.status_code, "text": r2.text}

    def upsert_deal_custom_field(self, deal_id: str, field_id: str, value: str) -> dict:
        """
        Create or update Deal Custom Field value.
        First try to find an existing datum; if found, PATCH it, else POST.
        """
        # Try to find existing datum via utils.ac_get if available
        dcv_id = None
        if ac_get:
            params_variants = [
                {
                    "filters[dealId]": deal_id,
                    "filters[customFieldId]": field_id,
                    "limit": 100,
                },
                {
                    "filters[dealid]": deal_id,
                    "filters[customfieldid]": field_id,
                    "limit": 100,
                },
            ]
            for params in params_variants:
                r = ac_get("dealCustomFieldData", params=params, timeout_s=30)
                if not r or not r.ok:
                    continue
                items = (r.json() or {}).get("dealCustomFieldData") or []
                for it in items:
                    if str(it.get("dealId")) == str(deal_id) and str(
                        it.get("customFieldId")
                    ) == str(field_id):
                        dcv_id = it.get("id")
                        break
                if dcv_id:
                    break

        base = self._root()
        if dcv_id:
            url = f"{base}/dealCustomFieldData/{dcv_id}"
            payload = {"dealCustomFieldDatum": {"fieldValue": str(value)}}
            pr = self.session.patch(
                url, headers=self._headers(), data=json.dumps(payload), timeout=(10, 30)
            )
            if not pr.ok:
                raise APIError(
                    f"upsert_deal_custom_field (update) failed: {pr.text}",
                    status=pr.status_code,
                )
            try:
                return pr.json()
            except Exception:
                return {"ok": True, "status_code": pr.status_code, "text": pr.text}

        url = f"{base}/dealCustomFieldData"
        payload = {
            "dealCustomFieldDatum": {
                "dealId": str(deal_id),
                "customFieldId": str(field_id),
                "fieldValue": str(value),
            }
        }
        cr = self.session.post(
            url, headers=self._headers(), data=json.dumps(payload), timeout=(10, 30)
        )
        if not cr.ok:
            raise APIError(
                f"upsert_deal_custom_field (create) failed: {cr.text}",
                status=cr.status_code,
            )
        try:
            return cr.json()
        except Exception:
            return {"ok": True, "status_code": cr.status_code, "text": cr.text}


# =========================
# Bridge (unified façade)
# =========================


# --- NEW helper: today's date string in Berlin ---
def _today_berlin_str() -> str:
    """
    Returns today's date formatted as dd.mm.yyyy in Europe/Berlin timezone when possible.
    """
    try:
        # Optional: set German locale; harmless if not present
        locale.setlocale(locale.LC_TIME, "de_DE.utf8")
    except Exception:
        pass
    if ZoneInfo:
        try:
            now = datetime.now(ZoneInfo("Europe/Berlin"))
        except Exception:
            now = datetime.now()
    else:
        now = datetime.now()
    return now.strftime("%d.%m.%Y")


# --- /NEW ---


class CRMBridge:
    """
    Single entry point to wire workflows that touch BOTH systems.
    """

    # --- NEW constants ---
    FIELD_ID_CAS_TOTAL_DIFF = "75"  # CAS Total Difference (currency)
    DEFAULT_NOTE_TEMPLATE = "Checked ({date})"
    # --- /NEW ---

    def __init__(
        self,
        it_client: Optional[ITApiClient] = None,
        ac_client: Optional[ActiveCampaignClient] = None,
        logger: Optional[logging.Logger] = None,
    ):
        self.it = it_client or ITApiClient()
        self.ac = ac_client or ActiveCampaignClient()
        self.log = logger or ROOT_LOG

    # ---------- Helpers ----------

    def _load_ac_selection(self) -> Tuple[str, Optional[str], Optional[str]]:
        """
        Load pipeline & stage selection from ~/.excel_processor_config.json.
        Returns (pipeline_id, stage_id, stage_title).
        """
        cfg: Dict[str, Any] = {}
        if load_config:
            try:
                cfg = load_config() or {}
            except Exception:
                cfg = {}
        else:
            path = os.path.join(os.path.expanduser("~"), ".excel_processor_config.json")
            try:
                with open(path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
            except Exception:
                cfg = {}

        ac_sel = cfg.get("activecampaign") or {}
        pipeline_id = (
            str(ac_sel.get("pipeline_id"))
            if ac_sel.get("pipeline_id") is not None
            else ""
        )
        stage_id = (
            str(ac_sel.get("stage_id")) if ac_sel.get("stage_id") is not None else None
        )
        stage_title = ac_sel.get("stage_title")
        if not pipeline_id or (not stage_id and not stage_title):
            raise ValueError(
                "Missing ActiveCampaign selection in ~/.excel_processor_config.json "
                "(need 'activecampaign.pipeline_id' AND ('stage_id' or 'stage_title'))."
            )
        return pipeline_id, stage_id, stage_title

    def _fmt_ts(self, note: Dict[str, Any]) -> str:
        return str(
            note.get("createdAt")
            or note.get("created_at")
            or note.get("created")
            or note.get("date")
            or note.get("timestamp")
            or ""
        ).strip()

    def _all_deals_in_stage(
        self, stage_id: str, pipeline_id: Optional[str]
    ) -> List[Dict[str, Any]]:
        """
        Fetch all deals in a stage using pagination until exhaustion.
        """
        self.log.info(
            "Fetching ALL deals for pipeline=%s stage_id=%s",
            pipeline_id or "ANY",
            stage_id,
        )
        results: List[Dict[str, Any]] = []
        offset = 0
        page_size = 100  # AC typically supports up to 100
        while True:
            params: Dict[str, Any] = {
                "limit": page_size,
                "offset": offset,
                "filters[stage]": stage_id,
            }
            if pipeline_id:
                params["filters[pipeline]"] = pipeline_id
            # Do NOT set status => fetch all (open/won/lost)
            data = self.ac.list_deals(params=params)
            deals = data.get("deals", []) or []
            results.extend(deals)

            meta = data.get("meta", {}) if isinstance(data, dict) else {}
            total = None
            try:
                total = int(meta.get("total")) if "total" in meta else None
            except Exception:
                total = None

            self.log.info(
                "AC batch offset=%d size=%d (total=%s)",
                offset,
                len(deals),
                total if total is not None else "unknown",
            )
            if not deals:
                break
            offset += page_size
            if (total is not None and offset >= total) or len(deals) < page_size:
                break

        self.log.info("Fetched %d deals total for stage=%s", len(results), stage_id)
        return results

    # ---------- Public: ALL deals in selected stage with organized notes ----------
    # NOTE: accept 'limit' for backward compatibility, but ignore it (we return ALL).
    def deals_in_selected_stage_with_notes(
        self, limit: Optional[int] = None
    ) -> List[str]:
        """
        - Reads pipeline + stage from config
        - Fetches ALL deals in that stage (no limit)
        - Fetches ALL IT leads once and maps notes by id
        - Returns pretty, grouped lines ready to print

        The 'limit' argument is accepted for backward compatibility but ignored.
        """
        if limit is not None:
            self.log.info(
                "deals_in_selected_stage_with_notes: 'limit' argument (%s) is ignored; returning ALL deals.",
                limit,
            )

        pipeline_id, stage_id, stage_title = self._load_ac_selection()
        self.log.info(
            "Building list for pipeline=%s stage_id=%s (%s)",
            pipeline_id,
            stage_id or "?",
            stage_title or "",
        )

        # Resolve stage_id if missing
        if not stage_id and stage_title:
            stage_id = self.ac.stage_id_by_name(stage_title, pipeline_id=pipeline_id)
            if not stage_id:
                raise ValueError(
                    f"Stage '{stage_title}' could not be resolved for pipeline {pipeline_id}."
                )

        # 1) All AC deals in the stage
        deals = self._all_deals_in_stage(stage_id, pipeline_id)

        # 2) One IT call to fetch all leads and map notes by id
        it_leads = self.it.list_leads()
        notes_by_id: Dict[str, List[Dict[str, Any]]] = {}
        for lead in it_leads or []:
            lid = str(lead.get("id"))
            if lid:
                notes_by_id[lid] = lead.get("notes") or []

        # 3) Formatting helpers
        def _indent_multiline(text: str, prefix: str = "    ") -> List[str]:
            # Preserve multi-line notes; indent subsequent lines
            text = text.replace("\r\n", "\n").replace("\r", "\n")
            return [
                prefix + line if i > 0 else line
                for i, line in enumerate(text.split("\n"))
            ]

        def format_block(d: Dict[str, Any]) -> List[str]:
            did = str(d.get("id"))
            title = (d.get("title") or "").strip() or f"(deal {did})"
            notes = notes_by_id.get(did, [])

            # Sort newest first; empties last
            def key_ts(n):
                ts = self._fmt_ts(n)
                return (0, ts) if ts else (1, "")

            ordered = sorted(notes, key=key_ts, reverse=True)

            lines: List[str] = []
            if ordered:
                first = ordered[0]
                first_ts = self._fmt_ts(first)
                first_text = str(
                    first.get("content") or first.get("note") or first.get("text") or ""
                )
                ts_box = f"[{first_ts}]" if first_ts else "[]"
                header = f"{ts_box} - {did} | {title} | [] {first_text}"
                lines.extend(
                    _indent_multiline(header, prefix="")
                )  # header may itself be multiline

                for n in ordered[1:]:
                    nts = self._fmt_ts(n)
                    ntext = str(
                        n.get("content") or n.get("note") or n.get("text") or ""
                    )
                    nbox = f"[{nts}]" if nts else "[]"
                    for line in _indent_multiline(f"{nbox} {ntext}"):
                        lines.append(line)
            else:
                lines.append(f"[] - {did} | {title} | (no notes)")

            lines.append("")  # blank line between deals
            return lines

        out_lines: List[str] = []
        for d in deals:
            out_lines.extend(format_block(d))

        self.log.info("Prepared %d line(s) for output.", len(out_lines))
        return out_lines

    # ---------------- NEW: preview/apply combined changes ----------------

    def preview_changes(
        self,
        deal_id: str,
        target_stage_id: Optional[str],
        add_checked_note: bool,
        custom_note_text: Optional[str],
        cas_total_diff: Optional[float],
    ) -> Dict[str, Any]:
        """
        Return a dict summarizing what will be posted, without doing anything.
        GUI shows this for confirmation/edit.
        """
        date_str = _today_berlin_str()
        note_text = (
            custom_note_text or ""
        ).strip() or self.DEFAULT_NOTE_TEMPLATE.format(date=date_str)
        return {
            "deal_id": str(deal_id),
            "target_stage_id": str(target_stage_id) if target_stage_id else None,
            "add_note": bool(add_checked_note),
            "note_text": note_text if add_checked_note else None,
            "set_field": (
                self.FIELD_ID_CAS_TOTAL_DIFF if cas_total_diff is not None else None
            ),
            "field_value": (
                float(cas_total_diff) if cas_total_diff is not None else None
            ),
        }

    def apply_changes(
        self,
        changes: Dict[str, Any],
        dry_run: bool = False,
    ) -> Dict[str, Any]:
        """
        Apply the requested changes in a safe order:
          1) Move stage (if provided)
          2) Set CAS Total Difference (field 75) if provided
          3) Add note (if requested)
        Returns a results dict per step.
        """
        deal_id = str(changes.get("deal_id"))
        target_stage_id = changes.get("target_stage_id")
        add_note = bool(changes.get("add_note"))
        note_text = changes.get("note_text")
        field_id = changes.get("set_field")
        field_value = changes.get("field_value")

        results: Dict[str, Any] = {"dry_run": dry_run, "deal_id": deal_id, "steps": []}

        if target_stage_id:
            step = {"action": "move_stage", "stage_id": str(target_stage_id)}
            if not dry_run:
                step["result"] = self.ac.move_deal_to_stage(
                    deal_id, str(target_stage_id)
                )
            results["steps"].append(step)

        if field_id and field_value is not None:
            step = {
                "action": "set_field",
                "field_id": str(field_id),
                "value": float(field_value),
            }
            if not dry_run:
                # AC currency fields accept string numbers; keep two decimals
                step["result"] = self.ac.upsert_deal_custom_field(
                    deal_id, str(field_id), f"{float(field_value):.2f}"
                )
            results["steps"].append(step)

        if add_note and (note_text or "").strip():
            step = {"action": "add_note", "text": note_text}
            if not dry_run:
                step["result"] = self.ac.create_deal_note(deal_id, note_text)
            results["steps"].append(step)

        return results


# Optional quick CLI for testing this module directly (console-only)
if __name__ == "__main__" and False:
    bridge = CRMBridge()
    try:
        lines = bridge.deals_in_selected_stage_with_notes(
            limit=5
        )  # limit accepted but ignored
        print("\n".join(lines))
    except Exception as e:
        ROOT_LOG.error("Error: %s", e)


# ─────────────────────────────────────────────────────────────────────────────
# Field + Notes tester CLI (list fields, list/add/delete deal notes, set field)
# + NEW: preview/apply combined changes
# ─────────────────────────────────────────────────────────────────────────────

import argparse
import sys  # ← needed for sys.exit

# Reuse the same utils import pattern for the tester part too
try:
    from processor.utils import (
        ac_root as _ac_root2,
        _ac_headers as _ac_headers2,
        ac_get as _ac_get2,
    )
except Exception:
    try:
        from utils import ac_root as _ac_root2, _ac_headers as _ac_headers2, ac_get as _ac_get2  # type: ignore
    except Exception:
        _ac_root2 = ac_root
        _ac_headers2 = _ac_headers
        _ac_get2 = ac_get


def _pretty(js: Any) -> str:
    try:
        return json.dumps(js, ensure_ascii=False, indent=2)
    except Exception:
        return str(js)


# ------------------------- Deal custom fields ------------------------- #


def list_deal_custom_fields() -> List[Dict[str, Any]]:
    """Fetch all deal custom fields."""
    base = _ac_root2() if _ac_root2 else ac_root()
    url = f"{base}/dealCustomFieldMeta"
    r = requests.get(
        url, headers=_ac_headers2() if _ac_headers2 else _ac_headers(), timeout=20
    )
    if not r.ok:
        raise RuntimeError(f"list_deal_custom_fields failed: {r.text}")
    return r.json().get("dealCustomFieldMeta", []) or []


def list_deal_notes(deal_id: str) -> List[Dict[str, Any]]:
    """GET /deals/{id}/notes"""
    base = _ac_root2() if _ac_root2 else ac_root()
    url = f"{base}/deals/{deal_id}/notes"
    r = requests.get(
        url, headers=_ac_headers2() if _ac_headers2 else _ac_headers(), timeout=20
    )
    if not r.ok:
        raise RuntimeError(f"list_deal_notes failed: {r.text}")
    return r.json().get("notes", []) or []


def add_deal_note(deal_id: str, text: str) -> Dict[str, Any]:
    """POST /deals/{id}/notes (simple wrapper)"""
    base = _ac_root2() if _ac_root2 else ac_root()
    url = f"{base}/deals/{deal_id}/notes"
    payload = {"note": {"note": text}}
    r = requests.post(
        url,
        headers=_ac_headers2() if _ac_headers2 else _ac_headers(),
        data=json.dumps(payload),
        timeout=20,
    )
    if not r.ok:
        raise RuntimeError(f"add_deal_note failed: {r.text}")
    return r.json()


def delete_deal_note(note_id: str) -> Dict[str, Any]:
    """DELETE /notes/{id}"""
    base = _ac_root2() if _ac_root2 else ac_root()
    url = f"{base}/notes/{note_id}"
    r = requests.delete(
        url, headers=_ac_headers2() if _ac_headers2 else _ac_headers(), timeout=20
    )
    if not r.ok:
        raise RuntimeError(f"delete_deal_note failed: {r.text}")
    return {"ok": True, "status": r.status_code}


def set_field_value(deal_id: str, field_id: str, value: str) -> Dict[str, Any]:
    """Create or update deal custom field value via the upsert logic."""
    client = ActiveCampaignClient()
    return client.upsert_deal_custom_field(deal_id, field_id, value)


def main_cli():
    ap = argparse.ArgumentParser(description="ActiveCampaign fields/notes test CLI")
    ap.add_argument("--deal-id", help="Deal ID to operate on")
    ap.add_argument(
        "--list-fields", action="store_true", help="List deal custom fields"
    )
    ap.add_argument("--list-notes", action="store_true", help="List notes of a deal")
    ap.add_argument("--add-note", help="Add note text to a deal")
    ap.add_argument("--delete-note", help="Delete note by note ID")
    ap.add_argument(
        "--set-field", help="Set field value; requires --field-id and --field-value"
    )
    ap.add_argument("--field-id", help="Field ID to set")
    ap.add_argument("--field-value", help="Field value")

    # --- NEW quick flow flags ---
    ap.add_argument(
        "--preview-change",
        action="store_true",
        help="Preview AC changes (stage, note, CAS diff) without applying",
    )
    ap.add_argument(
        "--apply-change",
        action="store_true",
        help="Apply AC changes (stage, note, CAS diff)",
    )
    ap.add_argument("--target-stage-id", help="Target Stage ID to move the deal to")
    ap.add_argument(
        "--add-checked-note",
        action="store_true",
        help="Include a 'Checked (dd.mm.yyyy)' note",
    )
    ap.add_argument(
        "--note-text",
        help="Override note text (defaults to 'Checked (dd.mm.yyyy)' if --add-checked-note)",
    )
    ap.add_argument(
        "--cas-total-diff",
        type=float,
        help="Value for field #75 (CAS Total Difference)",
    )
    # --- /NEW ---

    args = ap.parse_args()

    # 1) list fields
    if args.list_fields:
        fields = list_deal_custom_fields()
        print(_pretty(fields))
        return

    # 2) list notes
    if args.list_notes:
        if not args.deal_id:
            print("ERROR: --list-notes requires --deal-id")
            sys.exit(2)
        notes = list_deal_notes(args.deal_id)
        print(_pretty(notes))
        return

    # 3) add note
    if args.add_note:
        if not args.deal_id:
            print("ERROR: --add-note requires --deal-id")
            sys.exit(2)
        res = add_deal_note(args.deal_id, args.add_note)
        print(_pretty(res))
        return

    # 4) delete note
    if args.delete_note:
        res = delete_deal_note(args.delete_note)
        print(_pretty(res))
        return

    # 5) set custom field
    if args.set_field:
        if not args.deal_id or not args.field_id or args.field_value is None:
            print("ERROR: --set-field requires --deal-id, --field-id and --field-value")
            sys.exit(2)
        res = set_field_value(args.deal_id, args.field_id, str(args.field_value))
        print(_pretty(res))
        return

    # 7) Preview/apply combined changes (move stage, set CAS diff #75, add note)
    if args.preview_change or args.apply_change:
        if not args.deal_id:
            print("ERROR: --preview-change/--apply-change require --deal-id")
            sys.exit(2)
        bridge = CRMBridge()
        preview = bridge.preview_changes(
            deal_id=args.deal_id,
            target_stage_id=args.target_stage_id,
            add_checked_note=bool(args.add_checked_note),
            custom_note_text=args.note_text,
            cas_total_diff=args.cas_total_diff,
        )
        print("Preview:")
        print(_pretty(preview))
        if args.apply_change:
            result = bridge.apply_changes(preview, dry_run=False)
            print("Applied:")
            print(_pretty(result))
        return

    ap.print_help()


if __name__ == "__main__":
    # Keep this CLI available when running this file directly
    main_cli()
