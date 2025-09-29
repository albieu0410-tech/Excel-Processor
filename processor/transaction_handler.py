from __future__ import annotations
from dataclasses import dataclass
from datetime import datetime
from typing import Callable, Dict, Iterable, List, Optional, Tuple, Any
import re

# ---- local helpers -----------------------------------------------------------


def _ts() -> str:
    return datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")


def _clean_str(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return re.sub(r"\s+", " ", s)


def _as_float(v: Any) -> Optional[float]:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = _clean_str(v)
    if s == "":
        return None
    s = re.sub(r"[^\d,.\-]", "", s)
    if s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    if s.count(".") == 1 and s.count(",") >= 1:
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def _parse_date(*candidates: Any) -> Optional[datetime]:
    fmts = [
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%d/%b/%Y",
        "%d/%b/%Y %H:%M:%S",
        "%d/%m/%Y",
        "%d/%m/%Y %H:%M:%S",
        "%m/%d/%Y",
        "%m/%d/%Y %H:%M:%S",
    ]
    for v in candidates:
        s = _clean_str(v)
        if not s:
            continue
        for f in fmts:
            try:
                return datetime.strptime(s, f)
            except ValueError:
                pass
    return None


def _norm_type(s: str) -> str:
    t = _clean_str(s).lower()
    if t in {"deposit", "credit", "topup", "top-up", "add funds", "funds added"}:
        return "deposit"
    if t in {"withdraw", "withdrawal", "payout", "redeem", "cashout", "cash out"}:
        return "cashout"
    if "redeem" in t:
        return "cashout"
    if "withdraw" in t or "payout" in t:
        return "cashout"
    if "deposit" in t:
        return "deposit"
    return t or "unknown"


def _norm_status(s: str) -> str:
    v = _clean_str(s).lower()
    v2 = re.sub(r"[^a-z]", "", v)
    if v2 in {"approved", "complete", "completed", "success", "succeeded"}:
        return "approved"
    if v2 in {"reject", "rejected", "declined", "failed", "error"}:
        return "rejected"
    if v2 in {"cancel", "cancelled", "canceled"}:
        return "cancelled"
    if "waitingforofflinepayment" in v2 or ("waiting" in v2 and "offline" in v2):
        return "waitingforofflinepayment"
    if v2 in {"pending", "processing", "inprogress"}:
        return "pending" if v2 == "pending" else "processing"
    return v2 or "unknown"


def _currency(s: Any) -> str:
    c = _clean_str(s).upper()
    if re.fullmatch(r"[A-Z]{3}", c):
        return c
    if c in {"$", "USD"}:
        return "USD"
    if c in {"€", "EUR"}:
        return "EUR"
    if c in {"£", "GBP"}:
        return "GBP"
    return c or "USD"


def _shorten(s: str, limit: int = 60) -> str:
    if len(s) <= limit:
        return s
    return s[: limit - 1] + "…"


# -----------------------------------------------------------------------------


@dataclass
class Txn:
    row_index: int
    when: datetime
    amount: float
    sign: int  # +1 deposit, -1 cashout (0 if dropped)
    currency: str
    label: str  # "Deposit" or "Cashout" (or "Unknown")
    status_raw: str
    type_raw: str
    status_norm: str
    type_norm: str
    gateway_raw: str
    descr_raw: str
    drcr_raw: str
    sign_source: str
    accepted: bool
    reason: Optional[str] = None


FINAL_ACCEPT_STATUSES = {"approved"}
FINAL_DROP_STATUSES = {"rejected", "cancelled"}
NOT_FINAL_STATUSES = {"pending", "processing", "waitingforofflinepayment", "unknown"}

COL_ALIASES = {
    "type": {"type", "transaction type", "kind"},
    "status": {"status", "state"},
    "currency": {"currency", "curr"},
    "amount": {"amount", "amt", "value", "sum"},
    "gateway": {
        "paysystem",
        "gateway",
        "payment system",
        "payment provider",
        "payment method",
        "processor",
        "provider",
    },
    "description": {
        "description",
        "details",
        "reference",
        "memo",
        "notes",
        "narrative",
        "comment",
        "info",
    },
    "drcr": {
        "dr/cr",
        "drcr",
        "debit/credit",
        "direction",
        "txn direction",
        "transaction direction",
        "debitcredit",
        "type description",
    },
    "date": {
        "date",
        "transaction date",
        "approved date",
        "request date",
        "created",
        "processed at",
        "timestamp",
    },
}


def _get_first(d: Dict[str, Any], keys: Iterable[str]) -> Any:
    for k in keys:
        if k in d and d[k] not in (None, ""):
            return d[k]
    lower = {k.lower(): k for k in d.keys()}
    for k in keys:
        if k in lower:
            real = lower[k]
            if d[real] not in (None, ""):
                return d[real]
    return None


def interpret_row(i1_based: int, row: Dict[str, Any]) -> Txn:
    type_raw = _get_first(row, COL_ALIASES["type"]) or ""
    status_raw = _get_first(row, COL_ALIASES["status"]) or ""
    currency = _currency(_get_first(row, COL_ALIASES["currency"]))
    amount = _as_float(_get_first(row, COL_ALIASES["amount"])) or 0.0
    gateway_raw = _get_first(row, COL_ALIASES["gateway"]) or ""
    descr_raw = _get_first(row, COL_ALIASES["description"]) or ""
    drcr_raw = _get_first(row, COL_ALIASES["drcr"]) or ""

    dt = (
        _parse_date(
            _get_first(row, {"approved date", "transaction date"}),
            _get_first(row, {"request date"}),
            _get_first(row, COL_ALIASES["date"]),
        )
        or datetime.now()
    )

    t_norm = _norm_type(type_raw)
    s_norm = _norm_status(status_raw)
    gateway_clean = _clean_str(gateway_raw)
    descr_clean = _clean_str(descr_raw)
    drcr_clean = _clean_str(drcr_raw)

    if t_norm == "cashout":
        sign = -1
        label = "Cashout"
        sign_source = "type_norm='cashout'"
    elif t_norm == "deposit":
        sign = +1
        label = "Deposit"
        sign_source = "type_norm='deposit'"
    else:
        return Txn(
            row_index=i1_based,
            when=dt,
            amount=amount,
            sign=0,
            currency=currency,
            label="Unknown",
            status_raw=_clean_str(status_raw),
            type_raw=_clean_str(type_raw),
            status_norm=s_norm,
            type_norm=t_norm,
            gateway_raw=gateway_clean,
            descr_raw=descr_clean,
            drcr_raw=drcr_clean,
            sign_source="type_norm='unknown'",
            accepted=False,
            reason=f"unknown type '{_clean_str(type_raw)}'",
        )

    if s_norm in FINAL_ACCEPT_STATUSES:
        accepted = True
        reason = None
    elif s_norm in FINAL_DROP_STATUSES:
        accepted = False
        reason = f"bad status '{_clean_str(status_raw)}'"
    elif s_norm in NOT_FINAL_STATUSES:
        accepted = False
        reason = f"not final '{_clean_str(status_raw)}' (normalized='{s_norm}')"
    else:
        accepted = False
        reason = f"unknown status '{_clean_str(status_raw)}'"

    return Txn(
        row_index=i1_based,
        when=dt,
        amount=amount,
        sign=sign if accepted else 0,
        currency=currency,
        label=label,
        status_raw=_clean_str(status_raw),
        type_raw=_clean_str(type_raw),
        status_norm=s_norm,
        type_norm=t_norm,
        gateway_raw=gateway_clean,
        descr_raw=descr_clean,
        drcr_raw=drcr_clean,
        sign_source=sign_source,
        accepted=accepted,
        reason=reason,
    )


def process_rows(
    rows: Iterable[Dict[str, Any]],
    log: Optional[Callable[[str], None]] = None,
) -> Tuple[List[Txn], float, float]:
    txns: List[Txn] = []
    dep_total = 0.0
    cash_total = 0.0

    def _emit(s: str) -> None:
        if log:
            try:
                log(s)
            except Exception:
                print(s)
        else:
            print(s)

    for idx, raw in enumerate(rows, start=1):
        t = interpret_row(idx, raw)
        dt_str = t.when.strftime("%Y-%m-%d")
        context_parts: List[str] = []
        if t.gateway_raw:
            context_parts.append(f"gateway='{_shorten(t.gateway_raw)}'")
        if t.drcr_raw:
            context_parts.append(f"drcr='{_shorten(t.drcr_raw, 20)}'")
        if t.descr_raw:
            context_parts.append(f"descr='{_shorten(t.descr_raw)}'")
        context_suffix = f" | hints: {', '.join(context_parts)}" if context_parts else ""
        _emit(
            f"{_ts()} 🧾 row {t.row_index}: type_raw='{t.type_raw or '-'}' → "
            f"'{t.type_norm}' | status_raw='{t.status_raw or '-'}' → '{t.status_norm}' | "
            f"amount={t.amount:.2f} {t.currency} | sign_source={t.sign_source}{context_suffix}"
        )
        if t.accepted:
            if t.label == "Deposit":
                dep_total += t.amount
                sign = "+1"
            else:
                cash_total += t.amount
                sign = "-1"

            accept_context = []
            if t.drcr_raw:
                accept_context.append(f"drcr='{_shorten(t.drcr_raw, 20)}'")
            if t.gateway_raw:
                accept_context.append(f"gateway='{_shorten(t.gateway_raw)}'")
            if t.descr_raw:
                accept_context.append(f"descr='{_shorten(t.descr_raw)}'")
            accept_suffix = f"; ctx: {', '.join(accept_context)}" if accept_context else ""
            _emit(
                f"{_ts()} 🔎 row {t.row_index}: ✅ accept → dt={dt_str} "
                f"amt={t.amount:.2f} sign={sign} curr={t.currency} "
                f"label={t.label} (type_norm='{t.type_norm}', status_norm='{t.status_norm}', "
                f"source={t.sign_source}){accept_suffix}"
            )
            txns.append(t)
        else:
            reason = t.reason or "dropped"
            drop_context = []
            if t.drcr_raw:
                drop_context.append(f"drcr='{_shorten(t.drcr_raw, 20)}'")
            if t.gateway_raw:
                drop_context.append(f"gateway='{_shorten(t.gateway_raw)}'")
            if t.descr_raw:
                drop_context.append(f"descr='{_shorten(t.descr_raw)}'")
            drop_suffix = f"; ctx: {', '.join(drop_context)}" if drop_context else ""
            _emit(
                f"{_ts()} 🔎 row {t.row_index}: ⏭️ drop — {reason} "
                f"(type_norm='{t.type_norm}', status_norm='{t.status_norm}', "
                f"source={t.sign_source}){drop_suffix}"
            )

    return txns, dep_total, cash_total


def summarize(
    rows: Iterable[Dict[str, Any]],
    log: Optional[Callable[[str], None]] = None,
) -> Dict[str, Any]:
    txns, dep_total, cash_total = process_rows(rows, log=log)
    net = dep_total - cash_total
    if log:
        log(
            f"{_ts()} 🔎 recap: deposits={dep_total:.2f}, cashouts={cash_total:.2f}, net={net:.2f}"
        )
        log(
            f"{_ts()} TA logic: Deposits = {dep_total}, Cashouts = {cash_total}, Net = {net}"
        )
    return {
        "transactions": txns,
        "deposits": dep_total,
        "cashouts": cash_total,
        "net": net,
    }


# ---- BACKWARD-COMPAT SHIM ----------------------------------------------------
# Your gui.py imports this name.
def process_transactions(
    rows: Iterable[Dict[str, Any]],
    log: Optional[Callable[[str], None]] = None,
):
    """
    Backward-compatible wrapper for older callers.

    Returns a 4-tuple to keep legacy code happy:
       (transactions_list, deposits_total, cashouts_total, net_total)

    Newer code can call summarize(...) instead.
    """
    result = summarize(rows, log=log)
    return (
        result["transactions"],
        result["deposits"],
        result["cashouts"],
        result["net"],
    )
