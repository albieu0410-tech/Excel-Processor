"""
Normalization helpers for transaction parsing.

Exports (backward-compatible):
- normalize_tx_type(text) -> "Deposit" | "Redeem" | None
- normalize_status(s) -> str
- is_effective_deposit(status) -> bool
- is_effective_withdraw(status) -> bool
- is_rejected(status) -> bool
- parse_amount(x) -> float | None
- DATE_FORMATS (tuple of strptime formats)
- parse_date(date_str) -> datetime | None

Additional helpers (non-breaking):
- is_final_success(status) -> bool
- is_bad_or_pending(status) -> bool
- normalize_currency(x, default="EUR") -> str
"""

from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from typing import Optional

# ----------------------------
# Canonical keywords (broad/multilingual)
# ----------------------------

DEPOSIT_KEYWORDS = {
    # EN
    "deposit",
    "topup",
    "top-up",
    "credit",
    "funding",
    "purchase",
    "cardverification",
    "addfunds",
    "addfunds",
    "transferin",
    "transfer in",
    # DE
    "einzahlung",
    "gutschrift",
    # FR
    "dépôt",
    "depot",
    "versement",
    # ES/PT/IT
    "depósito",
    "deposito",
    "ingreso",
    "credito",
    "crédito",
    "carte",
    # NL
    "storting",
    "bijschrijving",
    # PL
    "wpłata",
    "wplata",
    # RU
    "ввод",
    "зачисление",
    # TR
    "yatirma",
    "yatırım",
    "yukleme",
    "yükleme",
    # Short-hands
    "dep",
}

WITHDRAW_KEYWORDS = {
    # EN
    "withdraw",
    "withdrawal",
    "cashout",
    "cash out",
    "redeem",
    "payout",
    "pay out",
    "refund",
    "chargeback",
    "reversal",
    "reversed",
    "transferout",
    "transfer out",
    "debit",
    "debitmemo",
    # DE
    "auszahlung",
    "abbuchung",
    "rueckbuchung",
    "rückbuchung",
    # FR
    "retrait",
    "remboursement",
    "contrepassation",
    # ES/PT/IT
    "retiro",
    "saque",
    "prelievo",
    "reembolso",
    "rimborso",
    "estorno",
    "estornado",
    # NL
    "opname",
    "terugbetaling",
    # PL
    "wypłata",
    "wyplata",
    "obciążenie",
    # RU
    "вывод",
    "spisanie",
    # TR
    "cekim",
    "çekim",
    # Short-hands
    "red",
    "wd",
    "wdr",
}

# ----------------------------
# Status normalization
# ----------------------------


def _norm(s: Optional[str]) -> str:
    """Lowercase, trim, collapse to [a-z0-9] only with typo-tolerant cleanup."""
    if s is None:
        return ""

    # Convert to string and normalize unicode
    s = str(s)
    import unicodedata

    s = unicodedata.normalize("NFKC", s)  # fold weird unicode
    s = s.strip().lower()

    # collapse runs of separators
    s = re.sub(r"[\s\-_]+", "_", s)

    # --- Fix common OCR/typo variants of "closed" ---
    # cIosed / c1osed / ciosed / cl0sed → closed
    s = re.sub(r"c[i1]osed", "closed", s)  # cIosed, c1osed
    s = s.replace("ciosed", "closed")  # ciosed
    s = s.replace("cl0sed", "closed")  # cl0sed

    # Fix specific whole tokens seen in the wild
    s = s.replace("rdm_ciosed", "rdm_closed")
    s = s.replace("rdm_c1osed", "rdm_closed")
    s = s.replace("rdm_cIosed", "rdm_closed")

    # Final cleanup: collapse to [a-z0-9] only
    return re.sub(r"[^a-z0-9]+", "", s)


TYPE_MAP = {
    # money in
    "deposit": "deposit",
    "credit": "deposit",
    "topup": "deposit",
    "top-up": "deposit",
    "add funds": "deposit",
    "funds added": "deposit",
    # money out
    "redeem": "cashout",
    "withdraw": "cashout",
    "withdrawal": "cashout",
    "payout": "cashout",
    "cashout": "cashout",
    "cash out": "cashout",
}

STATUS_MAP = {
    # accepted (final)
    "approved": "approved",
    "complete": "approved",
    "completed": "approved",
    "success": "approved",
    "succeeded": "approved",
    # final but dropped
    "rejected": "rejected",
    "reject": "rejected",
    "declined": "rejected",
    "failed": "rejected",
    "error": "rejected",
    "cancelled": "cancelled",
    "canceled": "cancelled",
    "cancel": "cancelled",
    # not final
    "pending": "pending",
    "processing": "processing",
    "in progress": "processing",
    "inprogress": "processing",
    "waiting for offline payment": "waitingforofflinepayment",
    "waitingforofflinepayment": "waitingforofflinepayment",
}

# Buckets with aliases
STATUS_ALIASES = {
    "accepted_deposit": {
        "dep_settled",
        "settled",
        "approved",
        "completed",
        "success",
        "ok",
        "done",
        # some exports label deposits as "posted"/"processed"
        "posted",
        "processed",
        "confirmed",
        "accepted",
        "captured",
        "closed",
    },
    "accepted_withdraw": {
        "pay_closed",
        "paid",
        "processed",
        "completed",
        "success",
        "ok",
        "done",
        "approved",
        "closed",
        "settled",
        "posted",
        "confirmed",
        "accepted",
        "captured",
        "payout_paid",
        "payout_closed",
        "rdm_closed",
    },
    "rejected": {
        "rejected",
        "declined",
        "failed",
        "canceled",
        "cancelled",
        "void",
        "chargeback",
        "reversed",
        "refund",
        "refunded",
        "expired",
        "denied",
        "error",
        "notprocessed",
        "insufficientfunds",
    },
    "pending": {
        "pending",
        "processing",
        "inprogress",
        "in_progress",
        "inprocess",
        "awaiting",
        "open",
        "onhold",
        "authorized",
        "authorised",
        "underreview",
        "inreview",
        "awaitpayment",
    },
    "waiting": {"waiting"},
    "unknown": {"unknown"},
}

# One-letter compact status used by some layouts
STATUS_COMPACT = {
    "a": "approved",
    "c": "canceled",
    "p": "pending",
    "w": "waiting",
}

OK_TOKENS_GENERIC = {
    "approved",
    "completed",
    "settled",
    "closed",
    "success",
    "successful",
    "paid",
    "posted",
    "processed",
    "ok",
    "done",
    "confirmed",
    "accepted",
    "captured",
}
BAD_TOKENS_GENERIC = {
    "rejected",
    "declined",
    "failed",
    "canceled",
    "cancelled",
    "void",
    "chargeback",
    "reversed",
    "refunded",
    "expired",
    "denied",
    "error",
    "notprocessed",
    "insufficientfunds",
}
PENDING_TOKENS_GENERIC = {
    "pending",
    "processing",
    "inprogress",
    "inprocess",
    "in_review",
    "inreview",
    "underreview",
    "onhold",
    "authorized",
    "authorised",
    "awaiting",
    "open",
    "awaitpayment",
}


def normalize_status(s: str) -> str:
    """
    Normalize raw status to a stable bucket-like label.
    - Recognizes compact A/C/P/W.
    - Maps a wide array of aliases to accepted_deposit / accepted_withdraw / rejected / pending / waiting.
    - Returns 'unknown' when empty.
    - Enhanced with OCR/typo healing for withdrawal statuses.
    """
    if s is None:
        return "unknown"

    # Enhanced preprocessing with OCR/typo healing
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKC", s).strip().lower()

    # normalize separators but KEEP underscores
    s = re.sub(r"[\s\-]+", "_", s)

    # --- OCR / typo healing for "closed" family ---
    # cIosed / c1osed / cl0sed / ciosed → closed
    s = re.sub(r"c[i1]osed", "closed", s)  # cIosed, c1osed
    s = s.replace("cl0sed", "closed")  # cl0sed
    s = s.replace("ciosed", "closed")  # ciosed

    # Build an alnum-only shadow to catch lost underscores etc.
    s_alnum = re.sub(r"[^a-z0-9]", "", s)

    # If we see rdm+closed even without separators, force canonical 'rdm_closed'
    if "rdmclosed" in s_alnum:
        s = "rdm_closed"

    # Other canonicalizations (optional but handy)
    if "payclosed" in s_alnum:
        s = "pay_closed"
    if "payoutclosed" in s_alnum:
        s = "payout_closed"
    if "payoutpaid" in s_alnum:
        s = "payout_paid"

    # Now continue with existing bucket mapping logic
    raw = s
    if not raw:
        return "unknown"

    n = _norm(raw)
    if not n:
        return "unknown"

    # Compact letters first
    if n in STATUS_COMPACT:
        # Keep canonical words consistent with other buckets
        comp = STATUS_COMPACT[n]
        if comp == "approved":
            # we won't know deposit vs withdraw here; callers decide contextually
            return "approved"
        return comp

    # Direct bucket aliases
    for bucket, values in STATUS_ALIASES.items():
        if n in values:
            return bucket

    # Generic fallbacks
    if n in OK_TOKENS_GENERIC:
        return "approved"
    if n in BAD_TOKENS_GENERIC:
        return "rejected"
    if n in PENDING_TOKENS_GENERIC:
        return "pending"

    return n


def is_effective_deposit(status: str) -> bool:
    """
    Historical behavior: treat statuses mapped to 'accepted_deposit' as successful deposits.
    Also accept generic 'approved' where direction is inferred elsewhere.
    """
    ns = normalize_status(status)
    return ns in {"accepted_deposit", "approved"}


def is_effective_withdraw(status: str) -> bool:
    """
    Historical behavior: treat statuses mapped to 'accepted_withdraw' as successful withdrawals.
    Also accept generic 'approved' where direction is inferred elsewhere.
    """
    ns = normalize_status(status)
    return ns in {"accepted_withdraw", "approved"}


def is_rejected(status: str) -> bool:
    return normalize_status(status) == "rejected"


# New convenience gates (optional to use)
def is_final_success(status: str) -> bool:
    ns = normalize_status(status)
    return (
        ns in {"accepted_deposit", "accepted_withdraw"}
        or ns in OK_TOKENS_GENERIC
        or ns == "approved"
    )


def is_bad_or_pending(status: str) -> bool:
    ns = normalize_status(status)
    return (
        ns in {"rejected", "pending", "waiting"}
        or ns in BAD_TOKENS_GENERIC
        or ns in PENDING_TOKENS_GENERIC
    )


# ----------------------------
# Transaction type classification
# ----------------------------


def normalize_tx_type(text: Optional[str]):
    """
    Classify a transaction string into 'Deposit' / 'Redeem' / None.
    Backward-compatible name required by summary_writer and others.
    """
    tkn = _norm(text or "")
    if not tkn:
        return None

    # quick shorthands
    if tkn in {"dep", "depo"}:
        return "Deposit"
    if tkn in {"red", "wd", "wdr"}:
        return "Redeem"

    # keyword containment
    for kw in WITHDRAW_KEYWORDS:
        if kw in tkn:
            return "Redeem"
    for kw in DEPOSIT_KEYWORDS:
        if kw in tkn:
            return "Deposit"

    # simple verb forms
    if tkn in {"credited", "credit"}:
        return "Deposit"
    if tkn in {"debited", "debit"}:
        return "Redeem"

    return None


# ----------------------------
# Amount parsing
# ----------------------------

CURRENCY_SYMBOLS_RE = re.compile(r"[A-Za-z$€£¥₽₺₹₩₪₿₫฿₴₦₲៛₡₱ƒ₵₸₭₮₼₨₥₤₢₣₠₧₯₳₰]")


def parse_amount(x) -> Optional[float]:
    """
    Robust amount parser:
      - Handles currency symbols/letters.
      - Handles thousands and decimal separators (., , , spaces, apostrophes).
      - Handles parentheses negatives and leading/trailing minus.
      - Returns float (negative if truly negative), or None if not parsable.
    """
    if x is None:
        return None

    if isinstance(x, (int, float)):
        try:
            return float(x)
        except Exception:
            return None

    s = str(x).strip()
    if not s:
        return None

    s = s.replace("\u00a0", " ")

    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1].strip()
    if s.endswith("-"):
        neg, s = True, s[:-1].strip()
    if s.startswith("-"):
        neg, s = True, s[1:].strip()

    # strip currency letters/symbols
    s = CURRENCY_SYMBOLS_RE.sub("", s)

    # Remove spaces and apostrophes used as thousands
    s = s.replace(" ", "").replace("'", "")

    # Decide separators
    if "," in s and "." in s:
        # last separator is decimal
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        if s.count(",") == 1:
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    # else dots or none → fine

    try:
        val = float(s)
    except Exception:
        return None

    return -val if neg else val


# ----------------------------
# Currency normalization (optional utility)
# ----------------------------

CURRENCY_MAP = {
    "eur": "EUR",
    "€": "EUR",
    "usd": "USD",
    "$": "USD",
    "gbp": "GBP",
    "£": "GBP",
    "inr": "INR",
    "₹": "INR",
    "uah": "UAH",
    "pln": "PLN",
    "rub": "RUB",
    "₽": "RUB",
    "try": "TRY",
    "₺": "TRY",
    "brl": "BRL",
    "cad": "CAD",
    "aud": "AUD",
    "chf": "CHF",
    "sek": "SEK",
    "nok": "NOK",
    "dkk": "DKK",
}


def normalize_currency(x: Optional[str], default: str = "EUR") -> str:
    if x is None:
        return default
    s = str(x).strip()
    if not s:
        return default
    s_up = s.upper()
    if len(s_up) == 3 and s_up.isalpha():
        return s_up
    n = _norm(s)
    if n in CURRENCY_MAP:
        return CURRENCY_MAP[n]
    m = re.search(r"[A-Za-z]{3}", s_up)
    if m:
        return m.group(0)
    return default


# ----------------------------
# Date parsing (your original formats preserved)
# ----------------------------

DATE_FORMATS = (
    # US month/day/year with AM/PM (covers e.g. "1/15/2020 1:21:27 PM")
    "%m/%d/%Y %I:%M:%S %p",
    "%m/%d/%Y %I:%M %p",
    "%m-%d-%Y %I:%M:%S %p",
    "%m-%d-%Y %I:%M %p",
    "%m.%d.%Y %I:%M:%S %p",
    "%m.%d.%Y %I:%M %p",
    # US month/day/year 24-hour
    "%m/%d/%Y %H:%M:%S",
    "%m/%d/%Y %H:%M",
    "%m/%d/%Y",
    "%m-%d-%Y %H:%M:%S",
    "%m-%d-%Y %H:%M",
    "%m-%d-%Y",
    "%m.%d.%Y %H:%M:%S",
    "%m.%d.%Y %H:%M",
    "%m.%d.%Y",
    # ISO / year-first
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y/%m/%d %H:%M:%S",
    "%Y/%m/%d %H:%M",
    "%Y.%m.%d %H:%M:%S",
    "%Y.%m.%d %H:%M",
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%Y.%m.%d",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%dT%H:%M",
    # Day-first common
    "%d-%m-%Y %H:%M:%S",
    "%d-%m-%Y %H:%M",
    "%d-%m-%Y",
    "%d.%m.%Y %H:%M:%S",
    "%d.%m.%Y %H:%M",
    "%d.%m.%Y",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y",
    # With month short name (e.g. "12/Jan/2020 05:30:00")
    "%d/%b/%Y %H:%M:%S",
    "%d/%b/%Y %H:%M",
    "%d/%b/%Y",
)


def parse_date(date_str: str):
    s = (date_str or "").strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return None
