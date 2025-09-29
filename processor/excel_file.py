import os
import csv
import openpyxl
from datetime import datetime
from typing import Optional, Tuple, List


class ExcelFile:
    """
    Lightweight wrapper around an openpyxl Workbook with a few conveniences
    used across the app (flags, selected sheets, date filters, etc.).

    - Accepts .xlsx/.xlsm directly.
    - Accepts .csv and converts it to an in‑memory workbook (single sheet).
    """

    # ----- Date filter modes -----
    DATE_ALL = "all"  # no filter
    DATE_BEFORE = "before"  # <= date_a
    DATE_AFTER = "after"  # >= date_a
    DATE_BETWEEN = "between"  # date_a <= d <= date_b

    def __init__(self, file_path: str):
        self.file_path = file_path

        # Will be set by parsers / finders
        self.calculation_sheet: Optional[str] = None
        self.payment_sheet: Optional[str] = None
        self.column_to_sum: Optional[str] = None
        self.claim_amount: Optional[float] = None

        # UI / flow flags
        self.processed: bool = False
        self.match: bool = False

        # User selections (from the dual list modal)
        self.selected_deposit_sheets: List[str] = []
        self.selected_cashout_sheets: List[str] = []

        # Optional date filter (applied by parsing functions that opt‑in)
        # mode = one of DATE_* constants; date_a/date_b are naive datetimes (no tz)
        self.date_filter_mode: str = ExcelFile.DATE_ALL
        self.date_a: Optional[datetime] = None
        self.date_b: Optional[datetime] = None

        # Load workbook (CSV is converted to a single‑sheet workbook)
        self.wb = self._load_any(file_path)

    # ---------- public helpers ----------

    def set_date_filter_before(self, dt: datetime):
        self.date_filter_mode = ExcelFile.DATE_BEFORE
        self.date_a = dt
        self.date_b = None

    def set_date_filter_after(self, dt: datetime):
        self.date_filter_mode = ExcelFile.DATE_AFTER
        self.date_a = dt
        self.date_b = None

    def set_date_filter_between(self, dt_start: datetime, dt_end: datetime):
        # Normalize ordering just in case
        if dt_start and dt_end and dt_start > dt_end:
            dt_start, dt_end = dt_end, dt_start
        self.date_filter_mode = ExcelFile.DATE_BETWEEN
        self.date_a = dt_start
        self.date_b = dt_end

    def clear_date_filter(self):
        self.date_filter_mode = ExcelFile.DATE_ALL
        self.date_a = None
        self.date_b = None

    def passes_date_filter(self, dt: Optional[datetime]) -> bool:
        """Utility used by parsers to honor the user's date limiter."""
        if not isinstance(dt, datetime) or self.date_filter_mode == ExcelFile.DATE_ALL:
            return True
        if self.date_filter_mode == ExcelFile.DATE_BEFORE:
            return dt <= (self.date_a or dt)
        if self.date_filter_mode == ExcelFile.DATE_AFTER:
            return dt >= (self.date_a or dt)
        if self.date_filter_mode == ExcelFile.DATE_BETWEEN:
            a = self.date_a or dt
            b = self.date_b or dt
            return a <= dt <= b
        return True

    def set_active_sheet(self, name: str):
        """Best‑effort: move the requested sheet to the front (used by some flows)."""
        if name in self.wb.sheetnames:
            ws = self.wb[name]
            # Move to index 0 if not already there
            if self.wb._sheets and self.wb._sheets[0] is not ws:
                self.wb._sheets.remove(ws)
                self.wb._sheets.insert(0, ws)

    # ---------- internal loaders ----------

    def _load_any(self, path: str):
        low = (path or "").lower()
        if low.endswith(".csv"):
            return self._csv_to_workbook(path)
        # .xlsx / .xlsm / etc. — open normally (data_only for speed when reading formulas)
        return openpyxl.load_workbook(path, data_only=True)

    def _csv_to_workbook(self, path: str):
        """
        Convert a CSV into an openpyxl Workbook with a single sheet named 'All'.
        - No dialect sniffing that would be too costly; use Python's csv with universal newline handling.
        - Keeps values as strings; numeric parsing will happen in the parser anyway.
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "All"

        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            for row in reader:
                ws.append(row)
        return wb
