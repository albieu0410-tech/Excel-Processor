# processor/tabula.py
import os
import tempfile
import pandas as pd


def _try_tabula(path, log):
    try:
        import tabula
    except Exception as e:
        if log:
            log(f"ℹ️ tabula not available: {e}")
        return []
    try:
        return tabula.read_pdf(
            path, pages="all", multiple_tables=True, lattice=True
        ) or tabula.read_pdf(path, pages="all", multiple_tables=True, stream=True)
    except Exception as e:
        if log:
            log(f"⚠️ tabula failed: {e}")
        return []


def _try_camelot(path, log):
    try:
        import camelot
    except Exception as e:
        if log:
            log(f"ℹ️ camelot not available: {e}")
        return []
    try:
        tables = camelot.read_pdf(path, pages="all", flavor="lattice")
        if tables.n == 0:
            tables = camelot.read_pdf(path, pages="all", flavor="stream")
        return [t.df for t in tables] if tables.n > 0 else []
    except Exception as e:
        if log:
            log(f"⚠️ camelot failed: {e}")
        return []


def pdf_to_excel(pdf_path: str, log=None) -> str:
    """
    Convert PDF tables into a single-sheet Excel.
    Headers are always col_0...col_N + page.
    """
    if log:
        log(
            f"🧩 Converting PDF → Excel (default headers): {os.path.basename(pdf_path)}"
        )

    # extract
    tables = _try_tabula(pdf_path, log)
    if not tables:
        if log:
            log("🔁 Falling back to camelot…")
        tables = _try_camelot(pdf_path, log)
    if not tables:
        raise RuntimeError("No tables found in PDF.")

    # normalize rows
    all_rows = []
    for page_num, t in enumerate(tables, start=1):
        df = t if isinstance(t, pd.DataFrame) else pd.DataFrame(t)
        for row in df.itertuples(index=False, name=None):
            vals = [str(v).strip() if v is not None else "" for v in row]
            vals.append(page_num)
            all_rows.append(vals)

    # build headers dynamically (max columns across all rows)
    max_cols = max(len(r) for r in all_rows)
    headers = [f"col_{i}" for i in range(max_cols - 1)] + ["page"]

    # pad rows
    for r in all_rows:
        if len(r) < max_cols:
            r.extend([""] * (max_cols - len(r)))

    # write excel
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    prefix = (
        "TA_"
        if base.upper().startswith("TA_")
        else ("IB_" if base.upper().startswith("IB_") else "TA_")
    )
    out_dir = os.path.dirname(pdf_path) or tempfile.gettempdir()
    out_path = os.path.join(out_dir, f"{prefix}{base}.xlsx")

    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "All"
    ws.append(headers)
    for row in all_rows:
        ws.append(row)
    wb.save(out_path)

    if log:
        log(
            f"✅ Created Excel: {os.path.basename(out_path)} with {len(all_rows)} rows, {len(headers)} cols"
        )
    return out_path
