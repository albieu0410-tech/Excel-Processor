from processor.utils import beautify_with_xlwings
from processor.excel_file import ExcelFile
from processor.summary_writer import (
    create_ta_payment_summary,
    create_summary_sheet,
    create_payment_summary_eur,  # EUR-normalized Payment Summary
    create_summary_sheet_eur,  # NEW: Summary EUR from Payment Summary EUR
)
from gui import launch_gui


def run_processing(
    file_path, mode="calculation", forced_sheet=None, apply_styling=True
):
    log = make_logger_ui()  # your GUI logger hook
    excel = ExcelFile(file_path, log)

    if mode == "checking":
        perform_checking(excel, log)  # your existing checking path
        excel.save()
        if apply_styling:
            beautify_with_xlwings(file_path, log)
        return

    # CALCULATION MODE
    if forced_sheet:
        try:
            excel.set_active_sheet(forced_sheet)
            log(f"📑 Using sheet: {forced_sheet}")
        except Exception as e:
            log(f"⚠️ Could not set active sheet '{forced_sheet}': {e}")

    # Build the base “Payment Summary” and the grouped “Summary”
    create_ta_payment_summary(excel, log)
    create_summary_sheet(excel, None, log)

    # Build the EUR-normalized copy (USD→EUR via ECB average over the observed span per year)
    try:
        create_payment_summary_eur(excel, log)
    except Exception as e:
        log(f"⚠️ Payment Summary EUR skipped — {e}")

    # NEW: Build the EUR summary (Currency → Type → Year) from “Payment Summary EUR”
    try:
        create_summary_sheet_eur(excel, log)
    except Exception as e:
        log(f"⚠️ Summary EUR skipped — {e}")

    # Save workbook
    excel.save()  # openpyxl save

    # Optional native Excel polish (colors, formats, freeze panes)
    if apply_styling:
        beautify_with_xlwings(file_path, log)


if __name__ == "__main__":
    launch_gui()
