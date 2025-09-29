# gui.py
import os
import customtkinter as ctk
import csv
import re
import glob
from datetime import datetime
from tkinter import filedialog, simpledialog, messagebox
import json
import threading
import requests
from typing import Iterable, Set, List, Dict, Optional, Tuple, Any
from processor.crm_bridge import CRMBridge, filter_unchecked_deals

from processor.tabula import pdf_to_excel
from processor.excel_file import ExcelFile
from processor.sheet_finder import find_sheets, find_column_to_sum, extract_claim_amount
from processor.transaction_handler import process_transactions  # optional
from processor.summary_writer import (
    create_summary_sheet,
    create_ta_payment_summary,
    create_currency_type_date_summary,  # pivot-ish summary
    create_payment_summary_eur,  # New function
    create_summary_sheet_eur,  # New function
)
from processor.utils import (
    log_message,
    load_config,
    save_config,
    normalize_deal_id,
    test_apis,  # <-- uses long timeout for IT endpoint
)

from processor.crm_bridge import CRMBridge, ActiveCampaignClient, _today_berlin_str
from processor.utils import log_message, load_config


_RE_RECAP = re.compile(
    r"recap:.*?deposits\s*=\s*([0-9]+(?:\.[0-9]{1,2})?).*?cashouts\s*=\s*([0-9]+(?:\.[0-9]{1,2})?).*?net\s*=\s*([0-9]+(?:\.[0-9]{1,2})?)",
    re.IGNORECASE,
)


def extract_net_from_log_text(text: str) -> Optional[float]:
    """
    Looks for: 'recap: deposits=162004.00, cashouts=132275.71, net=29728.29'
    Returns 29728.29 as float if found.
    """
    m = _RE_RECAP.search(text or "")
    if not m:
        return None
    try:
        return float(m.group(3))
    except Exception:
        return None


# -----------------------------------------------------------------------------
# AC confirmation dialog
# -----------------------------------------------------------------------------


# 2c) A small dialog class
class ACConfirmDialog(ctk.CTkToplevel):
    def __init__(
        self, master, bridge: CRMBridge, deal_id: str, current_title: str = ""
    ):
        super().__init__(master)
        self.title("ActiveCampaign – Confirm Changes")
        self.geometry("580x580")
        self.resizable(True, True)
        self.bridge = bridge
        self.deal_id = str(deal_id)

        # Fetch stages for dropdown
        try:
            ac = self.bridge.ac
            stages = ac.list_stages()  # optionally filter by pipeline if desired
        except Exception as e:
            stages = []
            messagebox.showerror("Stages", f"Could not fetch stages:\n{e}")

        stage_map = {
            f"{s.get('title')} (#{s.get('id')})": str(s.get("id")) for s in stages
        }
        self.stage_labels = sorted(stage_map.keys())
        self.stage_id_by_label = stage_map

        # Pre-fill CAS Total Difference from log recap if available
        # We try to read the master GUI's log textbox if accessible
        try:
            log_text = master.log_text.get("1.0", "end")
        except Exception:
            log_text = ""
        default_net = extract_net_from_log_text(log_text) or 0.0

        # Widgets
        pad = {"padx": 10, "pady": 6}

        self.lbl_deal = ctk.CTkLabel(self, text=f"Deal ID: {deal_id}")
        self.lbl_deal.grid(row=0, column=0, columnspan=2, sticky="w", **pad)

        self.lbl_title = ctk.CTkLabel(self, text=f"Title: {current_title}")
        self.lbl_title.grid(row=1, column=0, columnspan=2, sticky="w", **pad)

        self.stage_label = ctk.CTkLabel(self, text="Move to Stage:")
        self.stage_label.grid(row=2, column=0, sticky="e", **pad)

        self.stage_combo = ctk.CTkComboBox(
            self, values=self.stage_labels or ["(no stages)"]
        )
        self.stage_combo.grid(row=2, column=1, sticky="w", **pad)
        if self.stage_labels:
            self.stage_combo.set(self.stage_labels[0])

        self.chk_add_note_var = ctk.BooleanVar(value=True)
        self.chk_add_note = ctk.CTkCheckBox(
            self, text="Add Checked note", variable=self.chk_add_note_var
        )
        self.chk_add_note.grid(row=3, column=0, columnspan=2, sticky="w", **pad)

        # Default note text using the same helper as crm_bridge
        from processor.crm_bridge import _today_berlin_str  # reuse date format

        default_note = f"Checked ({_today_berlin_str()})"

        self.lbl_note = ctk.CTkLabel(self, text="Note text:")
        self.lbl_note.grid(row=4, column=0, sticky="e", **pad)
        self.note_entry = ctk.CTkEntry(self, width=380)
        self.note_entry.insert(0, default_note)
        self.note_entry.grid(row=4, column=1, sticky="w", **pad)

        self.lbl_cas = ctk.CTkLabel(self, text="CAS Total Difference (ID=75):")
        self.lbl_cas.grid(row=5, column=0, sticky="e", **pad)
        self.cas_entry = ctk.CTkEntry(self, width=180)
        self.cas_entry.insert(0, f"{default_net:.2f}")
        self.cas_entry.grid(row=5, column=1, sticky="w", **pad)

        # Buttons
        self.btn_preview = ctk.CTkButton(self, text="Preview", command=self.on_preview)
        self.btn_preview.grid(row=6, column=0, sticky="e", **pad)
        self.btn_apply = ctk.CTkButton(self, text="Apply", command=self.on_apply)
        self.btn_apply.grid(row=6, column=1, sticky="w", **pad)

        self.txt_result = ctk.CTkTextbox(self, width=540, height=160)
        self.txt_result.grid(row=7, column=0, columnspan=2, padx=10, pady=(4, 10))
        self.txt_result.configure(state="disabled")

    def _collect_changes(self) -> Dict[str, Any]:
        stage_label = (self.stage_combo.get() or "").strip()
        stage_id = (
            self.stage_id_by_label.get(stage_label)
            if stage_label in self.stage_id_by_label
            else None
        )
        add_note = bool(self.chk_add_note_var.get())
        note_text = (self.note_entry.get() or "").strip()
        cas_val_str = (self.cas_entry.get() or "").strip().replace(",", ".")
        cas_val = None
        if cas_val_str:
            try:
                cas_val = float(cas_val_str)
            except Exception:
                messagebox.showerror(
                    "Value error",
                    f"CAS Total Difference is not a number: {cas_val_str}",
                )
                return {}
        preview = self.bridge.preview_changes(
            deal_id=self.deal_id,
            target_stage_id=stage_id,
            add_checked_note=add_note,
            custom_note_text=note_text,
            cas_total_diff=cas_val,
        )
        return preview

    def on_preview(self):
        changes = self._collect_changes()
        if not changes:
            return
        out = json.dumps(changes, ensure_ascii=False, indent=2)
        self.txt_result.configure(state="normal")
        self.txt_result.delete("1.0", "end")
        self.txt_result.insert("end", f"Preview (no changes made yet):\n{out}")
        self.txt_result.configure(state="disabled")

    def on_apply(self):
        changes = self._collect_changes()
        if not changes:
            return
        if messagebox.askyesno(
            "Confirm", "Proceed with these changes to ActiveCampaign?"
        ):
            try:
                result = self.bridge.apply_changes(changes, dry_run=False)
                out = json.dumps(result, ensure_ascii=False, indent=2)
                self.txt_result.configure(state="normal")
                self.txt_result.delete("1.0", "end")
                self.txt_result.insert("end", f"Applied:\n{out}")
                self.txt_result.configure(state="disabled")
                messagebox.showinfo("Done", "Changes applied successfully.")
            except Exception as e:
                messagebox.showerror("ActiveCampaign error", str(e))


class ACBatchConfirmDialog(ctk.CTkToplevel):
    """
    One window to confirm/apply AC changes for MANY deals at once.
    Rows: [Deal ID] [Stage ▼] [✓ Add note] [Note text] [CAS #75]
    Buttons: Preview All / Apply All
    """

    def __init__(
        self,
        master,
        bridge: CRMBridge,
        deals: List[Dict[str, Any]],
        prefill_claim_by_id: Optional[Dict[str, float]] = None,
        prefill_claim_eur_by_id: Optional[Dict[str, float]] = None,
    ):
        super().__init__(master)
        self.title("ActiveCampaign – Batch Confirm Changes")
        self.geometry("900x560")
        self.resizable(True, True)
        self.bridge = bridge
        self.prefill_claim_eur_by_id = prefill_claim_eur_by_id or {}
        self.prefill_claim_by_id = prefill_claim_by_id or {}

        # Fetch stages once
        try:
            ac = self.bridge.ac
            self._stages = ac.list_stages() or []
        except Exception as e:
            self._stages = []
            messagebox.showerror("Stages", f"Could not fetch stages:\n{e}")

        self._stage_names = [f"{s.get('title')} (#{s.get('id')})" for s in self._stages]
        self._stage_label_to_id = {
            f"{s.get('title')} (#{s.get('id')})": str(s.get("id")) for s in self._stages
        }
        # Build UI
        pad = {"padx": 10, "pady": 6}
        header = ctk.CTkLabel(self, text="Review and confirm changes for all deals:")
        header.pack(anchor="w", padx=10, pady=(10, 2))

        self.rows_frame = ctk.CTkScrollableFrame(self, height=370)
        self.rows_frame.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        # Per-deal row widgets will be stored here
        self._rows: List[Dict[str, Any]] = []

        # Add one row per deal
        for d in deals:
            did = str(d.get("deal_id") or d.get("id") or "")
            title = (d.get("title") or d.get("name") or "").strip()
            default_note = f"Checked ({_today_berlin_str()})"
            default_cas = self._default_cas_for(did)

            row = ctk.CTkFrame(self.rows_frame)
            row.pack(fill="x", padx=6, pady=4)

            ctk.CTkLabel(row, text=f"Deal {did}", width=120, anchor="w").grid(
                row=0, column=0, **pad, sticky="w"
            )
            ctk.CTkLabel(row, text=f"Title: {title}", anchor="w").grid(
                row=1, column=0, columnspan=5, **pad, sticky="w"
            )

            # Stage
            ctk.CTkLabel(row, text="Move to Stage:").grid(
                row=0, column=1, **pad, sticky="e"
            )
            stage_combo = ctk.CTkComboBox(
                row, values=self._stage_names or ["(no stages)"], width=240
            )
            if self._stage_names:
                stage_combo.set(self._stage_names[0])
            stage_combo.grid(row=0, column=2, **pad, sticky="w")

            # Add note
            add_note_var = ctk.BooleanVar(value=True)
            add_note_chk = ctk.CTkCheckBox(
                row, text="Add Checked note", variable=add_note_var
            )
            add_note_chk.grid(row=0, column=3, **pad, sticky="w")

            # Note text
            ctk.CTkLabel(row, text="Note text:").grid(
                row=0, column=4, **pad, sticky="e"
            )
            note_entry = ctk.CTkEntry(row, width=320)
            note_entry.insert(0, default_note)
            note_entry.grid(row=0, column=5, **pad, sticky="w")

            # CAS #75
            ctk.CTkLabel(row, text="CAS Total Difference (ID=75):").grid(
                row=0, column=6, **pad, sticky="e"
            )
            cas_entry = ctk.CTkEntry(row, width=120)
            cas_entry.insert(0, f"{default_cas:.2f}")
            cas_entry.grid(row=0, column=7, **pad, sticky="w")

            self._rows.append(
                {
                    "deal_id": did,
                    "stage_combo": stage_combo,
                    "add_note_var": add_note_var,
                    "note_entry": note_entry,
                    "cas_entry": cas_entry,
                }
            )

        # Buttons + output
        btns = ctk.CTkFrame(self)
        btns.pack(fill="x", padx=10, pady=6)
        ctk.CTkButton(btns, text="Preview All", command=self.on_preview_all).pack(
            side="left", padx=4
        )
        ctk.CTkButton(btns, text="Apply All", command=self.on_apply_all).pack(
            side="left", padx=4
        )

        self.txt_result = ctk.CTkTextbox(self, height=140)
        self.txt_result.pack(fill="both", expand=False, padx=10, pady=(2, 10))
        self.txt_result.configure(state="disabled")

    # --- helpers ---

    def _stage_id_from_label(self, label: str) -> Optional[str]:
        """
        Resolve a stage id from a displayed label like 'Stage Name (#123)'.
        Tries exact label first, then a loose startswith match on the title.
        """
        if not label:
            return None
        if hasattr(self, "_stage_label_to_id") and label in self._stage_label_to_id:
            return self._stage_label_to_id[label]
        # loose match: compare just the title prefix, ignore "(#id)"
        try:
            needle = label.split(" (#", 1)[0].strip().lower()
            for full, sid in getattr(self, "_stage_label_to_id", {}).items():
                title = full.split(" (#", 1)[0].strip().lower()
                if title.startswith(needle):
                    return sid
        except Exception:
            pass
        return None

    def _default_cas_for(self, deal_id: str) -> float:
        # Priority 1: remembered claim from processing
        if deal_id in self.prefill_claim_eur_by_id:
            return float(self.prefill_claim_eur_by_id[deal_id])
        if deal_id in self.prefill_claim_by_id:
            return float(self.prefill_claim_by_id[deal_id])
        return 0.0

    def _collect_all_changes(self) -> List[Dict[str, Any]]:
        changes = []
        for r in self._rows:
            did = r["deal_id"]
            stage_label = (r["stage_combo"].get() or "").strip()
            stage_id = self._stage_id_from_label(stage_label)
            add_note = bool(r["add_note_var"].get())
            note_text = (r["note_entry"].get() or "").strip()
            cas_text = (r["cas_entry"].get() or "").strip().replace(",", ".")
            cas_val = None
            if cas_text:
                try:
                    cas_val = float(cas_text)
                except Exception:
                    messagebox.showerror(
                        "Value error",
                        f"Deal {did}: CAS Total Difference is not a number: {cas_text}",
                    )
                    return []
            preview = self.bridge.preview_changes(
                deal_id=did,
                target_stage_id=stage_id,
                add_checked_note=add_note,
                custom_note_text=note_text,
                cas_total_diff=cas_val,
            )
            changes.append(preview)
        return changes

    def _stage_id_by_name(self, label: str) -> Optional[str]:
        return self._stage_id_by_name_loose(label)

    def _stage_id_by_name_loose(self, label: str) -> Optional[str]:
        if label in self._stage_id_by_name:
            return self._stage_id_by_name[label]
        # fallback: try to match by startswith the stage title
        for name, sid in self._stage_id_by_name.items():
            if label and name.lower().startswith(label.lower()):
                return sid
        return None

    def _set_output(self, text: str):
        self.txt_result.configure(state="normal")
        self.txt_result.delete("1.0", "end")
        self.txt_result.insert("end", text)
        self.txt_result.configure(state="disabled")

    # --- UI actions ---
    def on_preview_all(self):
        changes = self._collect_all_changes()
        if not changes:
            return
        out = json.dumps({"preview": changes}, ensure_ascii=False, indent=2)
        self._set_output(out)

    def on_apply_all(self):
        changes = self._collect_all_changes()
        if not changes:
            return
        results = []
        for ch in changes:
            try:
                res = self.bridge.apply_changes(ch, dry_run=False)
            except Exception as e:
                res = {"deal_id": ch.get("deal_id"), "error": str(e)}
            results.append(res)
        out = json.dumps({"applied": results}, ensure_ascii=False, indent=2)
        self._set_output(out)
        messagebox.showinfo("Done", "All changes applied.")


class ExcelProcessorGUI(ctk.CTk):
    def __init__(self):
        self._claim_by_deal_id: Dict[str, float] = {}
        self._claim_eur_by_deal_id: Dict[str, float] = {}  # NEW: final claim in EUR
        super().__init__()

        # ---- Window / theme ----
        self.title("Excel Claim Processor")
        self.geometry("1100x780")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("green")

        # ---- State ----
        self.excel_files: List[ExcelFile] = []
        self._selected_paths: Dict[str, bool] = {}  # file_path -> checkbox state
        self.file_select_vars: List[ctk.BooleanVar] = []

        # visible-only selection helpers
        self._visible_paths: List[str] = []  # paths currently visible after filters
        self._var_by_path: Dict[str, ctk.BooleanVar] = {}  # path -> its checkbox var

        self.config = load_config()
        self.sharepoint_dir = self.config.get("sharepoint_dir")
        self.crm: Optional[CRMBridge] = None  # lazy-init

        self.deals: List[Dict[str, Any]] = []
        self.deal_ids: List[str] = []
        self.selected_deal = ctk.StringVar(value="")
        self.deal_buttons: List[ctk.CTkButton] = []

        # Keep last AC deals we fetched (used to parse opponents from titles)
        self._last_ac_deals: List[Dict[str, Any]] = []

        # Opponents
        self.opponent_by_deal_id: Dict[str, str] = {}
        self.opponent_choices: List[str] = ["(All opponents)"]
        self.opponent_filter_var = ctk.StringVar(value="(All opponents)")
        self.opponent_filter_var.trace_add(
            "write", lambda *_: self.update_file_listbox()
        )

        # date filter (applies to new files you process; stored per-file as well)
        # {"mode":"before/after/between","from":datetime|None,"to":datetime|None}
        self.global_date_filter: Optional[Dict[str, Any]] = None

        self.mode_var = ctk.StringVar(value="checking")

        # free-text filter for queued files
        self.file_filter_var = ctk.StringVar(value="")
        self.file_filter_var.trace_add("write", lambda *_: self.update_file_listbox())

        # progress (shown inside the log panel)
        self._progress_total = 0
        self._progress_count = 0

        self.setup_layout()

    # ---------------- UI LAYOUT ----------------
    def setup_layout(self):
        # ====== Menubar ======
        self._build_menubar()

        # ====== Toolbar ======
        toolbar = ctk.CTkFrame(self)
        toolbar.pack(padx=12, pady=(8, 6), fill="x")

        # Quick actions dropdown
        self.quick_action = ctk.CTkOptionMenu(
            toolbar,
            values=[
                "Quick Actions…",
                "Select Files",
                "AC Deals (filter & build JSON)",
                "Import AC Deal Links…",
                "Load Deals CSV",
                "Add File for Deal",
                "Add All Deals' Files",
                "Move Deals…",
                "Date Filter…",
                "Set SharePoint Folder",
                "Test APIs",
                "Save Selected",
                "Clear All",
                "Remove Selected",
            ],
            command=self._on_quick_action,
        )
        self.quick_action.set("Quick Actions…")
        self.quick_action.pack(side="left", padx=4)

        # Deals dropdown
        self.deals_action = ctk.CTkOptionMenu(
            toolbar,
            values=[
                "Deals…",
                "Load Deals CSV",
                "AC Deals (filter & build JSON)",
                "Import AC Deal Links…",
                "Add File for Deal",
                "Add All Deals' Files",
            ],
            command=self._on_deals_action,
        )
        self.deals_action.set("Deals…")
        self.deals_action.pack(side="left", padx=4)

        # explicit toolbar button to move the DOs (selected deals)
        self.btn_move_deals = ctk.CTkButton(
            toolbar, text="🚚 Move Deals…", command=self.move_deals
        )
        self.btn_move_deals.pack(side="left", padx=6)

        if self.crm is None:
            self.crm = CRMBridge()
        self.btn_ac_dialog = ctk.CTkButton(
            toolbar, text="📝 AC Confirm", command=self.open_ac_batch_for_selected
        )
        self.btn_ac_dialog.pack(side="left", padx=6)

        # File filter (search)
        srch = ctk.CTkFrame(toolbar)
        srch.pack(side="left", padx=12)
        ctk.CTkLabel(srch, text="Filter:").pack(side="left", padx=(0, 6))
        ctk.CTkEntry(
            srch,
            width=220,
            textvariable=self.file_filter_var,
            placeholder_text="type to filter queued files",
        ).pack(side="left")

        # Opponent filter (dropdown + refresh)
        opp_frame = ctk.CTkFrame(toolbar)
        opp_frame.pack(side="left", padx=12)
        ctk.CTkLabel(opp_frame, text="Opponent:").pack(side="left", padx=(0, 6))
        self.opp_menu = ctk.CTkOptionMenu(
            opp_frame,
            values=self.opponent_choices,
            variable=self.opponent_filter_var,
            command=self._on_opponent_choice,
        )
        self.opp_menu.pack(side="left")
        ctk.CTkButton(
            opp_frame, text="↻", width=36, command=self.refresh_opponents
        ).pack(side="left", padx=6)

        # UI scale dropdown
        self.scale_menu = ctk.CTkOptionMenu(
            toolbar,
            values=["90%", "100%", "110%", "125%", "150%"],
            command=self._on_change_scale,
        )
        self.scale_menu.set("100%")
        self.scale_menu.pack(side="right", padx=(6, 2))
        ctk.CTkLabel(toolbar, text="Scale").pack(side="right", padx=(8, 2))

        # Theme toggle
        self.dark_toggle = ctk.CTkSwitch(
            toolbar, text="Dark Mode", command=self.toggle_mode
        )
        self.dark_toggle.pack(side="right", padx=10)

        # Mode radios
        self.mode_check = ctk.CTkRadioButton(
            toolbar, text="Checking", variable=self.mode_var, value="checking"
        )
        self.mode_calc = ctk.CTkRadioButton(
            toolbar, text="Calculation", variable=self.mode_var, value="calculation"
        )
        self.mode_calc.pack(side="right", padx=(10, 6))
        self.mode_check.pack(side="right", padx=(10, 6))

        # ====== Queue panel ======
        queue = ctk.CTkFrame(self)
        queue.pack(padx=12, pady=(4, 6), fill="both", expand=True)

        bar = ctk.CTkFrame(queue)
        bar.pack(fill="x", padx=10, pady=(10, 0))
        ctk.CTkLabel(bar, text="Queued files (tick to select):").pack(
            side="left", padx=(0, 10)
        )
        ctk.CTkButton(
            bar, text="Select All", width=110, command=self.select_all_files
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            bar, text="Select None", width=110, command=self.select_none_files
        ).pack(side="left", padx=4)

        self.file_listbox = ctk.CTkScrollableFrame(queue, height=260)
        self.file_listbox.pack(padx=10, pady=10, fill="both", expand=True)

        # ====== Actions row ======
        actions = ctk.CTkFrame(self)
        actions.pack(pady=(2, 8), padx=12, fill="x")
        self.process_button = ctk.CTkButton(
            actions, text="⚙️ Process All", command=self.process_all
        )
        self.process_button.pack(side="left", padx=6)
        self.process_selected_button = ctk.CTkButton(
            actions, text="✅ Process Selected", command=self.process_selected
        )
        self.process_selected_button.pack(side="left", padx=6)
        self.save_selected_button = ctk.CTkButton(  # NEW
            actions, text="💾 Save Selected", command=self.save_selected_workbooks
        )
        self.save_selected_button.pack(side="left", padx=6)
        self.save_button = ctk.CTkButton(
            actions, text="💾 Save All", command=self.save_all_workbooks
        )
        self.save_button.pack(side="left", padx=6)

        # ====== Log (with progress bar) ======
        log_frame = ctk.CTkFrame(self)
        log_frame.pack(padx=12, pady=(4, 12), fill="both", expand=False)

        self.progress_var = ctk.DoubleVar(value=0.0)
        self.log_progress = ctk.CTkProgressBar(
            log_frame, mode="determinate", variable=self.progress_var
        )
        self.log_progress.set(0.0)
        self._progress_host = log_frame  # keep reference

        self.log_text = ctk.CTkTextbox(log_frame, height=220)
        self.log_text.pack(padx=10, pady=10, fill="both", expand=True)
        self.log_text.configure(state="disabled")

    def _build_menubar(self):
        import tkinter as tk

        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Select Files", command=self.select_files)
        file_menu.add_command(
            label="Save Selected", command=self.save_selected_workbooks
        )  # NEW
        file_menu.add_command(label="Save All", command=self.save_all_workbooks)
        file_menu.add_separator()
        file_menu.add_command(label="Clear All", command=self.clear_files)
        file_menu.add_command(
            label="Remove Selected", command=self.remove_selected_files
        )
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.destroy)
        menubar.add_cascade(label="File", menu=file_menu)

        deals_menu = tk.Menu(menubar, tearoff=0)
        deals_menu.add_command(label="Load Deals CSV", command=self.load_deals_csv)
        deals_menu.add_command(
            label="AC Deals (filter & build JSON)", command=self.show_ac_stage_deals
        )
        deals_menu.add_command(
            label="Import AC Deal Links…", command=self.import_ac_deal_links
        )
        deals_menu.add_command(
            label="Add File for Deal", command=self.select_file_for_selected_deal
        )
        deals_menu.add_command(
            label="Add All Deals' Files", command=self.add_all_deals_files
        )
        deals_menu.add_separator()
        deals_menu.add_command(label="Move Deals…", command=self.move_deals)
        menubar.add_cascade(label="Deals", menu=deals_menu)

        sp_menu = tk.Menu(menubar, tearoff=0)
        sp_menu.add_command(
            label="Set SharePoint Folder", command=self.set_sharepoint_folder
        )
        menubar.add_cascade(label="SharePoint", menu=sp_menu)

        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Date Filter…", command=self.open_date_filter)
        tools_menu.add_command(label="Test APIs", command=self._on_test_apis_click)
        menubar.add_cascade(label="Tools", menu=tools_menu)

        opp_menu = tk.Menu(menubar, tearoff=0)
        opp_menu.add_command(label="Refresh Opponents", command=self.refresh_opponents)
        opp_menu.add_command(
            label="Clear Opponent Filter",
            command=lambda: self._set_opponent_filter("(All opponents)"),
        )
        menubar.add_cascade(label="Opponents", menu=opp_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(
            label="Light Theme", command=lambda: ctk.set_appearance_mode("Light")
        )
        view_menu.add_command(
            label="Dark Theme", command=lambda: ctk.set_appearance_mode("Dark")
        )
        view_menu.add_separator()
        for pct in ["90%", "100%", "110%", "125%", "150%"]:
            view_menu.add_command(
                label=f"Scale {pct}", command=lambda p=pct: self._on_change_scale(p)
            )
        menubar.add_cascade(label="View", menu=view_menu)

        self.configure(menu=menubar)

    # -------- Toolbar handlers --------
    def _on_quick_action(self, choice: str):
        try:
            self.quick_action.set("Quick Actions…")
        except Exception:
            pass
        m = {
            "Select Files": self.select_files,
            "AC Deals (filter & build JSON)": self.show_ac_stage_deals,
            "Import AC Deal Links…": self.import_ac_deal_links,
            "Load Deals CSV": self.load_deals_csv,
            "Add File for Deal": self.select_file_for_selected_deal,
            "Add All Deals' Files": self.add_all_deals_files,
            "Move Deals…": self.move_deals,
            "Date Filter…": self.open_date_filter,
            "Set SharePoint Folder": self.set_sharepoint_folder,
            "Test APIs": self._on_test_apis_click,
            "Save Selected": self.save_selected_workbooks,
            "Clear All": self.clear_files,
            "Remove Selected": self.remove_selected_files,
        }
        fn = m.get(choice)
        if fn:
            fn()

    def _on_deals_action(self, choice: str):
        try:
            self.deals_action.set("Deals…")
        except Exception:
            pass
        m = {
            "Load Deals CSV": self.load_deals_csv,
            "AC Deals (filter & build JSON)": self.show_ac_stage_deals,
            "Import AC Deal Links…": self.import_ac_deal_links,
            "Add File for Deal": self.select_file_for_selected_deal,
            "Add All Deals' Files": self.add_all_deals_files,
        }
        fn = m.get(choice)
        if fn:
            fn()

    def _on_change_scale(self, value: str):
        try:
            pct = int(value.strip().replace("%", ""))
            ctk.set_widget_scaling(pct / 100.0)
        except Exception:
            pass

    # ---------------- Helpers ----------------

    def _compute_net_from_payment_summary_eur(self, file) -> Optional[float]:
        """
        Read 'Payment Summary EUR' (if present) and compute net = Deposits - Redeems in EUR.
        Falls back to None if the sheet/columns aren't there.
        """
        try:
            if (
                not hasattr(file, "wb")
                or "Payment Summary EUR" not in file.wb.sheetnames
            ):
                return None
            sh = file.wb["Payment Summary EUR"]
            header = [c.value for c in sh[1]]
            try:
                amt_idx = header.index("Amount")
                typ_idx = header.index("Type")
            except ValueError:
                # best-effort fallback
                amt_idx, typ_idx = 2, 1
            net_eur = 0.0
            for row in sh.iter_rows(min_row=2, values_only=True):
                typ = row[typ_idx]
                amt = row[amt_idx]
                if not isinstance(amt, (int, float)):
                    continue
                if typ == "Deposit":
                    net_eur += float(amt)
                elif typ == "Redeem":
                    net_eur -= abs(float(amt))
            return round(net_eur, 2)
        except Exception:
            return None

    def toggle_mode(self):
        mode = "dark" if ctk.get_appearance_mode() == "Light" else "light"
        ctk.set_appearance_mode(mode)

    def log(self, message):
        log_message(self.log_text, message)

    # ---- progress helpers (shown inside log panel) ----
    def _progress_show(self, total: int):
        self._progress_total = max(total, 1)
        self._progress_count = 0
        try:
            self.log_progress.pack(padx=10, pady=(10, 0), fill="x")
        except Exception:
            pass
        self.progress_var.set(0.0)
        self.log_progress.set(0.0)
        self.update_idletasks()

    def _progress_step(self, step: int = 1):
        self._progress_count += max(step, 0)
        ratio = min(1.0, float(self._progress_count) / float(self._progress_total or 1))
        self.progress_var.set(ratio)
        self.log_progress.set(ratio)
        self.update_idletasks()

    def _progress_hide(self):
        try:
            self.log_progress.pack_forget()
        except Exception:
            pass
        self.update_idletasks()

    def _set_opponent_filter(self, value: str):
        try:
            self.opponent_filter_var.set(value)
        except Exception:
            pass
        self.update_file_listbox()

    def _on_opponent_choice(self, value: str):
        self.log(f"🎯 Opponent filter: {value}")

    def _is_legacy_excel(self, path: str) -> bool:
        return path.lower().endswith((".xls", ".xlsb"))

    # ---- SharePoint helpers ----
    def _is_under_sharepoint(self, path: str) -> bool:
        try:
            root = (self.sharepoint_dir or "").strip()
            if not (root and os.path.isdir(root)):
                return False
            # normalize & check common path
            root_cp = os.path.commonpath([os.path.abspath(root)])
            p_cp = os.path.commonpath([os.path.abspath(path), root_cp])
            return p_cp == root_cp
        except Exception:
            return False

    def _deal_id_from_sp_path(self, path: str) -> Optional[Tuple[str, Optional[str]]]:
        """
        Walk up from 'path' until a folder matches "^\d{3,10}\s*-\s*(.+)$"
        Returns (deal_id, folder_name) or None.
        """
        try:
            cur = os.path.abspath(path)
            root = os.path.abspath(self.sharepoint_dir or "")
            while True:
                cur, name = os.path.split(cur)
                if not name:
                    break
                # Stop once we exit SharePoint root
                if root and len(cur) < len(root):
                    break
                m = re.match(r"^\s*(\d{3,10})\s*-\s*(.+)$", name)
                if m:
                    return m.group(1), name
                if os.path.abspath(cur) == os.path.abspath(root):
                    break
            return None
        except Exception:
            return None

    def _maybe_enrich_from_sharepoint(self, ef: "ExcelFile"):
        """
        If file is inside SharePoint structure, auto-tag:
          - deal_id from folder name
          - opponent from IT/AC if available (or parse from AC title)
        """
        try:
            if not self._is_under_sharepoint(ef.file_path):
                return

            did_and_name = self._deal_id_from_sp_path(ef.file_path)
            if not did_and_name:
                return
            deal_id, sp_folder_name = did_and_name
            ef.deal_id = deal_id

            # Ensure CRM is ready
            if self.crm is None:
                self.crm = CRMBridge()

            # Get deal from AC to enrich (title → opponent), and cache opponent mapping
            deal = None
            try:
                if hasattr(self, "_fetch_ac_deals_by_ids"):
                    deals = self._fetch_ac_deals_by_ids([deal_id]) or []
                    deal = deals[0] if deals else None
                elif hasattr(self.crm, "get_deal"):
                    deal = self.crm.get_deal(deal_id)
            except Exception:
                deal = None

            opponent = None
            if deal and isinstance(deal, dict):
                # Prefer IT/known mapping (your refresh step may fill this); else parse from AC title
                title = deal.get("title") or deal.get("name") or ""
                opponent = self.opponent_by_deal_id.get(
                    deal_id
                ) or self._opponent_from_title(title)

            if opponent:
                ef.opponent = opponent
                self.opponent_by_deal_id[deal_id] = opponent

            # carry current date filter
            ef.date_filter = (
                dict(self.global_date_filter) if self.global_date_filter else None
            )

            # nice log
            opp_msg = f" — Opponent: {opponent}" if opponent else ""
            self.log(f"🔗 Auto-tagged from SharePoint: Deal {deal_id}{opp_msg}")
        except Exception as e:
            self.log(
                f"⚠️ SharePoint auto-tagging failed for {os.path.basename(ef.file_path)} — {e}"
            )

    # ---- FILE LIST (with visible-only selection) ----
    def update_file_listbox(self):
        # preserve existing selection states
        prev = dict(self._selected_paths)

        # clear UI and our mappings
        for w in self.file_listbox.winfo_children():
            try:
                w.destroy()
            except Exception:
                pass
        self.file_select_vars = []
        self._var_by_path.clear()
        self._visible_paths = []

        # active filters
        needle = (self.file_filter_var.get() or "").strip().lower()
        opp_filter = (self.opponent_filter_var.get() or "").strip()
        opp_active = opp_filter and opp_filter != "(All opponents)"
        opp_filter_low = opp_filter.lower()

        for file in self.excel_files:
            name = os.path.basename(file.file_path)

            # infer opponent if missing and we have mapping
            if getattr(file, "opponent", None) is None:
                did = getattr(file, "deal_id", None)
                if did and did in self.opponent_by_deal_id:
                    file.opponent = self.opponent_by_deal_id.get(did)

            # free-text filter
            if needle and needle not in name.lower():
                continue

            # opponent filter: match stored opponent OR substring in filename
            if opp_active:
                fopp = (getattr(file, "opponent", "") or "").strip().lower()
                if opp_filter_low not in fopp and opp_filter_low not in name.lower():
                    continue

            # visible → build label
            tag = ""
            if getattr(file, "date_filter", None):
                df = file.date_filter
                if df["mode"] == "before":
                    tag = f"  [before {df['to'].date()}]"
                elif df["mode"] == "after":
                    tag = f"  [after {df['from'].date()}]"
                elif df["mode"] == "between":
                    tag = f"  [between {df['from'].date()}–{df['to'].date()}]"

            deal_lbl = (
                f"  [Deal {getattr(file, 'deal_id', '').strip()}]"
                if getattr(file, "deal_id", None)
                else ""
            )
            opp_lbl = (
                f"  [Opponent: {getattr(file, 'opponent', '').strip()}]"
                if getattr(file, "opponent", None)
                else ""
            )

            label_txt = f"{name}{deal_lbl}{opp_lbl}{tag}"

            # checkbox var defaults to previous selection (or True if unseen)
            var = ctk.BooleanVar(value=prev.get(file.file_path, True))

            def _toggle(p=file.file_path, v=var):
                self._selected_paths[p] = v.get()

            cb = ctk.CTkCheckBox(
                self.file_listbox, text=label_txt, variable=var, command=_toggle
            )
            cb.pack(anchor="w", padx=8, pady=3)

            # track visible & mapping
            self.file_select_vars.append(var)
            self._selected_paths[file.file_path] = var.get()
            self._var_by_path[file.file_path] = var
            self._visible_paths.append(file.file_path)

    def select_all_files(self):
        for p in list(self._visible_paths):
            var = self._var_by_path.get(p)
            if var:
                var.set(True)
            self._selected_paths[p] = True

    def select_none_files(self):
        for p in list(self._visible_paths):
            var = self._var_by_path.get(p)
            if var:
                var.set(False)
            self._selected_paths[p] = False

    def _selected_deal_ids_from_visible(self) -> List[str]:
        """
        Return unique deal IDs for currently visible & selected files.
        """
        ids = []
        seen = set()
        for ef in self.excel_files:
            if ef.file_path in self._visible_paths and self._selected_paths.get(
                ef.file_path, False
            ):
                did = getattr(ef, "deal_id", None)
                if did and did not in seen:
                    seen.add(did)
                    ids.append(did)
        return ids

    # ---------------- API test action ----------------
    def _on_test_apis_click(self):
        def _runner():
            try:
                test_apis(self.log)
            except Exception as e:
                self.log(f"❌ API test failed: {e}")

        threading.Thread(target=_runner, daemon=True).start()

    # ---------------- ActiveCampaign x IT: stage flow ----------------
    def show_ac_stage_deals(self):
        """
        Fetch AC deals in saved stage, do ONE IT call for all leads,
        filter OUT cases with 'Checked' note (first non-empty line),
        THEN filter out cases that have 'Checked_' files in SharePoint.

        Export JSON (to current working dir) listing excel_files for kept cases
        and pdf_files (only if no excel files exist).

        Offer converting TA_*.pdf → Excel, log results, optionally queue converted.
        Then ask whether to queue Excel files: all / specific / none.
        """
        try:
            self.btn_ac_stage_deals.configure(state="disabled", text="Loading…")
        except Exception:
            pass

        def _run():
            try:
                if self.crm is None:
                    self.crm = CRMBridge()

                pipeline_id, stage_id, stage_title = self.crm._load_ac_selection()
                if not stage_id and stage_title:
                    stage_id = self.crm.ac.stage_id_by_name(
                        stage_title, pipeline_id=pipeline_id
                    )
                    if not stage_id:
                        raise ValueError(
                            f"Stage '{stage_title}' could not be resolved for pipeline {pipeline_id}."
                        )

                self.log(
                    f"🔗 AC: fetching deals for pipeline={pipeline_id}, stage={stage_id or stage_title}…"
                )
                deals = self.crm._all_deals_in_stage(stage_id, pipeline_id)
                self._last_ac_deals = deals or []
                self.log(f"ℹ️ AC: fetched {len(deals)} deal(s).")

                self._notes_first_then_sharepoint_and_queue(
                    deals, pipeline_id, stage_id, stage_title
                )

            except Exception as e:
                try:
                    messagebox.showerror("ActiveCampaign / IT", str(e))
                except Exception:
                    pass
                self.log(f"❌ AC stage deals retrieval failed: {e}")
            finally:
                try:
                    self.btn_ac_stage_deals.configure(
                        state="normal", text="🔗 AC Deals"
                    )
                except Exception:
                    pass

        threading.Thread(target=_run, daemon=True).start()

    # ---------------- Import AC Deal Links (from emails) ----------------
    def import_ac_deal_links(self):
        """
        Paste email text that contains AC deal links (…/deals/<id>).
        Extract deal IDs, then run notes-first → SharePoint filter, export JSON,
        conversion prompts, and queue prompts—scoped to these IDs.

        NOTE: This path ALSO scans SharePoint for Checked_* files (after the note check).
        """
        ids = self._prompt_paste_deal_links()
        if not ids:
            self.log("ℹ️ No deal IDs detected.")
            return

        def _run():
            try:
                if self.crm is None:
                    self.crm = CRMBridge()

                self.log(
                    f"📥 Imported {len(ids)} deal ID(s) from links: {', '.join(ids[:20])}{'…' if len(ids)>20 else ''}"
                )

                if not (self.sharepoint_dir and os.path.isdir(self.sharepoint_dir)):
                    self.log(
                        "⚠️ SharePoint root not set; Checked_* scan will be skipped unless you set it (Tools ▸ SharePoint)."
                    )

                deals = self._fetch_ac_deals_by_ids(ids)
                if not deals:
                    deals = [{"id": did, "title": f"Deal {did}"} for did in ids]

                self._last_ac_deals = deals

                pipeline_id, stage_id, stage_title = self.crm._load_ac_selection()
                # 🔴 IMPORTANT: this function includes SharePoint Checked_ scan AFTER note filtering
                self._notes_first_then_sharepoint_and_queue(
                    deals, pipeline_id, stage_id, stage_title, restrict_ids=set(ids)
                )

            except Exception as e:
                self.log(f"❌ Import AC Deal Links failed: {e}")

        threading.Thread(target=_run, daemon=True).start()

    def _prompt_paste_deal_links(self) -> List[str]:
        top = ctk.CTkToplevel(self)
        top.title("Import AC Deal Links")
        top.geometry("720x360")
        top.grab_set()
        top.focus_force()

        ctk.CTkLabel(
            top,
            text="Paste email text that contains AC links (…/deals/<id>). We’ll extract the IDs.",
        ).pack(pady=(12, 6))

        box = ctk.CTkTextbox(top)
        box.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        out: Dict[str, Any] = {"ids": []}
        done = ctk.BooleanVar(value=False)

        def on_ok():
            txt = box.get("1.0", "end")
            ids = self._extract_deal_ids_from_text(txt)
            out["ids"] = ids
            done.set(True)
            top.destroy()

        btns = ctk.CTkFrame(top)
        btns.pack(fill="x", padx=12, pady=10)
        ctk.CTkButton(btns, text="Extract & Continue", command=on_ok).pack(
            side="right", padx=6
        )
        self.wait_variable(done)
        return out["ids"]

    def _extract_deal_ids_from_text(self, text: str) -> List[str]:
        ids: List[str] = []
        if text:
            for m in re.finditer(r"/deals/(\d+)(?:\b|/|[\?#])", text):
                ids.append(m.group(1))
            for token in re.findall(r"\b\d{3,8}\b", text):
                if token not in ids:
                    ids.append(token)
        seen, out = set(), []
        for i in ids:
            if i not in seen:
                seen.add(i)
                out.append(i)
        return out

    def _fetch_ac_deals_by_ids(self, ids: List[str]) -> List[Dict[str, Any]]:
        try:
            if hasattr(self.crm, "get_deals_by_ids"):
                deals = self.crm.get_deals_by_ids(ids)
                if isinstance(deals, list):
                    return deals
        except Exception:
            pass

        out = []
        try:
            if hasattr(self.crm, "get_deal"):
                for did in ids:
                    try:
                        d = self.crm.get_deal(did)
                        if isinstance(d, dict):
                            out.append(d)
                    except Exception:
                        continue
        except Exception:
            pass
        return out

    # --------------- Shared pipeline: notes-first -> SP filter -> export -> convert -> queue ---------------
    def _notes_first_then_sharepoint_and_queue(
        self,
        deals: List[Dict[str, Any]],
        pipeline_id: Optional[str],
        stage_id: Optional[str],
        stage_title: Optional[str],
        restrict_ids: Optional[set] = None,
    ):
        # --- ONE IT call ---
        self.log("🧠 IT: fetching ALL leads (single request)…")
        it_leads = self.crm.it.list_leads()
        self.log(f"ℹ️ IT: received {len(it_leads)} lead(s).")

        # Build fast maps from IT results
        notes_by_id: Dict[str, Any] = {}
        opp_by_id: Dict[str, str] = {}
        for lead in it_leads or []:
            lid = str(lead.get("id") or "").strip()
            if not lid:
                continue
            raw_notes = lead.get("notes") or lead.get("note")
            if isinstance(raw_notes, str):
                notes_by_id[lid] = [raw_notes]
            else:
                notes_by_id[lid] = raw_notes or []

            it_opp = self._extract_opponent(lead)
            if it_opp:
                opp_by_id[lid] = it_opp

        # Fallback: parse opponents from AC titles when IT gives none
        for d in deals:
            did = str(d.get("id") or "").strip()
            if not did:
                continue
            if did in opp_by_id:
                continue
            title = d.get("title") or d.get("name") or ""
            parsed = self._opponent_from_title(title)
            if parsed:
                opp_by_id[did] = parsed

        # Update opponent choices in UI
        self._apply_opponent_choices_from_map(opp_by_id)

        # Notes-first filter (multi-line 'Checked')
        sp_root_ok = bool(self.sharepoint_dir and os.path.isdir(self.sharepoint_dir))
        if sp_root_ok:
            self.log(f"📂 SharePoint root: {self.sharepoint_dir}")
        else:
            self.log("⚠️ SharePoint root not set or missing; skipping Checked_ scan.")

        kept: List[Tuple[str, str]] = []  # (deal_id, title)
        dropped_note = dropped_sp = 0

        for d in deals:
            did = str(d.get("id") or "").strip()
            if not did:
                continue
            if restrict_ids and did not in restrict_ids:
                continue
            title = (d.get("title") or d.get("name") or "").strip() or f"(deal {did})"

            ev = self._find_checked_note_evidence(notes_by_id.get(did, []))
            if ev:
                self.log(f"🚫 IT: 'Checked' note for {did} — dropping.")
                self.log(f"    ↳ first line: {ev['first_line']}")
                if ev.get("preview"):
                    self.log(f"    ↳ preview: {ev['preview']}")
                dropped_note += 1
                continue

            if sp_root_ok:
                folder = self._resolve_deal_folder(did)
                if folder:
                    sp_found, sp_samples = self._sharepoint_checked_samples(folder)
                    if sp_found:
                        self.log(
                            f"🚫 SP: Checked_* file(s) present for {did} — dropping. Samples: {', '.join(sp_samples)}"
                        )
                        dropped_sp += 1
                        continue

            kept.append((did, title))

        self.log(
            f"🧹 Filter result — total:{len(deals)} kept:{len(kept)} "
            f"dropped_by_sharepoint:{dropped_sp} dropped_by_notes:{dropped_note}"
        )

        # Build export payload (excel + (pdf if no excel)), with opponent
        case_infos: Dict[str, Dict[str, Any]] = {}
        if kept and sp_root_ok:
            for did, title in kept:
                folder = self._resolve_deal_folder(did)
                excel_files = self._enumerate_excel_files(folder) if folder else []
                pdf_files = []
                if not excel_files:
                    pdf_files = self._enumerate_pdf_files(folder) if folder else []
                    if pdf_files:
                        self.log(
                            f"📄 {did} has no Excel; found {len(pdf_files)} PDF(s)."
                        )
                opponent = opp_by_id.get(did) or self._opponent_from_title(title)
                case_infos[did] = {
                    "deal_id": did,
                    "title": title,
                    "opponent": opponent,
                    "folder": folder,
                    "excel_files": excel_files,
                    "pdf_files": pdf_files if not excel_files else [],
                }
                self.log(
                    f"📦 {did} | {title} — {len(excel_files)} Excel file(s){' — Opponent: '+opponent if opponent else ''}"
                )
        else:
            for did, title in kept:
                opponent = opp_by_id.get(did) or self._opponent_from_title(title)
                case_infos[did] = {
                    "deal_id": did,
                    "title": title,
                    "opponent": opponent,
                    "folder": None,
                    "excel_files": [],
                    "pdf_files": [],
                }

        export_payload = {
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "pipeline_id": pipeline_id,
            "stage_id": stage_id,
            "stage_title": stage_title,
            "sharepoint_root": self.sharepoint_dir if sp_root_ok else None,
            "cases": list(case_infos.values()),
        }
        self._auto_export_json(export_payload)

        # Offer converting TA_*.pdf → Excel
        ta_pdf_map: Dict[str, List[str]] = {}
        for did, info in case_infos.items():
            if info.get("excel_files"):
                continue
            ta_pdfs = [
                p
                for p in (info.get("pdf_files") or [])
                if os.path.basename(p).startswith("TA_")
            ]
            if ta_pdfs:
                ta_pdf_map[did] = ta_pdfs

        converted_paths: List[Tuple[str, str]] = []
        if ta_pdf_map:
            mode, ids_specific = self._prompt_convert_choice(
                allowed_ids=list(ta_pdf_map.keys())
            )
            do_ids: List[str] = []
            if mode == "all":
                do_ids = list(ta_pdf_map.keys())
            elif mode == "specific":
                valid = set(ta_pdf_map.keys())
                do_ids = [x for x in ids_specific if x in valid]
                missing = [x for x in ids_specific if x not in valid]
                if missing:
                    self.log(
                        f"⚠️ Ignored unknown/non-TA-PDF deal IDs: {', '.join(missing)}"
                    )
            else:
                self.log("🛑 Skipped PDF→Excel conversion by user choice.")

            total_conv = sum(len(ta_pdf_map[k]) for k in do_ids)
            if total_conv:
                self._progress_show(total_conv)

            for did in do_ids:
                for pdf_path in ta_pdf_map.get(did, []):
                    try:
                        self.log(
                            f"🧪 Converting TA PDF → Excel: {os.path.basename(pdf_path)} (Deal {did})"
                        )
                        xlsx_path = pdf_to_excel(pdf_path, log=self.log)
                        if xlsx_path and os.path.exists(xlsx_path):
                            converted_paths.append((did, xlsx_path))
                            case_infos[did]["excel_files"].append(xlsx_path)
                            self.log(
                                f"✅ Converted: {os.path.basename(pdf_path)} → {os.path.basename(xlsx_path)}"
                            )
                        else:
                            self.log(
                                f"⚠️ Conversion returned no file for: {os.path.basename(pdf_path)}"
                            )
                    except Exception as e:
                        self.log(
                            f"❌ Conversion failed for {os.path.basename(pdf_path)} — {e}"
                        )
                    finally:
                        self._progress_step(1)

            self._progress_hide()

            if converted_paths:
                if self._prompt_queue_converted(len(converted_paths)):
                    queued = 0
                    seen_paths = {ef.file_path for ef in self.excel_files}
                    self._progress_show(len(converted_paths))
                    for did, x in converted_paths:
                        if x in seen_paths:
                            self._progress_step(1)
                            continue
                        try:
                            ef = ExcelFile(x)
                            ef.deal_id = did
                            ef.opponent = opp_by_id.get(
                                did
                            ) or self._opponent_from_title(case_infos[did]["title"])
                            ef.date_filter = (
                                dict(self.global_date_filter)
                                if self.global_date_filter
                                else None
                            )
                            ef.selected_deposit_sheets = []
                            ef.selected_cashout_sheets = []
                            self.excel_files.append(ef)
                            self._selected_paths[ef.file_path] = True
                            queued += 1
                            seen_paths.add(x)
                        except Exception as e:
                            self.log(
                                f"❌ Failed to queue converted {os.path.basename(x)}: {e}"
                            )
                        finally:
                            self._progress_step(1)
                    self._progress_hide()
                    self.update_file_listbox()
                    self.log(f"📥 Queued {queued} converted file(s) for calculation.")
                else:
                    self.log("🗃️ Converted files NOT queued (per user choice).")

        # Ask whether to queue remaining Excel files
        if not kept:
            self.log("ℹ️ No deals remain to calculate — nothing to fetch.")
            return

        excel_map = {
            did: (info.get("excel_files") or []) for did, info in case_infos.items()
        }
        choice, specific_ids = self._prompt_fetch_choice(
            allowed_ids=[did for did, _ in kept]
        )
        if choice == "none":
            self.log("🙅 Fetch skipped by user.")
            return

        if choice == "all":
            to_fetch_ids = [did for did, _ in kept]
        else:
            valid = set([did for did, _ in kept])
            to_fetch_ids = [x for x in specific_ids if x in valid]
            missing = [x for x in specific_ids if x not in valid]
            if missing:
                self.log(f"⚠️ Ignored unknown deal IDs: {', '.join(missing)}")

        # queue the files for selected deals
        seen_paths = {ef.file_path for ef in self.excel_files}
        candidates = []
        for did in to_fetch_ids:
            for p in excel_map.get(did) or []:
                if p not in seen_paths:
                    candidates.append((did, p))

        self._progress_show(len(candidates) or 1)

        queued = 0
        no_excel: List[str] = []
        for did in to_fetch_ids:
            files = excel_map.get(did) or []
            if not files:
                no_excel.append(did)
                continue
            for p in files:
                if p in seen_paths:
                    self._progress_step(1)
                    continue
                try:

                    if self._is_legacy_excel(p):
                        self.log(
                            f"⚠️ Skipping legacy Excel (not supported): {os.path.basename(p)} - convert to .xlsx"
                        )
                        self._progress_step(1)
                        continue

                    ef = ExcelFile(p)
                    ef.deal_id = did
                    ef.opponent = opp_by_id.get(did) or self._opponent_from_title(
                        case_infos[did]["title"]
                    )
                    ef.date_filter = (
                        dict(self.global_date_filter)
                        if self.global_date_filter
                        else None
                    )
                    ef.selected_deposit_sheets = []
                    ef.selected_cashout_sheets = []
                    self.excel_files.append(ef)
                    self._selected_paths[ef.file_path] = True
                    queued += 1
                    seen_paths.add(p)
                except Exception as e:
                    self.log(f"❌ Failed to queue {os.path.basename(p)}: {e}")
                finally:
                    self._progress_step(1)

        self._progress_hide()
        self.update_file_listbox()
        self.log(f"✅ Queued {queued} Excel file(s) for calculation.")
        if no_excel:
            self.log(
                f"ℹ️ These deals had no Excel files to queue: {', '.join(no_excel[:20])}{'…' if len(no_excel)>20 else ''}"
            )

    # ---------- opponents helpers ----------
    def _extract_opponent(self, lead: Dict[str, Any]) -> Optional[str]:
        for k in (
            "Opponent",
            "opponent",
            "GamblingProvider",
            "Provider",
            "Casino",
            "OpponentName",
        ):
            v = lead.get(k)
            if isinstance(v, str) and v.strip():
                return v.strip()
        return None

    def _opponent_from_title(self, title: Optional[str]) -> Optional[str]:
        if not title:
            return None
        s = str(title)
        if "," not in s:
            return None
        opp = s.split(",")[-1].strip()
        opp = re.sub(r"\s*[-–|:].*$", "", opp).strip()
        return opp or None

    def _apply_opponent_choices_from_map(self, opp_by_id: Dict[str, str]):
        self.opponent_by_deal_id.update(opp_by_id)
        uniq = sorted(
            {v.strip() for v in self.opponent_by_deal_id.values() if v and v.strip()},
            key=lambda s: s.lower(),
        )
        self.opponent_choices = ["(All opponents)"] + uniq
        try:
            self.opp_menu.configure(values=self.opponent_choices)
        except Exception:
            pass

    def refresh_opponents(self):
        def _run():
            try:
                if self.crm is None:
                    self.crm = CRMBridge()
                self.log("🔄 Fetching opponents list from IT…")
                leads = self.crm.it.list_leads()
                opp_map: Dict[str, str] = {}
                for ld in leads or []:
                    lid = str(ld.get("id") or "").strip()
                    if not lid:
                        continue
                    opp = self._extract_opponent(ld)
                    if opp:
                        opp_map[lid] = opp

                if not opp_map and getattr(self, "_last_ac_deals", None):
                    self.log("ℹ️ IT returned no opponents; parsing from AC deal titles…")
                    for d in self._last_ac_deals:
                        did = str(d.get("id") or "").strip()
                        title = d.get("title") or d.get("name") or ""
                        parsed = self._opponent_from_title(title)
                        if did and parsed:
                            opp_map[did] = parsed

                self._apply_opponent_choices_from_map(opp_map)
                self.log(f"✅ Loaded {max(0, len(self.opponent_choices)-1)} opponents.")
            except Exception as e:
                self.log(f"❌ Opponent refresh failed: {e}")

        threading.Thread(target=_run, daemon=True).start()

    # ---------- helpers for 'Checked' detection + SharePoint scan + JSON export ----------
    def _find_checked_note_evidence(self, notes):
        """
        If ANY note has its FIRST non-empty line starting with 'checked' (case-insensitive),
        return {'first_line': <line>, 'preview': <first 200 chars>} else None.
        Handles multi-line notes.
        """

        def _to_text_items(obj):
            if not obj:
                return []
            if isinstance(obj, str):
                return [obj]
            if isinstance(obj, list):
                out = []
                for item in obj:
                    if isinstance(item, str):
                        out.append(item)
                    elif isinstance(item, dict):
                        txt = (
                            item.get("content")
                            or item.get("note")
                            or item.get("text")
                            or item.get("body")
                            or ""
                        )
                        if txt:
                            out.append(txt)
                return out
            if isinstance(obj, dict):
                txt = (
                    obj.get("content")
                    or obj.get("note")
                    or obj.get("text")
                    or obj.get("body")
                    or ""
                )
                return [txt] if txt else []
            return []

        for txt in _to_text_items(notes):
            if not txt:
                continue
            lines = re.split(r"[\r\n]+", str(txt))
            for line in lines:
                if not line.strip():
                    continue
                if re.match(r"(?i)^checked(\b|[^a-z0-9_])", line.strip()):
                    preview = str(txt).strip().replace("\r", " ").replace("\n", " ")
                    if len(preview) > 200:
                        preview = preview[:200] + "…"
                    return {"first_line": line.strip(), "preview": preview}
                break
        return None

    def _sharepoint_checked_samples(self, folder_path: str):
        try:
            samples = []
            for root, _dirs, files in os.walk(folder_path):
                for name in files:
                    if name.startswith("Checked_"):
                        samples.append(name)
                        if len(samples) >= 3:
                            return True, samples
            return (len(samples) > 0), samples
        except Exception as e:
            self.log(f"⚠️ SP scan error in '{folder_path}': {e}")
            return False, []

    def _enumerate_excel_files(self, folder_path: str):
        if not folder_path:
            return []
        exts = (".xlsx", ".xlsm")
        out = []
        try:
            for root, _dirs, files in os.walk(folder_path):
                for name in files:
                    low = name.lower()
                    if (
                        low.endswith(exts)
                        and not name.startswith("~$")
                        and not name.startswith("Checked_")
                    ):
                        out.append(os.path.join(root, name))
        except Exception as e:
            self.log(f"⚠️ Error enumerating Excel files in {folder_path}: {e}")
        out.sort()
        return out

    def _enumerate_pdf_files(self, folder_path: str):
        if not folder_path:
            return []
        out = []
        try:
            for root, _dirs, files in os.walk(folder_path):
                for name in files:
                    low = name.lower()
                    if low.endswith(".pdf") and not name.startswith("~$"):
                        out.append(os.path.join(root, name))
        except Exception as e:
            self.log(f"⚠️ Error enumerating PDF files in {folder_path}: {e}")
        out.sort()
        return out

    def _auto_export_json(self, payload: dict, base_name: str = "to_calculate_files"):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = f"{base_name}_{ts}.json"
        path = os.path.join(os.getcwd(), fname)  # export to code folder (cwd)
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            self.log(f"📤 Exported JSON to: {path}")
            return path
        except Exception as e:
            self.log(f"❌ Export failed: {e}")
        return None

    # ---------- small helper to parse user-entered ID lists ----------
    def _parse_id_list(self, text: str) -> List[str]:
        ids: List[str] = []
        if not text:
            return ids
        for tok in re.findall(r"\b\d{1,10}\b", text):
            if tok not in ids:
                ids.append(tok)
        return ids

    # ---------- prompt: fetch all / specific / none ----------
    def _prompt_fetch_choice(self, allowed_ids: List[str]) -> Tuple[str, List[str]]:
        top = ctk.CTkToplevel(self)
        top.title("Fetch Excel files")
        top.geometry("560x360")
        top.grab_set()
        top.focus_force()

        ctk.CTkLabel(
            top,
            text="Which deals should be queued for calculation?",
            font=("TkDefaultFont", 14),
        ).pack(pady=(12, 6))

        choice_var = ctk.StringVar(value="all")
        rb_frame = ctk.CTkFrame(top)
        rb_frame.pack(fill="x", padx=12, pady=8)
        ctk.CTkRadioButton(
            rb_frame, text="Fetch all", variable=choice_var, value="all"
        ).pack(anchor="w", padx=8, pady=2)
        ctk.CTkRadioButton(
            rb_frame,
            text="Fetch specific (enter IDs below)",
            variable=choice_var,
            value="specific",
        ).pack(anchor="w", padx=8, pady=2)
        ctk.CTkRadioButton(
            rb_frame, text="Fetch none", variable=choice_var, value="none"
        ).pack(anchor="w", padx=8, pady=2)

        ctk.CTkLabel(top, text="Deal IDs (comma/space/newline separated):").pack(
            anchor="w", padx=16, pady=(8, 4)
        )
        ids_box = ctk.CTkTextbox(top, height=120)
        ids_box.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        if allowed_ids:
            hint = ", ".join(allowed_ids[:25]) + ("…" if len(allowed_ids) > 25 else "")
            ctk.CTkLabel(top, text=f"Available: {hint}", text_color="#888").pack(
                anchor="w", padx=16, pady=(0, 6)
            )

        out = {"mode": "none", "ids": []}
        done = ctk.BooleanVar(value=False)

        def on_ok():
            mode = choice_var.get()
            ids = []
            if mode == "specific":
                ids = self._parse_id_list(ids_box.get("1.0", "end"))
            out["mode"] = mode
            out["ids"] = ids
            done.set(True)
            top.destroy()

        btns = ctk.CTkFrame(top)
        btns.pack(fill="x", padx=12, pady=10)
        ctk.CTkButton(btns, text="Continue", command=on_ok).pack(side="right", padx=6)
        ctk.CTkButton(
            btns, text="Cancel", command=lambda: (top.destroy(), done.set(True))
        ).pack(side="right", padx=6)

        self.wait_variable(done)
        return out["mode"], out["ids"]

    # ---------- prompt: convert TA PDFs (all / specific / skip) ----------
    def _prompt_convert_choice(self, allowed_ids: List[str]) -> Tuple[str, List[str]]:
        top = ctk.CTkToplevel(self)
        top.title("Convert TA PDFs to Excel?")
        top.geometry("560x380")
        top.grab_set()
        top.focus_force()

        ctk.CTkLabel(
            top,
            text="TA_*.pdf files were found for some deals with no Excel.\nConvert them to Excel now?",
            font=("TkDefaultFont", 14),
            justify="left",
        ).pack(pady=(12, 6), padx=12, anchor="w")

        choice_var = ctk.StringVar(value="all")
        rb = ctk.CTkFrame(top)
        rb.pack(fill="x", padx=12, pady=8)
        ctk.CTkRadioButton(
            rb, text="Convert all", variable=choice_var, value="all"
        ).pack(anchor="w", padx=8, pady=2)
        ctk.CTkRadioButton(
            rb,
            text="Convert only specific deals (IDs below)",
            variable=choice_var,
            value="specific",
        ).pack(anchor="w", padx=8, pady=2)
        ctk.CTkRadioButton(
            rb, text="Don't convert", variable=choice_var, value="none"
        ).pack(anchor="w", padx=8, pady=2)

        ctk.CTkLabel(top, text="Deal IDs (comma/space/newline separated):").pack(
            anchor="w", padx=16, pady=(8, 4)
        )
        ids_box = ctk.CTkTextbox(top, height=120)
        ids_box.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        if allowed_ids:
            hint = ", ".join(allowed_ids[:25]) + ("…" if len(allowed_ids) > 25 else "")
            ctk.CTkLabel(
                top, text=f"Deals with TA PDFs: {hint}", text_color="#888"
            ).pack(anchor="w", padx=16, pady=(0, 6))

        out = {"mode": "none", "ids": []}
        done = ctk.BooleanVar(value=False)

        def on_ok():
            mode = choice_var.get()
            ids = []
            if mode == "specific":
                ids = self._parse_id_list(ids_box.get("1.0", "end"))
            out["mode"] = mode
            out["ids"] = ids
            done.set(True)
            top.destroy()

        btns = ctk.CTkFrame(top)
        btns.pack(fill="x", padx=12, pady=10)
        ctk.CTkButton(btns, text="Continue", command=on_ok).pack(side="right", padx=6)
        ctk.CTkButton(
            btns, text="Cancel", command=lambda: (top.destroy(), done.set(True))
        ).pack(side="right", padx=6)

        self.wait_variable(done)
        return out["mode"], out["ids"]

    # ---------- prompt: queue converted files? ----------
    def _prompt_queue_converted(self, converted_count: int) -> bool:
        top = ctk.CTkToplevel(self)
        top.title("Queue converted files?")
        top.geometry("460x180")
        top.grab_set()
        top.focus_force()

        ctk.CTkLabel(
            top,
            text=f"{converted_count} file(s) were converted from TA PDFs.\nQueue them now for calculation?",
            font=("TkDefaultFont", 14),
            justify="left",
        ).pack(pady=(18, 6), padx=16, anchor="w")

        res = {"ok": False}
        done = ctk.BooleanVar(value=False)

        def yes():
            res["ok"] = True
            done.set(True)
            top.destroy()

        def no():
            res["ok"] = False
            done.set(True)
            top.destroy()

        btns = ctk.CTkFrame(top)
        btns.pack(fill="x", padx=12, pady=12)
        ctk.CTkButton(btns, text="Yes, queue them", command=yes).pack(
            side="right", padx=6
        )
        ctk.CTkButton(btns, text="No", command=no).pack(side="right", padx=6)

        self.wait_variable(done)
        return res["ok"]

    # ---------------- File selection ----------------
    def select_files(self):
        initial = (
            self.sharepoint_dir
            if self.sharepoint_dir and os.path.isdir(self.sharepoint_dir)
            else None
        )
        paths = filedialog.askopenfilenames(
            title="Select Excel/CSV/PDF files",
            initialdir=initial,
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm *.xls *.xlsb *.xltx *.xltm *.ods *.csv"),
                ("Checked Excel files", "Checked_*.xlsx"),
                ("PDF files", "*.pdf"),
                ("All files", "*.*"),
            ],
        )
        added = 0
        for path in paths:
            try:
                if len(self.excel_files) >= 10:
                    break

                if self._is_legacy_excel(path):
                    self.log(
                        f"⚠️ Skipping legacy Excel file: {os.path.basename(path)} — convert to .xlsx"
                    )
                    continue

                if path.lower().endswith(".pdf"):
                    self.log(f"📄 PDF selected: {os.path.basename(path)} — converting…")
                    xlsx_path = pdf_to_excel(path, log=self.log)
                    excel_file = ExcelFile(xlsx_path)
                else:
                    excel_file = ExcelFile(path)

                # Auto-enrich from SharePoint if applicable
                self._maybe_enrich_from_sharepoint(excel_file)

                if self.global_date_filter:
                    excel_file.date_filter = dict(self.global_date_filter)
                else:
                    excel_file.date_filter = None

                excel_file.selected_deposit_sheets = []
                excel_file.selected_cashout_sheets = []

                self.excel_files.append(excel_file)
                self._selected_paths[excel_file.file_path] = True  # default selected
                added += 1
            except Exception as e:
                self.log(f"❌ Failed to add {os.path.basename(path)} — {e}")

        self.update_file_listbox()
        if added:
            self.log(f"➕ Added {added} file(s).")
        if len(self.excel_files) >= 10:
            self.log("⚠️ File limit (10) reached.")

    def clear_files(self):
        self.excel_files.clear()
        self._selected_paths.clear()
        self.update_file_listbox()
        self.log("Cleared all files.")

    def remove_selected_files(self):
        keep = []
        for ef in self.excel_files:
            if not self._selected_paths.get(ef.file_path, True):
                keep.append(ef)
        self.excel_files = keep
        self._selected_paths = {ef.file_path: True for ef in self.excel_files}
        self.update_file_listbox()
        self.log("Removed selected files.")

    # ---------------- Processing ----------------
    def _process_files(self, files: List[ExcelFile]):
        self._progress_show(len(files))
        processed_dids: List[str] = []  # collect successful deal IDs
        for file in files:
            try:
                from openpyxl import load_workbook

                file.wb = load_workbook(file.file_path, data_only=True)

                filename = os.path.basename(file.file_path)
                did_lbl = (
                    f" (Deal {getattr(file, 'deal_id', '').strip()})"
                    if getattr(file, "deal_id", None)
                    else ""
                )
                self.log(
                    f"📄 Processing {filename}{did_lbl} in {self.mode_var.get()} Mode..."
                )

                # Ask the user for sheet selections before parsing
                self.open_sheet_selector(file)

                mode = self.mode_var.get()
                if mode == "calculation":
                    create_ta_payment_summary(file, self.log)
                    try:
                        create_currency_type_date_summary(file, self.log)
                    except Exception as e:
                        self.log(f"⚠️ Summary (Currency→Type→Year) skipped — {e}")
                    try:
                        create_payment_summary_eur(file, self.log)
                    except Exception as e:
                        self.log(f"⚠️ Payment Summary EUR skipped — {e}")
                    try:
                        create_summary_sheet_eur(file, self.log)
                    except Exception as e:
                        self.log(f"⚠️ Yearly Summary EUR skipped — {e}")
                else:
                    if filename.startswith("IB_"):
                        find_sheets(file)
                        find_column_to_sum(file)
                        file.claim_amount = extract_claim_amount(file)

                        claim_eur = self._compute_net_from_payment_summary_eur(file)
                        if claim_eur is not None and getattr(file, "deal_id", None):
                            self._claim_eur_by_deal_id[str(file.deal_id)] = float(
                                claim_eur
                            )
                    elif getattr(file, "deal_id", None) and isinstance(
                        file.claim_amount, (int, float)
                    ):
                        self._claim_eur_by_deal_id[str(file.deal_id)] = float(
                            file.claim_amount
                        )
                    elif filename.startswith("TA_"):
                        create_ta_payment_summary(file, self.log)

                        if "Payment Summary" not in file.wb.sheetnames:
                            raise ValueError("❌ Payment Summary was not created.")

                        try:
                            create_summary_sheet(file, None, self.log)
                        except Exception as e:
                            self.log(f"⚠️ Yearly Summary skipped — {e}")
                        try:
                            create_currency_type_date_summary(file, self.log)
                        except Exception as e:
                            self.log(f"⚠️ Summary (Currency→Type→Year) skipped — {e}")
                        try:
                            create_payment_summary_eur(file, self.log)
                        except Exception as e:
                            self.log(f"⚠️ Payment Summary EUR skipped — {e}")
                        try:
                            create_summary_sheet_eur(file, self.log)
                        except Exception as e:
                            self.log(f"⚠️ Yearly Summary EUR skipped — {e}")

                        sh = file.wb["Payment Summary"]
                        header = [c.value for c in sh[1]]
                        try:
                            amt_idx = header.index("Amount")
                            typ_idx = header.index("Type")
                        except ValueError:
                            amt_idx, typ_idx = 2, 1

                        net_total = 0.0
                        for row in sh.iter_rows(min_row=2, values_only=True):
                            typ = row[typ_idx]
                            amt = row[amt_idx]
                            if not isinstance(amt, (int, float)):
                                continue
                            if typ == "Deposit":
                                net_total += float(amt)
                            elif typ == "Redeem":
                                net_total -= abs(float(amt))
                        file.claim_amount = round(net_total, 2)

                        claim_eur = self._compute_net_from_payment_summary_eur(file)
                        if claim_eur is not None:
                            if getattr(file, "deal id", None):
                                self._claim_eur_by_deal_id[str(file.deal_id)] = float(
                                    claim_eur
                                )
                        else:
                            if getattr(file, "deal_id", None) and isinstance(
                                file.claim_amount, (int, float)
                            ):
                                self._claim_eur_by_deal_id[str(file.deal_id)] = float(
                                    file.claim_amount
                                )
                        if getattr(file, "deal_id", None) and isinstance(
                            file.claim_amount, (int, float)
                        ):
                            self._claim_by_deal_id[str(file.deal_id)] = float(
                                file.claim_amount
                            )
                        self.log(
                            f"🧾 Final claim value (from Payment Summary): {file.claim_amount}"
                        )
                    else:
                        raise ValueError(
                            "❌ Unsupported file prefix in Checking Mode (use TA_ or IB_)."
                        )

                file.match = True
                file.processed = True
                self.log(f"✅ Finished processing: {filename}")

                if getattr(file, "deal_id", None):
                    processed_dids.append(str(file.deal_id))

            except Exception as e:
                self.log(f"❌ Error in {file.file_path}: {str(e)}")
            finally:
                self._progress_step(1)
        self._progress_hide()

        # After processing, ask to move deals to another stage
        unique_dids = sorted({d for d in processed_dids if d})
        if unique_dids:
            self._prompt_move_deals_post_processing(unique_dids)

        try:
            deals_for_ui = [
                {"deal_id": did, "title": self._title_for_deal(did)}
                for did in unique_dids
            ]
            dlg = ACBatchConfirmDialog(
                self,
                self.crm or CRMBridge(),
                deals=deals_for_ui,
                prefill_claim_eur_by_id=self._claim_eur_by_deal_id,
                prefill_claim_by_id=self._claim_by_deal_id,
            )
            try:
                dlg.grab_set()
                dlg.focus_force()
            except Exception:
                pass
            self.wait_window(dlg)
        except Exception as e:
            self.log(f"❌ Batch AC dialog failed: {e}")

    def process_all(self):
        if not self.excel_files:
            self.log("⚠️ No files selected.")
            return
        self._process_files(self.excel_files)

    def process_selected(self):
        # Only consider items that are currently visible under filters
        visible_selected_paths = [
            p for p in self._visible_paths if self._selected_paths.get(p, False)
        ]
        selected = [
            ef for ef in self.excel_files if ef.file_path in visible_selected_paths
        ]

        if not selected:
            self.log("ℹ️ No files ticked to process in the current view.")
            return

        self._process_files(selected)

    # ---------------- Move Selected Deals (toolbar button) ----------------
    def move_selected_deals(self):
        """Legacy method - now calls the enhanced move_deals method."""
        self.move_deals()

    def move_deals(self):
        """
        Entry point for the 'Move Deals…' button/menu.
        Lets user add IDs or AC links, validates (optional), then opens the stage picker.
        """
        # 1) start with IDs implied by currently visible+selected files
        selected_ids = self._selected_deal_ids_from_visible()

        # 2) ask user for additional IDs or links
        try:
            extra = simpledialog.askstring(
                "Move Deals…",
                "Optional: paste extra Deal IDs or AC links (comma/space/newline separated):",
                parent=self,
            )
        except Exception:
            extra = None

        # 3) parse the extra text to IDs (handles links like .../deals/<id>)
        extra_ids = []
        if extra:
            if hasattr(self, "_extract_deal_ids_from_text"):
                extra_ids = self._extract_deal_ids_from_text(extra)
            else:
                # very small fallback if your extractor isn't present
                extra_ids = list({m for m in re.findall(r"\b\d{1,10}\b", extra)})

        # 4) combine, dedupe (keep order: selected first, then extras)
        all_ids = []
        seen = set()
        for did in (selected_ids or []) + (extra_ids or []):
            s = str(did).strip()
            if s and s not in seen:
                seen.add(s)
                all_ids.append(s)

        if not all_ids:
            self.log("ℹ️ No Deal IDs to move.")
            return

        # 5) (Optional) validate IDs exist in ActiveCampaign (nice-to-have)
        #    If validation fails for some reason, we still proceed.
        try:
            if self.crm is None:
                self.crm = CRMBridge()
            found_deals = self._fetch_ac_deals_by_ids(
                all_ids
            )  # you already have this helper
            found_set = {str(d.get("id")) for d in (found_deals or []) if d.get("id")}
            missing = [d for d in all_ids if d not in found_set]
            if missing:
                self.log(
                    f"⚠️ These IDs were not found in ActiveCampaign: {', '.join(missing[:12])}"
                    + ("…" if len(missing) > 12 else "")
                )
            # Keep all_ids anyway—AC might still accept moves if IDs are valid but not yet fetched
        except Exception as e:
            self.log(f"ℹ️ Skipped AC validation ({e}); continuing.")

        # 6) open your existing stage selection dialog for these IDs
        self._prompt_move_deals_post_processing(all_ids)

    def _get_current_selection(self) -> Optional[Dict[str, Any]]:
        try:
            if getattr(self, "deals", None):
                for d in self.deals:
                    if d.get("_selected"):
                        return {
                            "deal id": str(d.get("id")),
                            "title": d.get("title") or "",
                        }
        except Exception:
            pass

        dlg = ctk.CTkInputDialog(text="Enter ActiveCampaign Deal ID:", title="Deal ID")
        deal_id = dlg.get_input() if dlg else None
        if not deal_id:
            return None
        return {"deal_id": str(deal_id), "title": ""}

    def open_ac_dialog(self):
        sel = self._get_current_selection()
        if not sel:
            messagebox.showwarning("ActiveCampaign", "No deal selected.")
            return
        dlg = ACConfirmDialog(
            self,
            self.crm or CRMBridge(),
            deal_id=sel["deal_id"],
            current_title=sel.get("title", ""),
        )
        try:
            dlg.grab_set()
            dlg.focus_force()
        except Exception:
            pass
        self.wait_window(dlg)

    def open_ac_dialog_for_deals(self, deal_ids: List[str]):
        if self.crm is None:
            self.crm = CRMBridge()
        for did in deal_ids:
            title = self._title_for_deal_id(did)
            dlg = ACConfirmDialog(self, self.crm, deal_id=str(did), current_title=title)
            try:
                dlg.grab_set()
                dlg.focus_force()
            except Exception:
                pass
            self.wait_window(dlg)

    def open_ac_confirm_dialog(self):
        self.open_ac_dialog()

    def open_ac_batch_for_selected(self):

        deal_ids = self._selected_deal_ids_from_visible()
        if not deal_ids:
            messagebox.showwarning("ActiveCampaign", "No deals selected.")
            return

        deals_for_ui = [
            {"deal_id": did, "title": self._title_for_deal(did)} for did in deal_ids
        ]

        dlg = ACBatchConfirmDialog(
            self,
            self.crm or CRMBridge(),
            deals=deals_for_ui,
            prefill_claim_eur_by_id=self._claim_eur_by_deal_id,
            prefill_claim_by_id=self._claim_by_deal_id,
        )
        try:
            dlg.grab_set()
            dlg.focus_force()

        except Exception:
            pass
        self.wait_window(dlg)

    # ---------------- Save ----------------
    def save_all_workbooks(self):
        for file in self.excel_files:
            if file.processed:
                new_path = os.path.join(
                    os.path.dirname(file.file_path),
                    f"Checked_{os.path.basename(file.file_path)}",
                )
                try:
                    file.wb.save(new_path)
                    self.log(f"💾 Saved: {new_path}")
                except Exception as e:
                    self.log(f"❌ Save failed for {new_path}: {e}")

    def save_selected_workbooks(self):
        """
        Save only the currently visible & ticked files that have been processed.
        """
        count = 0
        for ef in self.excel_files:
            if ef.file_path in self._visible_paths and self._selected_paths.get(
                ef.file_path, False
            ):
                if getattr(ef, "processed", False):
                    new_path = os.path.join(
                        os.path.dirname(ef.file_path),
                        f"Checked_{os.path.basename(ef.file_path)}",
                    )
                    try:
                        ef.wb.save(new_path)
                        self.log(f"💾 Saved (selected): {new_path}")
                        count += 1
                    except Exception as e:
                        self.log(f"❌ Save failed for {new_path}: {e}")
                else:
                    self.log(
                        f"⚠️ Skipped (not processed): {os.path.basename(ef.file_path)}"
                    )
        if count == 0:
            self.log("ℹ️ No processed files in selection to save.")

    # ---------------- SharePoint & Deals ----------------
    def set_sharepoint_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.config["sharepoint_dir"] = folder
            save_config(self.config)
            self.sharepoint_dir = folder
            self.log(f"📂 SharePoint folder set to: {folder}")

    def load_deals_csv(self):
        path = filedialog.askopenfilename(
            title="Select Deals CSV",
            filetypes=[("CSV Files", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            self.log("ℹ️ CSV selection cancelled.")
            return

        try:
            with open(path, "r", encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f)
                rows = list(reader)
                headers = [h.strip() for h in (reader.fieldnames or [])]
        except Exception as e:
            self.log(f"❌ Could not read CSV: {e}")
            return

        deal_id_col = self._detect_deal_id_column(headers)
        if not deal_id_col:
            self.log(f"❌ Could not find Deal ID column in headers: {headers}")
            return

        for btn in self.deal_buttons:
            btn.destroy()
        self.deal_buttons.clear()

        self.deals = rows
        self.deal_ids = []

        if not hasattr(self, "deal_frame"):
            self.deal_frame = ctk.CTkScrollableFrame(
                self, width=220, height=120, fg_color="transparent"
            )

        for row in rows:
            did = normalize_deal_id(row.get(deal_id_col))
            if did:
                self.deal_ids.append(did)
                btn = ctk.CTkButton(
                    self.deal_frame,
                    text=did,
                    width=200,
                    height=32,
                    command=lambda d=did: self.select_deal(d),
                )
                btn.pack(pady=2, padx=5, fill="x")
                self.deal_buttons.append(btn)

        if not self.deal_ids:
            btn = ctk.CTkButton(
                self.deal_frame, text="(no deals found)", width=200, state="disabled"
            )
            btn.pack(pady=2, padx=5, fill="x")
            self.deal_buttons.append(btn)
            self.selected_deal.set("")
            self.log("⚠️ No valid Deal IDs found in CSV.")
            return

        self.selected_deal.set(self.deal_ids[0])
        self.log(f"✅ Loaded {len(self.deal_ids)} deals from CSV.")

    def select_file_for_selected_deal(self):
        deal_id = (self.selected_deal.get() or "").strip()
        if not deal_id or deal_id == "(no deals)":
            self.log("⚠️ Select a Deal ID first.")
            return
        if not self.sharepoint_dir or not os.path.isdir(self.sharepoint_dir):
            self.log("⚠️ Set your SharePoint ROOT folder first.")
            return

        folder = self._resolve_deal_folder(deal_id)
        if not folder:
            self.log(
                f"❌ No folder found for Deal ID '{deal_id}' under {self.sharepoint_dir}"
            )
            return

        path = filedialog.askopenfilenames(
            title=f"Select file(s) for Deal {deal_id}",
            initialdir=folder,
            filetypes=[
                ("Excel files (modern)", "*.xlsx *.xlsm"),
                ("All files", "*.*"),
            ],
        )

        added = 0
        for p in path or []:
            if self._is_legacy_excel(p):
                self.log(
                    f"⚠️ Skipping legacy Excel file: {os.path.basename(p)} — convert to .xlsx"
                )
                continue
            ef = ExcelFile(p)
            # Auto-enrich from SharePoint if applicable
            self._maybe_enrich_from_sharepoint(ef)
            ef.deal_id = deal_id  # tag with Deal ID
            ef.opponent = self.opponent_by_deal_id.get(deal_id)  # tag opponent if known
            ef.date_filter = (
                dict(self.global_date_filter) if self.global_date_filter else None
            )
            ef.selected_deposit_sheets = []
            ef.selected_cashout_sheets = []
            self.excel_files.append(ef)
            self._selected_paths[ef.file_path] = True
            added += 1
        self.update_file_listbox()
        if added:
            self.log(f"📎 Added {added} file(s) from Deal {deal_id}.")

    def select_deal(self, deal_id: str):
        self.selected_deal.set(deal_id)
        self.log(f"✅ Selected deal: {deal_id}")

    # ---------------- Helpers (deals) ----------------

    def _title_for_deal(self, deal_id: str) -> str:
        """
        Best-effort title for a deal:
        1) Use cached AC deals from the last fetch (fast, no network)
        2) Use the SharePoint folder name suffix after "<ID> - "
        3) Query AC once (slow path) via existing helper _fetch_ac_deals_by_ids
        """
        did = str(deal_id).strip()
        if not did:
            return ""

        # 1) cached AC deals (from recent stage/list operations)
        try:
            for d in getattr(self, "_last_ac_deals", []) or []:
                if str(d.get("id")) == did:
                    return (d.get("title") or d.get("name") or "").strip()
        except Exception:
            pass

        # 2) SharePoint folder name: "<ID> - <Case Title, Opponent, ...>"
        try:
            folder = self._resolve_deal_folder(did)
            if folder:
                base = os.path.basename(folder)
                if " - " in base:
                    return base.split(" - ", 1)[1].strip()
        except Exception:
            pass

        # 3) AC fetch (single id) via your existing helper
        try:
            deals = self._fetch_ac_deals_by_ids([did])
            if deals:
                return (deals[0].get("title") or deals[0].get("name") or "").strip()
        except Exception:
            pass

        return ""

    def _compute_net_from_payment_summary_eur(self, file) -> Optional[float]:
        """
        Read 'Payment Summary EUR' (if present) and compute net = Deposits - Redeems in EUR.
        Falls back to None if the sheet/columns aren't there.
        """
        try:
            if (
                not hasattr(file, "wb")
                or "Payment Summary EUR" not in file.wb.sheetnames
            ):
                return None
            sh = file.wb["Payment Summary EUR"]
            header = [c.value for c in sh[1]]
            try:
                amt_idx = header.index("Amount")
                typ_idx = header.index("Type")
            except ValueError:
                # best-effort fallback
                amt_idx, typ_idx = 2, 1
            net_eur = 0.0
            for row in sh.iter_rows(min_row=2, values_only=True):
                typ = row[typ_idx]
                amt = row[amt_idx]
                if not isinstance(amt, (int, float)):
                    continue
                if typ == "Deposit":
                    net_eur += float(amt)
                elif typ == "Redeem":
                    net_eur -= abs(float(amt))
            return round(net_eur, 2)
        except Exception:
            return None

    def _detect_deal_id_column(self, headers):
        lowered = [h.lower() for h in headers]
        for key in ["deal id", "deal_id", "dealid", "id", "case id", "caseid"]:
            if key in lowered:
                return headers[lowered.index(key)]
        for i, h in enumerate(lowered):
            if "deal" in h and "id" in h:
                return headers[i]
        for alias in ["no", "ref", "reference", "nummer", "nr"]:
            if alias in lowered:
                return headers[lowered.index(alias)]
        return None

    def _resolve_deal_folder(self, deal_id: str):
        try:
            for entry in os.listdir(self.sharepoint_dir):
                full = os.path.join(self.sharepoint_dir, entry)
                if os.path.isdir(full) and entry.startswith(f"{deal_id} -"):
                    return full
            norm_target = re.sub(r"\W+", "", deal_id).lower()
            candidates = []
            for entry in os.listdir(self.sharepoint_dir):
                full = os.path.join(self.sharepoint_dir, entry)
                if not os.path.isdir(full):
                    continue
                prefix = entry.split(" - ", 1)[0]
                if re.sub(r"\W+", "", prefix).lower() == norm_target:
                    candidates.append(full)
            if len(candidates) == 1:
                return candidates[0]
            if len(candidates) > 1:
                preferred = [
                    c
                    for c in candidates
                    if re.search(r",\s*[^,]+$", os.path.basename(c))
                ]
                return preferred[0] if preferred else candidates[0]
            return None
        except Exception as e:
            self.log(f"⚠️ Folder scan error: {e}")
            return None

    def _pick_files_in_folder(self, folder_path: str):
        if not folder_path:
            return []
        matches = []
        matches += glob.glob(os.path.join(folder_path, "TA_*.xlsx"))
        matches += glob.glob(os.path.join(folder_path, "IB_*.xlsx"))
        if not matches:
            for path in glob.glob(os.path.join(folder_path, "*.xlsx")):
                name = os.path.basename(path)
                if not name.startswith("~$"):
                    matches.append(path)
        seen = set()
        ordered = []
        for p in matches:
            if p not in seen:
                seen.add(p)
                ordered.append(p)
        return ordered

    def add_all_deals_files(self):
        if not self.deal_ids:
            self.log("⚠️ Load a Deals CSV first.")
            return
        if not (self.sharepoint_dir and os.path.isdir(self.sharepoint_dir)):
            self.log("⚠️ Set your SharePoint ROOT folder first.")
            return

        total_added = 0
        missing = []
        for did in self.deal_ids:
            folder = self._resolve_deal_folder(did)
            if not folder:
                missing.append(did)
                continue
            files = self._pick_files_in_folder(folder)
            if not files:
                self.log(f"⚠️ No Excel files found for Deal {did} in {folder}")
                continue
            for p in files:
                ef = ExcelFile(p)
                # Auto-enrich from SharePoint if applicable
                self._maybe_enrich_from_sharepoint(ef)
                ef.deal_id = did
                ef.opponent = self.opponent_by_deal_id.get(did)
                ef.date_filter = (
                    dict(self.global_date_filter) if self.global_date_filter else None
                )
                ef.selected_deposit_sheets = []
                ef.selected_cashout_sheets = []
                self.excel_files.append(ef)
                self._selected_paths[ef.file_path] = True
            total_added += len(files)

        self.update_file_listbox()
        if total_added:
            self.log(f"✅ Queued {total_added} file(s) from all deals.")
        if missing:
            self.log(
                f"❌ No folder found for {len(missing)} deal(s): {', '.join(missing[:8])}{'…' if len(missing)>8 else ''}"
            )

    # ---------------- Post-processing: Move deals to other stages ----------------
    def _prompt_move_deals_post_processing(self, deal_ids: List[str]):
        """
        Ask user whether to move processed deals to new stage(s).
        User can choose a different target stage per deal (or '(Don't move)').
        Also supports adding extra IDs manually with a single stage selection.
        """
        try:
            if self.crm is None:
                self.crm = CRMBridge()

            pipeline_id, current_stage_id, current_stage_title = (
                self.crm._load_ac_selection()
            )
            stages = self._list_stages_for_pipeline(pipeline_id)

            if not stages:
                self.log(
                    "⚠️ Could not load stages for current pipeline; skipping move dialog."
                )
                return

            # Build modal with a dropdown per deal + bulk/extra-IDs section
            top = ctk.CTkToplevel(self)
            top.title("Move selected/processed deals to stage(s)?")
            height = min(680, 220 + 36 * max(1, len(deal_ids)))
            top.geometry(f"700x{height}")
            top.grab_set()
            top.focus_force()

            ctk.CTkLabel(
                top,
                text="Select a target stage for each deal (or pick 'Don't move').\n"
                "You can also add extra Deal IDs below and choose one stage for all extras.",
                justify="left",
            ).pack(pady=(12, 6))

            # Scrollable per-deal rows container
            frm = ctk.CTkScrollableFrame(top, height=300)
            frm.pack(fill="both", expand=True, padx=12, pady=(6, 8))

            # --- Bulk / Extra IDs section (above per-deal rows) ---
            bulk = ctk.CTkFrame(top)
            bulk.pack(fill="x", padx=12, pady=(4, 4))
            ctk.CTkLabel(
                bulk, text="Additional Deal IDs (comma/space/newline separated):"
            ).pack(anchor="w")
            extra_ids_entry = ctk.CTkTextbox(bulk, height=64)
            extra_ids_entry.pack(fill="x", padx=0, pady=(4, 6))

            ctk.CTkLabel(bulk, text="Stage for extra IDs:").pack(
                anchor="w", pady=(4, 2)
            )
            extra_stage_var = ctk.StringVar(value="(Don't move)")
            extra_stage_menu = ctk.CTkOptionMenu(
                bulk, values=[], variable=extra_stage_var
            )
            extra_stage_menu.pack(anchor="w")

            # Stage lists (shared by per-deal dropdowns AND the "extra" stage menu)
            stage_names = ["(Don't move)"] + [
                s.get("title") or s.get("name") or str(s.get("id")) for s in stages
            ]
            name_to_id = {"(Don't move)": None}
            for s in stages:
                nm = s.get("title") or s.get("name") or str(s.get("id"))
                name_to_id[nm] = str(s.get("id"))

            # Apply stage list to the "extra" stage dropdown
            try:
                extra_stage_menu.configure(values=stage_names)
                extra_stage_var.set("(Don't move)")
            except Exception:
                pass

            # Per-deal rows
            var_by_deal: Dict[str, ctk.StringVar] = {}
            for did in deal_ids:
                row = ctk.CTkFrame(frm)
                row.pack(fill="x", padx=6, pady=3)
                ctk.CTkLabel(row, text=f"Deal {did}", width=160, anchor="w").pack(
                    side="left"
                )
                v = ctk.StringVar(value="(Don't move)")
                var_by_deal[did] = v
                ctk.CTkOptionMenu(row, values=stage_names, variable=v).pack(
                    side="left", padx=6
                )

            done = ctk.BooleanVar(value=False)

            def on_ok():
                # Collect per-deal assignments
                assignments = []
                for did, v in var_by_deal.items():
                    stage_name = v.get()
                    stage_id = name_to_id.get(stage_name)
                    if stage_id:
                        assignments.append((did, stage_id, stage_name))

                # Parse extra IDs and apply the chosen extra stage
                extra_text = extra_ids_entry.get("1.0", "end")
                extra_ids = self._parse_id_list(extra_text)
                extra_stage_name = extra_stage_var.get()
                extra_stage_id = name_to_id.get(extra_stage_name)

                if extra_ids and extra_stage_id:
                    for did in extra_ids:
                        assignments.append((did, extra_stage_id, extra_stage_name))
                elif extra_ids and not extra_stage_id:
                    self.log(
                        "ℹ️ Extra IDs provided but stage for extras is '(Don't move)'; they will be ignored."
                    )

                top.destroy()
                if not assignments:
                    self.log("↪️ No moves selected.")
                    done.set(True)
                    return

                # Execute moves with progress
                self.log(f"🚚 Moving {len(assignments)} deal(s) to selected stages…")
                self._progress_show(len(assignments))
                moved = 0
                for did, sid, sname in assignments:
                    try:
                        if self._move_deal_to_stage(did, sid):
                            self.log(f"  ✅ Deal {did} → {sname}")
                            moved += 1
                        else:
                            self.log(f"  ⚠️ Failed to move deal {did} → {sname}")
                    except Exception as e:
                        self.log(f"  ❌ Error moving deal {did} → {sname}: {e}")
                    finally:
                        self._progress_step(1)
                self._progress_hide()
                self.log(f"✅ Stage move finished: {moved}/{len(assignments)} moved.")
                done.set(True)

            # Buttons
            btns = ctk.CTkFrame(top)
            btns.pack(fill="x", padx=12, pady=10)
            ctk.CTkButton(btns, text="Confirm Moves", command=on_ok).pack(
                side="right", padx=6
            )
            ctk.CTkButton(
                btns, text="Cancel", command=lambda: (top.destroy(), done.set(True))
            ).pack(side="right", padx=6)

            self.wait_variable(done)

        except Exception as e:
            self.log(f"❌ Move deals dialog failed: {e}")

    def _list_stages_for_pipeline(
        self, pipeline_id: Optional[str]
    ) -> List[Dict[str, Any]]:
        if not pipeline_id:
            return []
        try:
            if hasattr(self.crm, "list_pipeline_stages"):
                st = self.crm.list_pipeline_stages(pipeline_id)
                if isinstance(st, list):
                    return st
        except Exception:
            pass
        try:
            if hasattr(self.crm, "ac"):
                ac = self.crm.ac
                if hasattr(ac, "list_stages"):
                    st = ac.list_stages(pipeline_id)
                    if isinstance(st, list):
                        return st
                if hasattr(ac, "get_pipeline"):
                    p = ac.get_pipeline(pipeline_id)
                    if isinstance(p, dict):
                        st = p.get("stages") or p.get("data") or []
                        if isinstance(st, list):
                            return st
        except Exception:
            pass
        return []

    def _move_deal_to_stage(self, deal_id: str, stage_id: str) -> bool:
        try:
            if hasattr(self.crm, "move_deal_to_stage"):
                return bool(self.crm.move_deal_to_stage(deal_id, stage_id))
        except Exception:
            pass
        try:
            if hasattr(self.crm, "ac"):
                ac = self.crm.ac
                if hasattr(ac, "move_deal_to_stage"):
                    return bool(ac.move_deal_to_stage(deal_id, stage_id))
                if hasattr(ac, "update_deal"):
                    res = ac.update_deal(deal_id, {"stage": stage_id})
                    return bool(res)
        except Exception:
            pass
        return False

    # ---------------- Sheet selector (BLOCKING) ----------------
    def open_sheet_selector(self, excel_file: ExcelFile):
        top = ctk.CTkToplevel(self)
        top.title(f"Select Sheets — {os.path.basename(excel_file.file_path)}")
        top.geometry("600x520")
        top.grab_set()
        top.focus_force()

        ctk.CTkLabel(
            top,
            text="Tick ALL sheets that contain Deposits and ALL sheets that contain Cashouts.",
        ).pack(pady=10)

        body = ctk.CTkFrame(top)
        body.pack(fill="both", expand=True, padx=10, pady=10)

        left = ctk.CTkFrame(body)
        left.pack(side="left", fill="both", expand=True, padx=6)
        right = ctk.CTkFrame(body)
        right.pack(side="right", fill="both", expand=True, padx=6)

        ctk.CTkLabel(left, text="Deposits").pack(pady=(0, 6))
        ctk.CTkLabel(right, text="Cashouts").pack(pady=(0, 6))

        dep_vars, wd_vars = {}, {}
        for name in excel_file.wb.sheetnames:
            dep_vars[name] = ctk.BooleanVar(
                value=name in getattr(excel_file, "selected_deposit_sheets", [])
            )
            wd_vars[name] = ctk.BooleanVar(
                value=name in getattr(excel_file, "selected_cashout_sheets", [])
            )
            ctk.CTkCheckBox(left, text=name, variable=dep_vars[name]).pack(
                anchor="w", padx=8, pady=2
            )
            ctk.CTkCheckBox(right, text=name, variable=wd_vars[name]).pack(
                anchor="w", padx=8, pady=2
            )

        footer = ctk.CTkFrame(top)
        footer.pack(fill="x", padx=10, pady=10)

        done_flag = ctk.BooleanVar(value=False)

        def clear_all():
            for v in dep_vars.values():
                v.set(False)
            for v in wd_vars.values():
                v.set(False)

        def apply_and_close():
            excel_file.selected_deposit_sheets = [
                n for n, v in dep_vars.items() if v.get()
            ]
            excel_file.selected_cashout_sheets = [
                n for n, v in wd_vars.items() if v.get()
            ]
            excel_file.date_filter = (
                dict(self.global_date_filter) if self.global_date_filter else None
            )
            self.log(f"✔ Deposit sheets: {excel_file.selected_deposit_sheets or '—'}")
            self.log(f"✔ Cashout sheets: {excel_file.selected_cashout_sheets or '—'}")
            if excel_file.date_filter:
                df = excel_file.date_filter
                if df["mode"] == "before":
                    self.log(f"⏱️ Using date filter: BEFORE {df['to'].date()}")
                elif df["mode"] == "after":
                    self.log(f"⏱️ Using date filter: AFTER {df['from'].date()}")
                else:
                    self.log(
                        f"⏱️ Using date filter: BETWEEN {df['from'].date()} and {df['to'].date()}"
                    )
            done_flag.set(True)
            top.destroy()

        def on_close():
            done_flag.set(True)
            top.destroy()

        ctk.CTkButton(footer, text="Clear", command=clear_all).pack(side="left", padx=6)
        ctk.CTkButton(footer, text="Done", command=apply_and_close).pack(
            side="right", padx=6
        )
        top.protocol("WM_DELETE_WINDOW", on_close)

        self.wait_variable(done_flag)

    # ---------------- Date filter modal ----------------
    def open_date_filter(self):
        win = ctk.CTkToplevel(self)
        win.title("Date Filter")
        win.geometry("420x260")
        win.grab_set()

        mode = ctk.StringVar(
            value=(self.global_date_filter or {}).get("mode", "between")
        )
        ctk.CTkRadioButton(win, text="Before…", variable=mode, value="before").pack(
            anchor="w", padx=12, pady=4
        )
        ctk.CTkRadioButton(win, text="After…", variable=mode, value="after").pack(
            anchor="w", padx=12, pady=4
        )
        ctk.CTkRadioButton(win, text="Between…", variable=mode, value="between").pack(
            anchor="w", padx=12, pady=4
        )

        frm = ctk.CTkFrame(win)
        frm.pack(fill="x", padx=12, pady=10)

        ctk.CTkLabel(frm, text="From (YYYY-MM-DD):").grid(
            row=0, column=0, sticky="w", padx=6, pady=6
        )
        ctk.CTkLabel(frm, text="To   (YYYY-MM-DD):").grid(
            row=1, column=0, sticky="w", padx=6, pady=6
        )

        from_entry = ctk.CTkEntry(frm, width=180)
        to_entry = ctk.CTkEntry(frm, width=180)
        from_entry.grid(row=0, column=1, padx=6, pady=6, sticky="w")
        to_entry.grid(row=1, column=1, padx=6, pady=6, sticky="w")

        prev = self.global_date_filter or {}
        if prev.get("from"):
            from_entry.insert(0, str(prev["from"].date()))
        if prev.get("to"):
            to_entry.insert(0, str(prev["to"].date()))

        def parse_date(s):
            s = (s or "").strip()
            if not s:
                return None
            for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d"):
                try:
                    return datetime.strptime(s, fmt)
                except Exception:
                    continue
            return None

        def apply():
            m = mode.get()
            d_from = parse_date(from_entry.get())
            d_to = parse_date(to_entry.get())

            if m == "before":
                if not d_to:
                    messagebox.showerror(
                        "Date Filter", "Please enter the 'To' date for 'Before'."
                    )
                    return
                self.global_date_filter = {"mode": "before", "from": None, "to": d_to}
                self.log(f"⏱️ Date filter set: BEFORE {d_to.date()}")
            elif m == "after":
                if not d_from:
                    messagebox.showerror(
                        "Date Filter", "Please enter the 'From' date for 'After'."
                    )
                    return
                self.global_date_filter = {"mode": "after", "from": d_from, "to": None}
                self.log(f"⏱️ Date filter set: AFTER {d_from.date()}")
            else:
                if not (d_from and d_to):
                    messagebox.showerror(
                        "Date Filter", "Please enter both From and To for 'Between'."
                    )
                    return
                if d_from > d_to:
                    d_from, d_to = d_to, d_from
                self.global_date_filter = {
                    "mode": "between",
                    "from": d_from,
                    "to": d_to,
                }
                self.log(
                    f"⏱️ Date filter set: BETWEEN {d_from.date()} and {d_to.date()}"
                )

            for ef in self.excel_files:
                ef.date_filter = dict(self.global_date_filter)
            self.update_file_listbox()
            win.destroy()

        btns = ctk.CTkFrame(win)
        btns.pack(fill="x", padx=12, pady=8)
        ctk.CTkButton(btns, text="Apply", command=apply).pack(side="right", padx=6)

        def clear():
            self.global_date_filter = None
            for ef in self.excel_files:
                ef.date_filter = None
            self.update_file_listbox()
            self.log("⏱️ Date filter cleared.")
            win.destroy()

        ctk.CTkButton(btns, text="Clear", command=clear).pack(side="left", padx=6)


def launch_gui():
    app = ExcelProcessorGUI()
    app.mainloop()


if __name__ == "__main__":
    launch_gui()
