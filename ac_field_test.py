# ac_field_test.py
"""
Test utility for ActiveCampaign Deal custom fields and notes.

Functions:
  1) List all Deal custom fields:
       python ac_field_test.py --list-fields

  2) Set a specific Deal custom field value:
       python ac_field_test.py --deal-id 123456 --field-id 42 --value "Checked"

  3) List notes for a Deal:
       python ac_field_test.py --list-deal-notes --deal-id 123456

  4) Create a note on a Deal (NEW):
       python ac_field_test.py --add-deal-note "Your note text" --deal-id 123456

  5) Delete a note by Note ID (NEW):
       python ac_field_test.py --delete-note --note-id 98765

Env required:
  AC_API_URL (or AC_API_BASE_URL), AC_API_KEY
"""

from __future__ import annotations
import argparse
import json
import sys
from typing import Any, Dict, Optional, List

try:
    from processor.utils import ac_root, _ac_headers, ac_get
except Exception:
    from utils import ac_root, _ac_headers, ac_get  # type: ignore

import requests


def _pretty(js: Any) -> str:
    try:
        return json.dumps(js, ensure_ascii=False, indent=2)
    except Exception:
        return str(js)


# ------------------------- Deal custom fields ------------------------- #


def list_deal_custom_fields() -> List[Dict[str, Any]]:
    """Fetch all deal custom fields."""
    base = ac_root().rstrip("/")
    url = f"{base}/dealCustomFieldMeta"
    resp = requests.get(url, headers=_ac_headers(), timeout=(10, 30))
    if not resp.ok:
        raise RuntimeError(f"Failed to fetch fields: {resp.status_code} {resp.text}")
    js = resp.json() or {}
    return js.get("dealCustomFieldMeta", []) or []


def _find_existing_dcv(deal_id: str, field_id: str) -> Optional[Dict[str, Any]]:
    """Find existing dealCustomFieldDatum for (deal_id, field_id)."""
    params = {
        "filters[dealId]": deal_id,
        "filters[customFieldId]": field_id,
        "limit": 100,
    }
    r = ac_get("dealCustomFieldData", params=params, timeout_s=30)
    if not r.ok:
        params = {
            "filters[dealid]": deal_id,
            "filters[customfieldid]": field_id,
            "limit": 100,
        }
        r = ac_get("dealCustomFieldData", params=params, timeout_s=30)
        if not r.ok:
            return None
    js = r.json() or {}
    items = js.get("dealCustomFieldData") or js.get("dealCustomFieldDatum") or []
    for item in items:
        if str(item.get("dealId")) == str(deal_id) and str(
            item.get("customFieldId")
        ) == str(field_id):
            return item
    return None


def set_deal_custom_field(deal_id: str, field_id: str, value: str) -> Dict[str, Any]:
    """Upsert a Deal custom field value."""
    base = ac_root().rstrip("/")
    headers = _ac_headers()

    existing = _find_existing_dcv(deal_id, field_id)

    if existing and existing.get("id"):
        dcv_id = str(existing["id"])
        payload = {"dealCustomFieldDatum": {"fieldValue": value}}
        url = f"{base}/dealCustomFieldData/{dcv_id}"
        resp = requests.patch(
            url, headers=headers, data=json.dumps(payload), timeout=(10, 30)
        )
        out = {"action": "update", "status_code": resp.status_code, "ok": resp.ok}
        try:
            out["response"] = resp.json()
        except Exception:
            out["response"] = resp.text
        return out

    payload = {
        "dealCustomFieldDatum": {
            "dealId": str(deal_id),
            "customFieldId": str(field_id),
            "fieldValue": value,
        }
    }
    url = f"{base}/dealCustomFieldData"
    resp = requests.post(
        url, headers=headers, data=json.dumps(payload), timeout=(10, 30)
    )
    out = {"action": "create", "status_code": resp.status_code, "ok": resp.ok}
    try:
        out["response"] = resp.json()
    except Exception:
        out["response"] = resp.text
    return out


# ------------------------------ Deal notes ------------------------------ #


def list_deal_notes(deal_id: str) -> List[Dict[str, Any]]:
    """
    Fetch all notes for a given Deal via GET /deals/{id}/notes (pagination supported if present).
    """
    base = ac_root().rstrip("/")
    url = f"{base}/deals/{deal_id}/notes"
    headers = _ac_headers()

    notes: List[Dict[str, Any]] = []

    while True:
        resp = requests.get(url, headers=headers, timeout=(10, 30))
        if not resp.ok:
            raise RuntimeError(
                f"Failed to fetch deal notes: {resp.status_code} {resp.text}"
            )

        js = resp.json() or {}
        batch = js.get("notes") or []
        notes.extend(batch)

        meta = js.get("meta") or {}
        links = meta.get("links") or {}
        nxt = links.get("next")
        if not nxt:
            break
        url = f"{ac_root().rstrip('/')}{nxt}" if nxt.startswith("/") else nxt

    return notes


def create_deal_note(deal_id: str, text: str) -> Dict[str, Any]:
    """
    Create a note on a specific Deal.
    Prefers POST /deals/{id}/notes. If your account lacks this route, fallback to POST /notes with reltype/relid.
    """
    base = ac_root().rstrip("/")
    headers = _ac_headers()

    # Try deal-scoped endpoint first
    url = f"{base}/deals/{deal_id}/notes"
    payload = {"note": {"note": text}}
    resp = requests.post(
        url, headers=headers, data=json.dumps(payload), timeout=(10, 30)
    )
    if resp.ok:
        try:
            return {
                "ok": True,
                "status_code": resp.status_code,
                "response": resp.json(),
            }
        except Exception:
            return {"ok": True, "status_code": resp.status_code, "response": resp.text}

    # Fallback: global /notes with relation info
    url = f"{base}/notes"
    payload = {"note": {"note": text, "relid": int(deal_id), "reltype": "Deal"}}
    resp2 = requests.post(
        url, headers=headers, data=json.dumps(payload), timeout=(10, 30)
    )
    out = {"ok": resp2.ok, "status_code": resp2.status_code}
    try:
        out["response"] = resp2.json()
    except Exception:
        out["response"] = resp2.text
    return out


def delete_note(note_id: str) -> Dict[str, Any]:
    """
    Delete a note by its Note ID via DELETE /notes/{id}.
    """
    base = ac_root().rstrip("/")
    headers = _ac_headers()
    url = f"{base}/notes/{note_id}"
    resp = requests.delete(url, headers=headers, timeout=(10, 30))
    out = {"ok": resp.ok, "status_code": resp.status_code}
    try:
        out["response"] = resp.json()
    except Exception:
        out["response"] = resp.text
    return out


# --------------------------------- CLI --------------------------------- #


def main():
    ap = argparse.ArgumentParser(
        description="Test ActiveCampaign Deal custom fields and notes."
    )
    ap.add_argument(
        "--list-fields", action="store_true", help="List all Deal custom fields"
    )
    ap.add_argument("--deal-id", help="Deal ID")
    ap.add_argument("--field-id", help="Custom Field ID")
    ap.add_argument("--value", help="Value to set")
    ap.add_argument(
        "--list-deal-notes", action="store_true", help="List notes for a Deal"
    )
    ap.add_argument(
        "--add-deal-note", help="Create a note on a Deal (provide the note text)"
    )
    ap.add_argument(
        "--delete-note", action="store_true", help="Delete a note by --note-id"
    )
    ap.add_argument("--note-id", help="Note ID to delete")

    args = ap.parse_args()

    # 1) List all deal custom fields
    if args.list_fields:
        fields = list_deal_custom_fields()
        print("Available Deal Custom Fields:")
        for f in fields:
            fid = f.get("id")
            name = f.get("fieldLabel") or f.get("title")
            ftype = f.get("fieldType")
            print(f"  ID={fid} | Name={name} | Type={ftype}")
        return

    # 2) List notes for a given deal
    if args.list_deal_notes:
        if not args.deal_id:
            print("ERROR: --list-deal-notes requires --deal-id")
            sys.exit(2)
        notes = list_deal_notes(args.deal_id)
        if not notes:
            print(f"No notes found for deal {args.deal_id}.")
            return
        print(f"Notes for deal {args.deal_id}:")
        for n in notes:
            nid = n.get("id")
            body = n.get("note") or ""
            c_at = n.get("cdate") or n.get("createdTimestamp") or n.get("created_at")
            u_at = n.get("udate") or n.get("updatedTimestamp") or n.get("updated_at")
            snippet = body.strip().replace("\r", " ").replace("\n", " ")
            if len(snippet) > 160:
                snippet = snippet[:157] + "..."
            print(f"- #{nid} | created={c_at} | updated={u_at} | {snippet}")
        return

    # 3) Create a note on a Deal
    if args.add_deal_note is not None:
        if not args.deal_id:
            print("ERROR: --add-deal-note requires --deal-id")
            sys.exit(2)
        result = create_deal_note(args.deal_id, args.add_deal_note)
        print(_pretty(result))
        if not result.get("ok"):
            sys.exit(1)
        return

    # 4) Delete a note by Note ID
    if args.delete_note:
        if not args.note_id:
            print("ERROR: --delete-note requires --note-id")
            sys.exit(2)
        result = delete_note(args.note_id)
        print(_pretty(result))
        if not result.get("ok"):
            sys.exit(1)
        return

    # 5) Set a deal custom field
    if args.deal_id and args.field_id and args.value is not None:
        result = set_deal_custom_field(args.deal_id, args.field_id, args.value)
        print(_pretty(result))
        if not result.get("ok"):
            sys.exit(1)
        return

    ap.print_help()


if __name__ == "__main__":
    main()
