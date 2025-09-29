#!/usr/bin/env python3
# it_probe.py — quick tester for https://api.golden-tech.de/glts_ac
#
# Usage examples:
#   python it_probe.py
#   python it_probe.py --raw
#   python it_probe.py --token "%IT_AUTH_TOKEN%"
#   python it_probe.py --param id=15619 --param Email=test@example.com --raw
#   python it_probe.py --endpoint https://api.golden-tech.de/glts_ac --read-timeout 90 --connect-timeout 15
#
# Env fallbacks:
#   IT_BASE_URL (default: https://api.golden-tech.de/glts_ac)
#   IT_AUTH_TOKEN or AC_AUTH_TOKEN (Authorization header if set)
#   IT_READ_TIMEOUT (default 90)   IT_CONNECT_TIMEOUT (default 15)
#   IT_USER_AGENT (default ExcelProcessor-IT-Probe/1.0)

import os
import sys
import json
import time
import argparse
from typing import Any, Dict, List, Optional, Tuple

import requests


def env_int(name: str, default: int) -> int:
    try:
        return int(os.getenv(name, str(default)))
    except Exception:
        return default


DEFAULT_ENDPOINT = os.getenv("IT_BASE_URL", "https://api.golden-tech.de/glts_ac")
DEFAULT_TOKEN = os.getenv("IT_AUTH_TOKEN") or os.getenv("AC_AUTH_TOKEN")
DEFAULT_READ = env_int("IT_READ_TIMEOUT", 90)
DEFAULT_CONN = env_int("IT_CONNECT_TIMEOUT", 15)
DEFAULT_UA = os.getenv("IT_USER_AGENT", "ExcelProcessor-IT-Probe/1.0")


def pick_array(data: Any) -> Tuple[List[Dict[str, Any]], Optional[str]]:
    """Return (array, key_used) from typical shapes."""
    if isinstance(data, list):
        return data, "<top-level-list>"
    if isinstance(data, dict):
        for k in ("data", "leads", "results", "items"):
            v = data.get(k)
            if isinstance(v, list):
                return v, k
    return [], None


def redact_token(tok: Optional[str]) -> str:
    if not tok:
        return "(none)"
    if len(tok) <= 8:
        return "*" * len(tok)
    return tok[:4] + "…" + tok[-4:]


def main() -> None:
    ap = argparse.ArgumentParser(description="Probe the GLTS AC IT endpoint")
    ap.add_argument(
        "--endpoint",
        default=DEFAULT_ENDPOINT,
        help="Base endpoint (default: env IT_BASE_URL or https://api.golden-tech.de/glts_ac)",
    )
    ap.add_argument(
        "--token",
        default=DEFAULT_TOKEN,
        help="Authorization header value (default: env IT_AUTH_TOKEN or AC_AUTH_TOKEN)",
    )
    ap.add_argument(
        "--read-timeout",
        type=int,
        default=DEFAULT_READ,
        help=f"Read timeout seconds (default: {DEFAULT_READ})",
    )
    ap.add_argument(
        "--connect-timeout",
        type=int,
        default=DEFAULT_CONN,
        help=f"Connect timeout seconds (default: {DEFAULT_CONN})",
    )
    ap.add_argument("--raw", action="store_true", help="Print raw JSON response")
    ap.add_argument(
        "--head", type=int, default=5, help="Show first N items from detected array"
    )
    ap.add_argument(
        "--fields",
        action="store_true",
        help="Also list union of keys from first 100 items",
    )
    ap.add_argument(
        "--param",
        action="append",
        default=[],
        help="Add query param as key=value (can be repeated)",
    )
    args = ap.parse_args()

    # Build query params dict from --param key=value pairs
    params: Dict[str, Any] = {}
    for kv in args.param:
        if "=" in kv:
            k, v = kv.split("=", 1)
            params[k.strip()] = v.strip()
        else:
            print(
                f"[warn] ignoring malformed --param '{kv}' (expected key=value)",
                file=sys.stderr,
            )

    headers = {
        "Accept": "application/json",
        "User-Agent": DEFAULT_UA,
    }
    if args.token:
        headers["Authorization"] = args.token

    print(f"--- IT probe ---")
    print(f"URL: {args.endpoint}")
    if params:
        print(f"Query: {json.dumps(params, ensure_ascii=False)}")
    print(f"Authorization: {redact_token(args.token)}")
    print(f"Timeouts: connect={args.connect_timeout}s, read={args.read_timeout}s")

    t0 = time.time()
    try:
        r = requests.get(
            args.endpoint,
            headers=headers,
            params=params,
            timeout=(args.connect_timeout, args.read_timeout),
        )
        elapsed = round(time.time() - t0, 3)
    except requests.ReadTimeout:
        print(f"\n❌ Read timeout after {args.read_timeout}s", file=sys.stderr)
        sys.exit(1)
    except requests.ConnectTimeout:
        print(f"\n❌ Connect timeout after {args.connect_timeout}s", file=sys.stderr)
        sys.exit(1)
    except requests.RequestException as e:
        print(f"\n❌ Request error: {e}", file=sys.stderr)
        sys.exit(1)

    ct = (r.headers.get("content-type") or "").lower()
    print(f"\nHTTP {r.status_code}  elapsed={elapsed}s  content-type={ct}")
    print(f"Final URL: {r.url}")

    # Try to parse JSON (server typically returns JSON even on 4xx)
    data: Any = None
    if (
        "application/json" in ct
        or r.text.strip().startswith("{")
        or r.text.strip().startswith("[")
    ):
        try:
            data = r.json()
        except Exception:
            pass

    if data is None:
        # Non-JSON body
        body = r.text
        if len(body) > 1000:
            body = body[:1000] + "…"
        print("\n[body]")
        print(body)
        sys.exit(0)

    if args.raw:
        print("\n[json]")
        print(json.dumps(data, ensure_ascii=False, indent=2))
        sys.exit(0)

    # Summary view
    top_keys = (
        list(data.keys())
        if isinstance(data, dict)
        else (["<list>"] if isinstance(data, list) else [])
    )
    print(f"\n[top-level keys] {top_keys}")

    arr, key_used = pick_array(data)
    if key_used:
        print(f"[array] found {len(arr)} item(s) under '{key_used}'")
        head_n = max(0, args.head)
        if head_n and arr:
            sample = arr[:head_n]

            # compact preview
            def compact(x: Any) -> Any:
                if not isinstance(x, dict):
                    return x
                # pick a few common fields if present
                picks = {
                    k: x.get(k)
                    for k in (
                        "id",
                        "title",
                        "email",
                        "owner",
                        "Opponent",
                        "GamblingProvider",
                    )
                }
                # include note snippet if present
                note = None
                if isinstance(x.get("notes"), list) and x["notes"]:
                    try:
                        # pick the latest by mdate/cdate
                        latest = sorted(
                            x["notes"],
                            key=lambda n: (n.get("mdate") or n.get("cdate") or ""),
                        )[-1]
                        note = latest.get("note")
                    except Exception:
                        pass
                if note:
                    picks["note"] = (
                        (note[:120] + "…")
                        if isinstance(note, str) and len(note) > 120
                        else note
                    )
                return {k: v for k, v in picks.items() if v is not None}

            print("\n[sample]")
            print(
                json.dumps([compact(x) for x in sample], ensure_ascii=False, indent=2)
            )

        if args.fields and arr:
            keys = set()
            for it in arr[:1000]:
                if isinstance(it, dict):
                    keys.update(it.keys())
            print("\n[fields]")
            print(", ".join(sorted(keys)))
    else:
        # No array detected; print the object
        print("\n[object]")
        print(json.dumps(data, ensure_ascii=False, indent=2))

    # Special callout for typical 422 validation payload
    if isinstance(data, dict) and "detail" in data and r.status_code == 422:
        print(
            "\n[note] Server returned 422 with 'detail' — base GET may require filters on this environment."
        )


if __name__ == "__main__":
    main()
