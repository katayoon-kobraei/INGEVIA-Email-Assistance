from __future__ import annotations

import sys
from pathlib import Path

APP_ROOT = Path(__file__).resolve().parent
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

import win32com.client

from src.config import APP_MODE, OUTPUT_ROOT
from src.output.sender_backfill import backfill_missing_sender_names


def main() -> int:
    if APP_MODE != "PROCESSOR":
        return 0
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        changed = backfill_missing_sender_names(OUTPUT_ROOT, outlook)
        print(f"Sender-name migration updated {changed} archived row(s).")
        return 0
    except Exception as exc:
        print(f"Sender-name migration skipped: {exc}")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
