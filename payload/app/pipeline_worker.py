from __future__ import annotations

import gc
import os
import sys
import traceback
from pathlib import Path

APP_ROOT = Path(__file__).resolve().parent
os.chdir(APP_ROOT)
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))


def _configure_utf8_stdio() -> None:
    """Make all diagnostic output safe for arbitrary email subjects."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="backslashreplace")
        except Exception:
            pass


_configure_utf8_stdio()

from dotenv import load_dotenv

load_dotenv(APP_ROOT / ".env")

RESOURCE_EXIT_CODE = 75
MAILBOX_EXIT_CODE = 76

from src.outlook_errors import is_outlook_resource_error, is_outlook_mailbox_unavailable


def main() -> int:
    pythoncom = None
    try:
        if os.name == "nt":
            import pythoncom as _pythoncom

            pythoncom = _pythoncom
            pythoncom.CoInitialize()

        from src.pipeline import run

        run()
        return 0
    except Exception as exc:
        traceback.print_exc()
        if is_outlook_resource_error(exc):
            print("OUTLOOK_RESOURCE_EXHAUSTED: Outlook/MAPI resources are temporarily unavailable.")
            return RESOURCE_EXIT_CODE
        if is_outlook_mailbox_unavailable(exc):
            print("OUTLOOK_MAILBOX_TEMPORARILY_UNAVAILABLE: configured mailbox could not be opened in this Outlook session.")
            return MAILBOX_EXIT_CODE
        return 1
    finally:
        # Release Python/pywin32 wrappers while COM is still initialized.  The
        # process exit remains the final hard boundary, but this ordering avoids
        # leaving Outlook/MAPI wrappers pending until interpreter teardown.
        gc.collect()
        try:
            if pythoncom is not None:
                pythoncom.CoUninitialize()
        finally:
            gc.collect()


if __name__ == "__main__":
    raise SystemExit(main())
