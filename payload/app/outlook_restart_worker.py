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

from dotenv import load_dotenv

load_dotenv(APP_ROOT / ".env")


def main() -> int:
    pythoncom = None
    try:
        if os.name == "nt":
            import pythoncom as _pythoncom

            pythoncom = _pythoncom
            pythoncom.CoInitialize()

        from src.outlook_restart import run

        return int(run() or 0)
    except Exception:
        traceback.print_exc()
        return 1
    finally:
        # Same reasoning as pipeline_worker.py: all Outlook COM references in
        # this short-lived worker die with the process, so CoUninitialize +
        # gc.collect() here makes that as deterministic as possible before
        # Windows tears the process down.
        try:
            if pythoncom is not None:
                pythoncom.CoUninitialize()
        finally:
            gc.collect()


if __name__ == "__main__":
    raise SystemExit(main())