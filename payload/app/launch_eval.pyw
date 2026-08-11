from __future__ import annotations

import contextlib
import datetime as dt
import os
import runpy
import sys
import traceback
from pathlib import Path

APP_ROOT = Path(__file__).resolve().parent
os.chdir(APP_ROOT)
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

LOG_DIR = APP_ROOT / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "daily_evaluation.log"

with LOG_FILE.open("a", encoding="utf-8") as log, contextlib.redirect_stdout(log), contextlib.redirect_stderr(log):
    print(f"\n[{dt.datetime.now().isoformat(timespec='seconds')}] Daily evaluation started")
    try:
        runpy.run_path(str(APP_ROOT / "eval_project_agent.py"), run_name="__main__")
        print(f"[{dt.datetime.now().isoformat(timespec='seconds')}] Daily evaluation completed")
    except Exception:
        traceback.print_exc()
        print(f"[{dt.datetime.now().isoformat(timespec='seconds')}] Daily evaluation failed")
