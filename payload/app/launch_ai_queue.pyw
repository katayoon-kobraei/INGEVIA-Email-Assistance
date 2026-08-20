from __future__ import annotations

import datetime as dt
import os
import subprocess
import sys
from pathlib import Path

APP_ROOT = Path(__file__).resolve().parent
os.chdir(APP_ROOT)
if str(APP_ROOT) not in sys.path:
    sys.path.insert(0, str(APP_ROOT))

from dotenv import load_dotenv
from src.processing_lock import ProcessingAlreadyRunning, processing_lock

load_dotenv(APP_ROOT / ".env")

LOG_DIR = APP_ROOT / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "ai_queue.log"
WORKER = APP_ROOT / "ai_queue_worker.py"
ALREADY_RUNNING_EXIT_CODE = 20


def _python_console_exe() -> str:
    current = Path(sys.executable)
    if current.name.casefold() == "pythonw.exe":
        candidate = current.with_name("python.exe")
        if candidate.is_file():
            return str(candidate)
    return str(current)


def _creationflags() -> int:
    return getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0


def _append(text: str) -> None:
    with LOG_FILE.open("a", encoding="utf-8") as handle:
        handle.write(text.rstrip() + "\n")


def main() -> int:
    started = dt.datetime.now().isoformat(timespec="seconds")
    try:
        # Reuse the same machine-wide lock as Outlook ingestion. This protects
        # shared index/queue state while still guaranteeing the AI queue itself
        # never opens Outlook.
        with processing_lock():
            completed = subprocess.run(
                [_python_console_exe(), str(WORKER)],
                cwd=str(APP_ROOT),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=_creationflags(),
                check=False,
            )
            output = "\n".join(
                part.strip() for part in (completed.stdout, completed.stderr)
                if part and part.strip()
            ).strip()
            _append(f"\n[{started}] AI queue run started")
            if output:
                _append(output)
            _append(f"[{dt.datetime.now().isoformat(timespec='seconds')}] AI queue exit code {completed.returncode}")
            return completed.returncode
    except ProcessingAlreadyRunning:
        # Normal and expected if the 30-minute Outlook capture happens to overlap
        # this 2-minute trigger. The next queue trigger will run two minutes later.
        return ALREADY_RUNNING_EXIT_CODE
    except Exception as exc:
        _append(f"[{started}] AI queue supervisor failed: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
