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
LOG_FILE = LOG_DIR / "outlook_restart.log"
WORKER = APP_ROOT / "outlook_restart_worker.py"
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
        # Reuse the same machine-wide lock as Outlook ingestion and the AI
        # queue, so this can never quit Outlook out from under an
        # in-progress capture run, and a capture run can never start while
        # Outlook is mid-restart. Both scheduled restart times (21:00 and
        # 23:00) sit outside the 05:30-20:00 capture window, so contention
        # should be rare in practice.
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
            _append(f"\n[{started}] Outlook daily restart started")
            if output:
                _append(output)
            _append(
                f"[{dt.datetime.now().isoformat(timespec='seconds')}] "
                f"Outlook daily restart exit code {completed.returncode}"
            )
            return completed.returncode
    except ProcessingAlreadyRunning:
        _append(f"[{started}] Outlook daily restart skipped -- processing lock held by another run.")
        return ALREADY_RUNNING_EXIT_CODE
    except Exception as exc:
        _append(f"[{started}] Outlook restart supervisor failed: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())