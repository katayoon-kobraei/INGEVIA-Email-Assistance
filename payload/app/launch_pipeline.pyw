from __future__ import annotations

import datetime as dt
import os
import subprocess
import sys
import time
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
LOG_FILE = LOG_DIR / "scheduler.log"
WORKER = APP_ROOT / "pipeline_worker.py"
RESOURCE_EXIT_CODE = 75
MAILBOX_EXIT_CODE = 76
ALREADY_RUNNING_EXIT_CODE = 20
RETRY_SECONDS = max(5, int(os.environ.get("OUTLOOK_RESOURCE_RETRY_SECONDS") or "60"))
MAILBOX_RETRY_SECONDS = max(5, int(os.environ.get("OUTLOOK_MAILBOX_RETRY_SECONDS") or "60"))


def _python_console_exe() -> str:
    current = Path(sys.executable)
    if current.name.casefold() == "pythonw.exe":
        candidate = current.with_name("python.exe")
        if candidate.is_file():
            return str(candidate)
    return str(current)


def _creationflags() -> int:
    return getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0


def _worker_env() -> dict[str, str]:
    """Force UTF-8 inside the console worker even on Spanish Windows.

    The supervisor already decodes captured output as UTF-8, but without these
    variables the child process can still *encode* stdout using cp1252 before
    subprocess captures it. A subject containing an emoji therefore used to
    raise UnicodeEncodeError inside an otherwise healthy Outlook run.
    """
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8:backslashreplace"
    return env


def _run_worker() -> tuple[int, str]:
    completed = subprocess.run(
        [_python_console_exe(), str(WORKER)],
        cwd=str(APP_ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=_worker_env(),
        creationflags=_creationflags(),
        check=False,
    )
    output = "\n".join(
        part.strip() for part in (completed.stdout, completed.stderr) if part and part.strip()
    ).strip()
    return completed.returncode, output


def _append_log(text: str) -> None:
    with LOG_FILE.open("a", encoding="utf-8") as handle:
        handle.write(text.rstrip() + "\n")


def _emit(text: str, interactive: bool) -> None:
    _append_log(text)
    if interactive:
        try:
            print(text, flush=True)
        except Exception:
            pass


def main() -> int:
    interactive = "--interactive" in sys.argv
    started = dt.datetime.now().isoformat(timespec="seconds")

    try:
        with processing_lock():
            _emit(f"\n[{started}] Email processing started", interactive)
            code, output = _run_worker()
            if output:
                _emit(output, interactive)

            if code in {RESOURCE_EXIT_CODE, MAILBOX_EXIT_CODE}:
                if code == RESOURCE_EXIT_CODE:
                    delay = RETRY_SECONDS
                    message = (
                        "Outlook/MAPI resource exhaustion detected. Releasing the first worker completely "
                        f"and retrying once in {delay} seconds..."
                    )
                else:
                    delay = MAILBOX_RETRY_SECONDS
                    message = (
                        "The configured Outlook mailbox is temporarily unavailable. Releasing the first "
                        f"worker completely and retrying once in {delay} seconds..."
                    )
                _emit(message, interactive)
                time.sleep(delay)
                first_code = code
                code, output = _run_worker()
                if output:
                    _emit(output, interactive)

                if code == first_code:
                    if code == RESOURCE_EXIT_CODE:
                        _emit(
                            "Outlook resources are still unavailable after the automatic retry. "
                            "Close and reopen Outlook, then run Process New Emails again.",
                            interactive,
                        )
                    else:
                        _emit(
                            "The configured mailbox is still unavailable after the automatic retry. "
                            "The next scheduled run will try again and the lookback window will catch up "
                            "without losing emails.",
                            interactive,
                        )

            finished = dt.datetime.now().isoformat(timespec="seconds")
            if code == 0:
                _emit(f"[{finished}] Email processing completed", interactive)
            else:
                _emit(f"[{finished}] Email processing failed (exit code {code})", interactive)
            return code

    except ProcessingAlreadyRunning:
        message = "Email processing is already running on this computer. Please wait for the current run to finish."
        _emit(message, interactive)
        return ALREADY_RUNNING_EXIT_CODE
    except Exception as exc:
        _emit(f"Pipeline supervisor failed: {exc}", interactive)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
