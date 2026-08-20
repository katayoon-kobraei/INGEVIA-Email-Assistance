from __future__ import annotations

"""
outlook_restart.py

Nightly maintenance: gracefully closes and relaunches OUTLOOK.EXE so its
accumulated COM/MAPI/GDI resource usage resets on a schedule, without ever
requiring the boss to restart the PC itself. This is the direct fix for
"Outlook has exhausted all shared resources" errors that build up on a
machine that stays logged in for days or weeks at a time -- see
src/outlook_errors.py for the error text that failure mode produces, and
pipeline_worker.py for how the main capture run reacts to it mid-day.

Invoked twice a night (21:00 and 23:00, both outside the 05:30-20:00
capture window -- see setup_outlook_restart_scheduler.ps1) by the exact
same idempotent entry point, run(). If the 21:00 attempt already
succeeded, the 23:00 run just confirms Outlook is running and does
nothing further. The second trigger exists specifically to cover the case
where Outlook is already stuck at 21:00 and needs another hour before it
will close.

GRACEFUL ONLY -- this deliberately never force-kills OUTLOOK.EXE. If
Outlook won't quit cleanly (including the case where it's already stuck
in the very resource-exhausted state this feature exists to prevent),
run() leaves it alone and logs clearly rather than risking the boss
losing an unsaved draft to a hard kill. That was an explicit choice, not
an oversight -- see the two-attempt nightly schedule above for how a
stuck Outlook still gets handled without a force-kill.
"""

import os
import subprocess
import time
import winreg
from pathlib import Path

QUIT_TIMEOUT_SECONDS = 45
QUIT_POLL_INTERVAL_SECONDS = 2


def _hidden_process_options() -> dict:
    return {"creationflags": subprocess.CREATE_NO_WINDOW} if hasattr(subprocess, "CREATE_NO_WINDOW") else {}


def _is_outlook_running() -> bool:
    """True if any OUTLOOK.EXE process is currently running for this
    Windows session. Uses tasklist rather than adding a new psutil
    dependency to the project -- it's always present on Windows and
    this only ever needs to run twice a night."""
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq OUTLOOK.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=15,
            **_hidden_process_options(),
        )
    except Exception:
        # Can't tell either way -- assume it might still be running so a
        # needed quit attempt is never silently skipped.
        return True
    return "OUTLOOK.EXE" in (result.stdout or "").upper()


def _resolve_outlook_exe() -> str | None:
    """Looks up the real OUTLOOK.EXE path from the Windows "App Paths"
    registry key. Works across MSI and Click-to-Run installs, 32-bit
    and 64-bit Office, without hardcoding a Program Files path that
    varies machine to machine."""
    for hive, subkey in (
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"),
    ):
        try:
            with winreg.OpenKey(hive, subkey) as key:
                path, _ = winreg.QueryValueEx(key, None)
                if path and Path(path).is_file():
                    return path
        except OSError:
            continue
    return None


def _launch_outlook() -> tuple[bool, str]:
    exe = _resolve_outlook_exe()
    try:
        if exe:
            subprocess.Popen([exe], **_hidden_process_options())
            return True, f"Launched Outlook from {exe}."
        # Registry lookup failed -- fall back to the same shell App Paths
        # resolution Windows itself uses for the Run dialog/Start menu.
        os.startfile("outlook")
        return True, "Launched Outlook via shell association (registry path not found)."
    except Exception as exc:
        return False, f"Could not launch Outlook: {exc}"


def _quit_outlook_gracefully() -> tuple[bool, str]:
    """Attempts Application.Quit() over COM, then polls tasklist until
    the process actually exits (Quit() returning does not mean the
    process has fully terminated yet). Never force-kills -- see the
    module docstring."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        try:
            outlook_app = win32com.client.Dispatch("Outlook.Application")
        except Exception as exc:
            return False, (
                "Could not attach to Outlook via COM -- it may already be in a "
                f"resource-exhausted state: {exc}"
            )
        try:
            outlook_app.Quit()
        except Exception as exc:
            return False, f"Outlook.Quit() raised an error: {exc}"
        finally:
            outlook_app = None
    finally:
        pythoncom.CoUninitialize()

    waited = 0
    while waited < QUIT_TIMEOUT_SECONDS:
        if not _is_outlook_running():
            return True, f"Outlook closed gracefully after {waited}s."
        time.sleep(QUIT_POLL_INTERVAL_SECONDS)
        waited += QUIT_POLL_INTERVAL_SECONDS

    return False, (
        f"Outlook did not finish closing within {QUIT_TIMEOUT_SECONDS}s of Quit(). "
        "Leaving it running rather than force-killing it -- the next scheduled "
        "attempt (or tonight's second attempt) will try again."
    )


def run() -> int:
    """Ensures Outlook ends this run freshly launched, without ever
    force-killing it. Returns 0 on success (including the no-op case
    where Outlook simply wasn't running yet), 1 if a running Outlook
    could not be closed gracefully this attempt."""
    if not _is_outlook_running():
        launched, message = _launch_outlook()
        print(f"Outlook was not running. {message}")
        return 0 if launched else 1

    quit_ok, quit_message = _quit_outlook_gracefully()
    print(quit_message)
    if not quit_ok:
        return 1

    launched, launch_message = _launch_outlook()
    print(launch_message)
    return 0 if launched else 1