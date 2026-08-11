from __future__ import annotations

import os
import subprocess
from typing import Any


DANGEROUS_EXTENSIONS = {
    ".exe", ".scr", ".bat", ".cmd", ".com", ".pif", ".vbs", ".vbe",
    ".js", ".jse", ".ws", ".wsf", ".ps1", ".jar", ".msi",
    ".docm", ".xlsm", ".pptm",
}

DEFENDER_PATH = r"C:\Program Files\Windows Defender\MpCmdRun.exe"


def _hidden_process_options() -> dict[str, Any]:
    """Prevent command-line scanners from opening a visible console window."""
    if os.name != "nt":
        return {}

    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    return {
        "startupinfo": startupinfo,
        "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0),
    }


def is_dangerous_extension(filename: str) -> bool:
    return os.path.splitext(filename)[1].lower() in DANGEROUS_EXTENSIONS


def scan_with_defender(file_path: str) -> bool:
    result = subprocess.run(
        [DEFENDER_PATH, "-Scan", "-ScanType", "3", "-File", file_path],
        capture_output=True,
        text=True,
        stdin=subprocess.DEVNULL,
        check=False,
        **_hidden_process_options(),
    )
    return result.returncode == 0


def check_attachment(file_path: str) -> tuple[bool, str]:
    filename = os.path.basename(file_path)
    if is_dangerous_extension(filename):
        return False, f"Blocked file type: {os.path.splitext(filename)[1]}"
    if not scan_with_defender(file_path):
        return False, "Flagged by antivirus scan"
    return True, "Clean"
