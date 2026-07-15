import subprocess
import os

DANGEROUS_EXTENSIONS = {
    ".exe", ".scr", ".bat", ".cmd", ".com", ".pif", ".vbs", ".vbe",
    ".js", ".jse", ".ws", ".wsf", ".ps1", ".jar", ".msi",
    ".docm", ".xlsm", ".pptm",
}

DEFENDER_PATH = r"C:\Program Files\Windows Defender\MpCmdRun.exe"

def is_dangerous_extension(filename):
    return os.path.splitext(filename)[1].lower() in DANGEROUS_EXTENSIONS

def scan_with_defender(file_path):
    result = subprocess.run(
        [DEFENDER_PATH, "-Scan", "-ScanType", "3", "-File", file_path],
        capture_output=True, text=True
    )
    return result.returncode == 0

def check_attachment(file_path):
    filename = os.path.basename(file_path)
    if is_dangerous_extension(filename):
        return False, f"Blocked file type: {os.path.splitext(filename)[1]}"
    if not scan_with_defender(file_path):
        return False, "Flagged by antivirus scan"
    return True, "Clean"