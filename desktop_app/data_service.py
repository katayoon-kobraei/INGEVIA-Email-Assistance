from __future__ import annotations

import csv
import json
import os
import re
import subprocess
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from dotenv import load_dotenv


PROJECT_ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = PROJECT_ROOT / ".env"
load_dotenv(ENV_PATH)

OUTPUT_ROOT = Path(os.environ.get("OUTPUT_ROOT") or r"C:\EmailAssistant\Output")
ARCHIVE_ROOT = Path(os.environ.get("ARCHIVE_ROOT") or str(OUTPUT_ROOT))
PROCESSED_CATEGORY_NAME = os.environ.get("PROCESSED_CATEGORY_NAME") or "IA - PROCESADO"
FLAG_PROCESSED_EMAILS = (os.environ.get("FLAG_PROCESSED_EMAILS") or "true").lower() == "true"
ARCHIVE_JUNK_EMAILS = (os.environ.get("ARCHIVE_JUNK_EMAILS") or "true").lower() == "true"
JUNK_ARCHIVE_FOLDER_NAME = os.environ.get("JUNK_ARCHIVE_FOLDER_NAME") or "Archivo"
CHECK_PENDING_RESPONSES = (os.environ.get("CHECK_PENDING_RESPONSES") or "true").lower() == "true"
PENDING_FOLDER_NAME = os.environ.get("PENDING_FOLDER_NAME") or "PENDIENTE DE RESPUESTA"
DESCRIPTIONS_XLSX_PATH = Path(os.environ.get("DESCRIPTIONS_XLSX_PATH") or "") if os.environ.get("DESCRIPTIONS_XLSX_PATH") else None
DESCRIPTION_MAX_CHARS = int(os.environ.get("DESCRIPTION_MAX_CHARS") or "220")
BOSS_EMAIL = os.environ.get("BOSS_EMAIL") or "m.vera@ingevia.com"
ADMINISTRACION_EMAIL = os.environ.get("ADMINISTRACION_EMAIL") or "administracion@ingevia.com"
BILLING_OUTPUT_ROOT = Path(os.environ.get("BILLING_OUTPUT_ROOT") or r"X:\EMAILS")
IGNORE_DOMAINS = os.environ.get("IGNORE_DOMAINS") or "ingevia.com"
IGNORE_SENDERS = os.environ.get("IGNORE_SENDERS") or "fmunoz@munozbosch.com"
PLENERGY_SENDER_DOMAINS = os.environ.get("PLENERGY_SENDER_DOMAINS") or "plenergy.es,plainco.es"
TASK_NAME = "Email AI Assistant"
MAIL_ITEM_CLASS = 43

PROCESSED_IDS_PATH = OUTPUT_ROOT / "_processed_ids.json"
PENDING_LIST_PATH = OUTPUT_ROOT / "pendientes.csv"
PRIORITY_LIST_PATH = OUTPUT_ROOT / "priorities.csv"
REPORT_LOG_PATH = OUTPUT_ROOT / "email_report_log.csv"
REPORT_XLSX_PATH = OUTPUT_ROOT / "Informe de Emails.xlsx"


@dataclass(slots=True)
class OutlookFlagSummary:
    total: int
    rows: list[dict[str, Any]]
    estimated: bool = False
    error: str | None = None


def _parse_date(value: str | None) -> datetime | None:
    if not value:
        return None
    for fmt in (
        "%Y-%m-%d %H:%M",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None


def _read_csv(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    try:
        with path.open("r", newline="", encoding="utf-8-sig") as handle:
            return list(csv.DictReader(handle))
    except (OSError, csv.Error):
        return []


def _email_key(date_text: str | None, subject: str | None) -> tuple[str, str]:
    return (
        str(date_text or "").strip(),
        " ".join(str(subject or "").split()).casefold(),
    )


def load_processed_ids() -> set[str]:
    try:
        if PROCESSED_IDS_PATH.exists():
            data = json.loads(PROCESSED_IDS_PATH.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return {str(item) for item in data if item}
    except (OSError, json.JSONDecodeError):
        pass
    return set()


def load_priority_lookup() -> dict[tuple[str, str], int]:
    lookup: dict[tuple[str, str], int] = {}
    for row in _read_csv(PRIORITY_LIST_PATH):
        try:
            priority = int(row.get("Priority") or 0)
        except (TypeError, ValueError):
            continue
        if 1 <= priority <= 5:
            lookup[_email_key(row.get("Date"), row.get("Subject"))] = priority
    return lookup


def load_report_lookup() -> dict[tuple[str, str], dict[str, Any]]:
    lookup: dict[tuple[str, str], dict[str, Any]] = {}
    for row in _read_csv(REPORT_LOG_PATH):
        date_text = " ".join(
            part for part in (str(row.get("Date") or ""), str(row.get("Time") or "")) if part
        ).strip()
        normalized = {
            **row,
            "DateTime": date_text,
            "FechaHora": date_text,
            "SenderName": row.get("Sender Name", ""),
            "SenderEmail": row.get("Sender Email", ""),
            "Summary": row.get("Summary", ""),
            "Path": row.get("Path", ""),
            "NombreRemitente": row.get("Sender Name", ""),
            "EmailRemitente": row.get("Sender Email", ""),
            "Resumen": row.get("Summary", ""),
            "Ruta": row.get("Path", ""),
            "_parsed_date": _parse_date(date_text),
        }
        lookup[_email_key(date_text, row.get("Subject"))] = normalized
    return lookup


def load_pending_lookup() -> set[tuple[str, str]]:
    return {
        _email_key(row.get("Date"), row.get("Subject"))
        for row in _read_csv(PENDING_LIST_PATH)
    }


def load_index_rows() -> list[dict[str, Any]]:
    rows = _read_csv(OUTPUT_ROOT / "index.csv")
    priority_lookup = load_priority_lookup()
    report_lookup = load_report_lookup()
    pending_lookup = load_pending_lookup()

    for row in rows:
        date_text = str(row.get("Date") or "")
        subject = str(row.get("Subject") or "")
        key = _email_key(date_text, subject)
        row["_parsed_date"] = _parse_date(date_text)
        try:
            row["_attachment_count"] = int(row.get("Attachments") or 0)
        except (TypeError, ValueError):
            row["_attachment_count"] = 0
        priority = priority_lookup.get(key, 0)
        row["Priority"] = priority if priority else ""
        row["_priority"] = priority
        report_row = report_lookup.get(key, {})
        row["Summary"] = report_row.get("Summary", "")
        row["_summary"] = report_row.get("Summary", "")
        row["_pending"] = key in pending_lookup
        row["_status"] = "REVISAR" if row.get("Project Folder") == "UNSORTED" else "PROCESADO"

    rows.sort(key=lambda row: row.get("_parsed_date") or datetime.min, reverse=True)
    return rows


def load_report_rows() -> list[dict[str, Any]]:
    priority_lookup = load_priority_lookup()
    pending_lookup = load_pending_lookup()
    rows: list[dict[str, Any]] = []
    for row in _read_csv(REPORT_LOG_PATH):
        date_text = " ".join(
            part for part in (str(row.get("Date") or ""), str(row.get("Time") or "")) if part
        ).strip()
        subject = str(row.get("Subject") or "")
        key = _email_key(date_text, subject)
        priority = priority_lookup.get(key, 0)
        normalized = {
            **row,
            "DateTime": date_text,
            "FechaHora": date_text,
            "SenderName": row.get("Sender Name", ""),
            "SenderEmail": row.get("Sender Email", ""),
            "Subject": subject,
            "Summary": row.get("Summary", ""),
            "Path": row.get("Path", ""),
            "NombreRemitente": row.get("Sender Name", ""),
            "EmailRemitente": row.get("Sender Email", ""),
            "Asunto": subject,
            "Resumen": row.get("Summary", ""),
            "Ruta": row.get("Path", ""),
            "Priority": priority if priority else "",
            "Prioridad": priority if priority else "",
            "_priority": priority,
            "_pending": key in pending_lookup,
            "_parsed_date": _parse_date(date_text),
        }
        rows.append(normalized)
    rows.sort(key=lambda row: row.get("_parsed_date") or datetime.min, reverse=True)
    return rows


def load_activity_rows() -> list[dict[str, Any]]:
    """Load an optional activity log, falling back to archived emails.

    Older installations may contain ``activity_log.csv``. The current backend
    records filed messages in ``index.csv`` and produces a human-readable report,
    so the UI can safely fall back to those rows when the activity log is absent.
    """
    rows = _read_csv(OUTPUT_ROOT / "activity_log.csv")
    if not rows:
        return load_index_rows()
    priority_lookup = load_priority_lookup()
    pending_lookup = load_pending_lookup()
    for row in rows:
        date_text = str(row.get("Date") or "")
        subject = str(row.get("Subject") or "")
        key = _email_key(date_text, subject)
        priority = priority_lookup.get(key, 0)
        row["_parsed_date"] = _parse_date(date_text)
        row["Priority"] = priority if priority else ""
        row["_priority"] = priority
        row["_pending"] = key in pending_lookup
    rows.sort(key=lambda row: row.get("_parsed_date") or datetime.min, reverse=True)
    return rows


def load_pending_rows() -> list[dict[str, Any]]:
    priority_lookup = load_priority_lookup()
    normalized: list[dict[str, Any]] = []
    for row in _read_csv(PENDING_LIST_PATH):
        date = str(row.get("Date") or "")
        sender = str(row.get("Sender") or "")
        subject = str(row.get("Subject") or "")
        priority = priority_lookup.get(_email_key(date, subject), 0)
        normalized.append(
            {
                "Date": date,
                "Sender": sender,
                "Subject": subject,
                "Priority": priority if priority else "",
                "Fecha": date,
                "Remitente": sender,
                "Asunto": subject,
                "Prioridad": priority if priority else "",
                "_priority": priority,
                "_parsed_date": _parse_date(date),
            }
        )
    normalized.sort(key=lambda row: row.get("_parsed_date") or datetime.min, reverse=True)
    return normalized


def pending_count() -> int:
    return len(load_pending_rows())


def processed_ids_count() -> int:
    ids = load_processed_ids()
    return len(ids) if ids else len(load_index_rows())


def load_attachment_rows() -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if not OUTPUT_ROOT.is_dir():
        return rows

    for metadata_path in OUTPUT_ROOT.rglob("metadata.json"):
        try:
            metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue

        for attachment in metadata.get("attachments", []):
            filename = str(attachment.get("filename", ""))
            status = str(attachment.get("status", "unknown"))
            reason = str(attachment.get("reason", ""))
            folder = metadata_path.parent
            path = folder / filename
            date = metadata.get("timestamp", "")
            project = metadata.get("project_folder", "")
            subject = metadata.get("subject", "")
            rows.append(
                {
                    "Fecha": date,
                    "Proyecto": project,
                    "Asunto": subject,
                    "Archivo": filename,
                    "Estado": status,
                    "Motivo": reason,
                    "Carpeta": str(folder),
                    "Ruta": str(path),
                    "Date": date,
                    "Project": project,
                    "Subject": subject,
                    "File": filename,
                    "Status": status,
                    "Reason": reason,
                    "Folder": str(folder),
                    "Path": str(path),
                }
            )

    rows.sort(key=lambda row: str(row.get("Fecha", "")), reverse=True)
    return rows



def get_report_excel_info() -> dict[str, Any]:
    """Return status information for the generated Excel email report."""
    path = REPORT_XLSX_PATH
    info: dict[str, Any] = {
        "path": str(path),
        "name": path.name,
        "exists": False,
        "size_bytes": 0,
        "size_text": "—",
        "modified": "—",
        "entries": len(load_report_rows()),
    }

    try:
        if not path.is_file():
            return info

        stat = path.stat()
        size_bytes = int(stat.st_size)
        if size_bytes < 1024:
            size_text = f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            size_text = f"{size_bytes / 1024:.1f} KB"
        else:
            size_text = f"{size_bytes / (1024 * 1024):.1f} MB"

        info.update(
            {
                "exists": True,
                "size_bytes": size_bytes,
                "size_text": size_text,
                "modified": datetime.fromtimestamp(stat.st_mtime).strftime(
                    "%Y-%m-%d %H:%M"
                ),
            }
        )
    except OSError:
        pass

    return info

def get_scheduler_status() -> dict[str, Any] | None:
    if os.name != "nt":
        return None
    try:
        command = (
            f"Get-ScheduledTaskInfo -TaskName '{TASK_NAME}' | "
            "Select-Object LastRunTime,LastTaskResult,NextRunTime | ConvertTo-Json"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            capture_output=True,
            text=True,
            timeout=10,
            check=False,
        )
        if result.returncode != 0 or not result.stdout.strip():
            return None
        return json.loads(result.stdout)
    except (OSError, subprocess.SubprocessError, json.JSONDecodeError):
        return None


def open_path(path: str | Path) -> tuple[bool, str]:
    if not path:
        return False, "No path was configured / No se configuró ninguna ruta."
    target = Path(path)
    if not target.exists():
        return False, f"Path does not exist / No existe la ruta:\n{target}"
    try:
        if os.name == "nt":
            os.startfile(str(target))  # type: ignore[attr-defined]
        elif os.name == "posix":
            subprocess.Popen(["xdg-open", str(target)])
        else:
            return False, "This operating system cannot open the path automatically."
        return True, ""
    except OSError as exc:
        return False, str(exc)


def _safe_get(item: Any, property_name: str, default: Any = None) -> Any:
    try:
        value = getattr(item, property_name)
        return default if value is None else value
    except Exception:
        return default


def _categories(item: Any) -> set[str]:
    raw = str(_safe_get(item, "Categories", "") or "")
    return {part.strip() for part in re.split(r"[,;]", raw) if part.strip()}


def _outlook_item_row(item: Any, direction: str) -> dict[str, Any]:
    timestamp = _safe_get(item, "ReceivedTime") or _safe_get(item, "SentOn")
    if timestamp is not None:
        try:
            date_text = timestamp.strftime("%Y-%m-%d %H:%M")
        except Exception:
            date_text = str(timestamp)
    else:
        date_text = ""

    contact = (
        _safe_get(item, "SenderEmailAddress", "")
        if direction == "ENTRANTE"
        else _safe_get(item, "To", "")
    )
    subject = str(_safe_get(item, "Subject", "(Sin asunto)") or "(Sin asunto)")
    categories = str(_safe_get(item, "Categories", "") or "")
    entry_id = str(_safe_get(item, "EntryID", "") or "")
    return {
        "Fecha": date_text,
        "Dirección": direction,
        "Contacto": contact or "",
        "Asunto": subject,
        "Categoría": categories,
        "Date": date_text,
        "Direction": direction,
        "Contact": contact or "",
        "Subject": subject,
        "Category": categories,
        "EntryID": entry_id,
    }


def load_outlook_flagged(limit: int = 250) -> OutlookFlagSummary:
    """Return only Outlook messages processed by this local installation.

    The local ``_processed_ids.json`` is used as the source of truth. This avoids
    counting thousands of historical or manually flagged messages in the boss's
    mailbox that happen to use the same Outlook category.
    """
    if os.name != "nt":
        return OutlookFlagSummary(0, [], error="Outlook is only available on Windows.")

    processed_ids = load_processed_ids()
    if not processed_ids:
        return OutlookFlagSummary(0, [])

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        return OutlookFlagSummary(0, [], error=f"pywin32 missing: {exc}")

    pythoncom.CoInitialize()
    try:
        namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        rows: list[dict[str, Any]] = []
        total = 0
        for entry_id in processed_ids:
            try:
                item = namespace.GetItemFromID(entry_id)
            except Exception:
                continue
            if _safe_get(item, "Class") != MAIL_ITEM_CLASS:
                continue
            categories = _categories(item)
            flag_status = int(_safe_get(item, "FlagStatus", 0) or 0)
            is_stamped = PROCESSED_CATEGORY_NAME in categories or flag_status == 2
            if not is_stamped:
                continue
            total += 1
            if len(rows) < limit:
                direction = "SALIENTE" if bool(_safe_get(item, "Sent", False)) else "ENTRANTE"
                rows.append(_outlook_item_row(item, direction))

        rows.sort(key=lambda row: row.get("Fecha", ""), reverse=True)
        return OutlookFlagSummary(total=total, rows=rows[:limit], estimated=False)
    except Exception as exc:
        return OutlookFlagSummary(0, [], error=str(exc))
    finally:
        pythoncom.CoUninitialize()
