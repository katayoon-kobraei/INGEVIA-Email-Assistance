from __future__ import annotations

import csv
import json
import os
import subprocess
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = PROJECT_ROOT / ".env"
load_dotenv(ENV_PATH)

OUTPUT_ROOT = Path(os.environ.get("OUTPUT_ROOT", r"C:\EmailAssistant\Output"))
ARCHIVE_ROOT = Path(os.environ.get("ARCHIVE_ROOT", str(OUTPUT_ROOT)))
PROCESSED_CATEGORY_NAME = os.environ.get("PROCESSED_CATEGORY_NAME", "IA - PROCESADO")
FLAG_PROCESSED_EMAILS = os.environ.get("FLAG_PROCESSED_EMAILS", "true").lower() == "true"
TASK_NAME = "Email AI Assistant"


@dataclass(slots=True)
class OutlookFlagSummary:
    total: int
    rows: list[dict[str, Any]]
    estimated: bool = False
    error: str | None = None


def _parse_date(value: str | None) -> datetime | None:
    if not value:
        return None
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(value)
    except ValueError:
        return None


def load_index_rows() -> list[dict[str, Any]]:
    index_path = OUTPUT_ROOT / "index.csv"
    if not index_path.exists():
        return []

    with index_path.open("r", newline="", encoding="utf-8-sig") as handle:
        rows = list(csv.DictReader(handle))

    for row in rows:
        row["_parsed_date"] = _parse_date(row.get("Date"))
        try:
            row["_attachment_count"] = int(row.get("Attachments") or 0)
        except (TypeError, ValueError):
            row["_attachment_count"] = 0
        row["_status"] = "REVISAR" if row.get("Project Folder") == "UNSORTED" else "PROCESADO"

    rows.sort(key=lambda r: r.get("_parsed_date") or datetime.min, reverse=True)
    return rows


def processed_ids_count() -> int:
    path = OUTPUT_ROOT / "_processed_ids.json"
    try:
        if path.exists():
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return len(set(str(item) for item in data))
    except (OSError, json.JSONDecodeError):
        pass
    return len(load_index_rows())


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
            folder = metadata_path.parent
            rows.append(
                {
                    "Fecha": metadata.get("timestamp", ""),
                    "Proyecto": metadata.get("project_folder", ""),
                    "Asunto": metadata.get("subject", ""),
                    "Archivo": filename,
                    "Estado": status,
                    "Motivo": attachment.get("reason", ""),
                    "Carpeta": str(folder),
                    "Ruta": str(folder / filename),
                }
            )

    rows.sort(key=lambda r: str(r.get("Fecha", "")), reverse=True)
    return rows


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
    target = Path(path)
    if not target.exists():
        return False, f"No existe la ruta:\n{target}"
    try:
        if os.name == "nt":
            os.startfile(str(target))  # type: ignore[attr-defined]
        elif os.name == "posix":
            subprocess.Popen(["xdg-open", str(target)])
        else:
            return False, "Este sistema operativo no permite abrir la ruta automáticamente."
        return True, ""
    except OSError as exc:
        return False, str(exc)


def _categories(item: Any) -> set[str]:
    raw = str(getattr(item, "Categories", "") or "")
    return {part.strip() for part in raw.split(",") if part.strip()}


def _outlook_item_row(item: Any, direction: str, time_field: str) -> dict[str, Any]:
    timestamp = getattr(item, time_field, None)
    if timestamp is not None:
        try:
            date_text = timestamp.strftime("%Y-%m-%d %H:%M")
        except Exception:
            date_text = str(timestamp)
    else:
        date_text = ""

    contact = ""
    try:
        contact = item.SenderEmailAddress if direction == "ENTRANTE" else item.To
    except Exception:
        pass

    return {
        "Fecha": date_text,
        "Dirección": direction,
        "Contacto": contact or "",
        "Asunto": str(getattr(item, "Subject", "") or "(Sin asunto)"),
        "Categoría": str(getattr(item, "Categories", "") or ""),
        "EntryID": str(getattr(item, "EntryID", "") or ""),
    }


def load_outlook_flagged(limit: int = 250, fallback_scan_limit: int = 5000) -> OutlookFlagSummary:
    """Return emails stamped by this assistant in Outlook.

    The preferred path uses Outlook's category restriction and gives an exact
    count. If the local Outlook provider rejects that query, the function falls
    back to scanning the newest messages and marks the result as estimated.
    """
    if os.name != "nt":
        return OutlookFlagSummary(0, [], error="Outlook solo está disponible en Windows.")

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        return OutlookFlagSummary(0, [], error=f"Falta pywin32: {exc}")

    pythoncom.CoInitialize()
    try:
        namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        folders = (
            (6, "ENTRANTE", "ReceivedTime"),
            (5, "SALIENTE", "SentOn"),
        )
        category = PROCESSED_CATEGORY_NAME.replace("'", "''")
        query = f'@SQL="urn:schemas-microsoft-com:office:office#Keywords" ci_phrasematch \'{category}\''

        rows: list[dict[str, Any]] = []
        total = 0
        estimated = False

        for folder_id, direction, time_field in folders:
            folder = namespace.GetDefaultFolder(folder_id)
            items = folder.Items
            try:
                restricted = items.Restrict(query)
                try:
                    restricted.Sort(f"[{time_field}]", True)
                except Exception:
                    pass
                total += int(restricted.Count)
                for position, item in enumerate(restricted, start=1):
                    if len(rows) >= limit:
                        break
                    rows.append(_outlook_item_row(item, direction, time_field))
            except Exception:
                estimated = True
                try:
                    items.Sort(f"[{time_field}]", True)
                except Exception:
                    pass
                scanned = 0
                for item in items:
                    scanned += 1
                    if scanned > fallback_scan_limit:
                        break
                    try:
                        is_ours = PROCESSED_CATEGORY_NAME in _categories(item)
                        completed = int(getattr(item, "FlagStatus", 0) or 0) == 2
                    except Exception:
                        continue
                    if is_ours and completed:
                        total += 1
                        if len(rows) < limit:
                            rows.append(_outlook_item_row(item, direction, time_field))

        rows.sort(key=lambda r: r.get("Fecha", ""), reverse=True)
        return OutlookFlagSummary(total=total, rows=rows[:limit], estimated=estimated)
    except Exception as exc:
        return OutlookFlagSummary(0, [], error=str(exc))
    finally:
        pythoncom.CoUninitialize()
