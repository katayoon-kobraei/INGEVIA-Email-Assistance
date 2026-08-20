from __future__ import annotations

import csv
import gc
from email.header import decode_header, make_header
from email.utils import parseaddr
import json
import os
import re
import shutil
import subprocess
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path, PureWindowsPath
from typing import Any

from dotenv import load_dotenv

from src.output.outlook_visibility import (
    is_hidden_outlook_item,
    load_hidden_ids as _load_hidden_outlook_ids,
    sync_hidden_ids_from_outlook,
)
from src.dehu import is_dehu_sender


PROJECT_ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = PROJECT_ROOT / ".env"
load_dotenv(ENV_PATH)

APP_MODE = (os.environ.get("APP_MODE") or "PROCESSOR").strip().upper()
if APP_MODE not in {"PROCESSOR", "VIEWER"}:
    APP_MODE = "PROCESSOR"
IS_PROCESSOR = APP_MODE == "PROCESSOR"
IS_VIEWER = APP_MODE == "VIEWER"

OUTPUT_ROOT = Path(os.environ.get("OUTPUT_ROOT") or r"C:\EmailAssistant\Output")
SHARED_DATA_PATH = Path(os.environ.get("SHARED_DATA_PATH") or str(OUTPUT_ROOT))
ARCHIVE_ROOT = Path(os.environ.get("ARCHIVE_ROOT") or str(OUTPUT_ROOT))
PROJECT_YEAR_ROOT = Path(os.environ.get("PROJECT_YEAR_ROOT") or str(ARCHIVE_ROOT / f"TRABAJOS {datetime.now().year}"))
PROCESSED_CATEGORY_NAME = os.environ.get("PROCESSED_CATEGORY_NAME") or "IA - PROCESADO"
FLAG_PROCESSED_EMAILS = (os.environ.get("FLAG_PROCESSED_EMAILS") or "true").lower() == "true"
ARCHIVE_JUNK_EMAILS = (os.environ.get("ARCHIVE_JUNK_EMAILS") or "true").lower() == "true"
JUNK_ARCHIVE_FOLDER_NAME = os.environ.get("JUNK_ARCHIVE_FOLDER_NAME") or "Archivo"
CHECK_PENDING_RESPONSES = (os.environ.get("CHECK_PENDING_RESPONSES") or "true").lower() == "true"
PENDING_FOLDER_NAME = os.environ.get("PENDING_FOLDER_NAME") or "PENDIENTE DE RESPUESTA"
DESCRIPTIONS_XLSX_PATH = Path(os.environ.get("DESCRIPTIONS_XLSX_PATH") or "") if os.environ.get("DESCRIPTIONS_XLSX_PATH") else None
DESCRIPTION_MAX_CHARS = int(os.environ.get("DESCRIPTION_MAX_CHARS") or "220")
CURRENT_YEAR = datetime.now().year
BOSS_EMAIL = os.environ.get("BOSS_EMAIL") or "m.vera@ingevia.com"
TARGET_MAILBOX = os.environ.get("TARGET_MAILBOX") or BOSS_EMAIL
ADMINISTRACION_EMAIL = os.environ.get("ADMINISTRACION_EMAIL") or "administracion@ingevia.com"
BILLING_OUTPUT_ROOT = Path(os.environ.get("BILLING_OUTPUT_ROOT") or r"X:\EMAILS")
IGNORE_DOMAINS = os.environ.get("IGNORE_DOMAINS") or "ingevia.com"
IGNORE_SENDERS = os.environ.get("IGNORE_SENDERS") or "fmunoz@munozbosch.com"
PLENERGY_SENDER_DOMAINS = os.environ.get("PLENERGY_SENDER_DOMAINS") or "plenergy.es,plainco.es"
TASK_NAME = "INGEVIA Email AI Assistant"
MAIL_ITEM_CLASS = 43

# Same folder src/config.py's PROMPTS_DIR resolves to -- kept as an
# independent constant here (like PROJECT_YEAR_ROOT above) rather than
# importing src.config, since the desktop app must stay usable on a
# VIEWER install that never installs the Outlook/Gemini dependencies
# that package pulls in.
PROMPTS_DIR = Path(os.environ.get("PROMPTS_DIR") or str(PROJECT_ROOT / "src" / "prompts"))
PROMPT_FILES = [
    "project_prompt.md",
    "address_prompt.md",
    "plenergy_address_prompt.md",
    "billing_prompt.md",
    "department_prompt.md",
    "post_filing_prompt.md",
]

PROCESSED_IDS_PATH = OUTPUT_ROOT / "_processed_ids.json"
PENDING_LIST_PATH = OUTPUT_ROOT / "pendientes.csv"
PRIORITY_LIST_PATH = OUTPUT_ROOT / "priorities.csv"
REPORT_LOG_PATH = OUTPUT_ROOT / "email_report_log.csv"
REPORT_XLSX_PATH = OUTPUT_ROOT / "Informe de Emails.xlsx"


def _hidden_process_options() -> dict[str, Any]:
    """Run Windows helper commands without flashing a console window."""
    if os.name != "nt":
        return {}

    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = subprocess.SW_HIDE
    return {
        "startupinfo": startupinfo,
        "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0),
    }


def list_prompt_files() -> list[Path]:
    """Every editable AI prompt file, in a fixed, sensible order (not
    alphabetical). Falls back to whatever .md files actually exist in
    PROMPTS_DIR if that doesn't match the expected set -- so a renamed
    or newly-added prompt file still shows up instead of silently
    disappearing from the editor."""
    if not PROMPTS_DIR.is_dir():
        return []
    ordered = [PROMPTS_DIR / name for name in PROMPT_FILES if (PROMPTS_DIR / name).is_file()]
    known = {p.name for p in ordered}
    extra = sorted(p for p in PROMPTS_DIR.glob("*.md") if p.name not in known)
    return ordered + extra


def read_prompt_file(path: Path) -> str:
    return Path(path).read_text(encoding="utf-8")


def write_prompt_file(path: Path, content: str) -> tuple[bool, str]:
    """Saves new prompt content. Keeps a timestamped backup of the
    previous version first, under PROMPTS_DIR/_backups, so an
    accidental bad edit is always recoverable -- these files directly
    control how Gemini classifies and files every email. Returns
    (True, "") on success or (False, error_message) on failure."""
    path = Path(path)
    try:
        if path.is_file():
            backup_dir = PROMPTS_DIR / "_backups"
            backup_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            shutil.copy2(path, backup_dir / f"{path.stem}.{stamp}{path.suffix}")
        path.write_text(content, encoding="utf-8")
        return True, ""
    except OSError as exc:
        return False, str(exc)


# Column order project_descriptions.py reads by POSITION (not header
# text) from every sheet of DESCRIPTIONS_XLSX_PATH -- see that module's
# docstring. Kept here as documentation; save_folder_data_row() below
# writes new/updated rows in this exact order.
FOLDER_DATA_COLUMNS = [
    "ID proyecto",
    "Empresa\\ Proyecto",
    "Dirección",
    "Año",
    "Empresa\\ Proyecto Descripción",
    "Subfolder Descripción",
]


def save_folder_data_row(
    project_id: str,
    company: str,
    address: str,
    year: Any,
    company_description: str = "",
    subfolder_description: str = "",
) -> tuple[bool, str, bool]:
    """Adds one row to DESCRIPTIONS_XLSX_PATH (Folder_Data.xlsx), the
    reference spreadsheet project_descriptions.py reads for
    company/site context on every classification call -- see that
    module for how it's consumed. Changes take effect on the very next
    email processed (that module reloads the workbook whenever its
    mtime changes), no restart needed.

    Matches an existing row by 'ID proyecto' (column A, trimmed,
    case-insensitive) across every sheet in the workbook and updates it
    in place instead of creating a duplicate. A brand-new ID is
    appended to the sheet named after `year` if one exists, otherwise
    to the workbook's last sheet (today that's the only sheet, "Trabajo
    IA", which already mixes every year).

    Returns (ok, message, was_update). Messages are bilingual (ES / EN)
    since this module is shared by both language UI windows -- see the
    same pattern in open_path() above.
    """
    import openpyxl  # local import: keeps this dependency out of every

    # other desktop_app codepath that doesn't touch spreadsheets.

    path = DESCRIPTIONS_XLSX_PATH
    if not path:
        return False, (
            "No hay ningún archivo de descripciones configurado (DESCRIPTIONS_XLSX_PATH).\n"
            "No description spreadsheet is configured (DESCRIPTIONS_XLSX_PATH)."
        ), False
    if not path.exists():
        return False, f"No se encontró el archivo / File not found:\n{path}", False

    project_id = (project_id or "").strip()
    if not project_id:
        return False, "Falta el ID de proyecto / Missing project ID.", False

    try:
        year_value: Any = int(str(year).strip())
    except (TypeError, ValueError):
        year_value = str(year).strip()

    new_values = [
        project_id,
        (company or "").strip(),
        (address or "").strip(),
        year_value,
        (company_description or "").strip(),
        (subfolder_description or "").strip(),
    ]

    try:
        wb = openpyxl.load_workbook(path)
    except PermissionError:
        return False, (
            "No se pudo guardar: el archivo está abierto en Excel. Ciérrelo e inténtelo de nuevo.\n"
            "Could not save: the file is open in Excel. Close it and try again."
        ), False
    except OSError as exc:
        return False, str(exc), False

    def _do_save() -> tuple[bool, str]:
        try:
            wb.save(path)
            return True, ""
        except PermissionError:
            return False, (
                "No se pudo guardar: el archivo está abierto en Excel. Ciérrelo e inténtelo de nuevo.\n"
                "Could not save: the file is open in Excel. Close it and try again."
            )
        except OSError as exc:
            return False, str(exc)

    target_key = project_id.casefold()
    for sheet in wb.worksheets:
        for (cell,) in sheet.iter_rows(min_row=2, min_col=1, max_col=1):
            if cell.value is None:
                continue
            if str(cell.value).strip().casefold() != target_key:
                continue
            for col_offset, value in enumerate(new_values[1:], start=2):
                sheet.cell(row=cell.row, column=col_offset, value=value)
            ok, error = _do_save()
            if not ok:
                return False, error, False
            return True, f"Fila actualizada en la hoja «{sheet.title}». / Row updated in sheet “{sheet.title}”.", True

    year_sheet_name = str(year).strip()
    target_sheet = wb[year_sheet_name] if year_sheet_name in wb.sheetnames else wb[wb.sheetnames[-1]]
    target_sheet.append(new_values)
    ok, error = _do_save()
    if not ok:
        return False, error, False
    return True, f"Fila añadida a la hoja «{target_sheet.title}». / Row added to sheet “{target_sheet.title}”.", False


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


def _normalized_subject(subject: str | None) -> str:
    return " ".join(str(subject or "").split()).casefold()


def _normalized_date(date_text: str | None) -> str:
    raw = str(date_text or "").strip()
    parsed = _parse_date(raw)
    return parsed.strftime("%Y-%m-%d %H:%M") if parsed else raw


def _email_key(date_text: str | None, subject: str | None) -> tuple[str, str]:
    return (_normalized_date(date_text), _normalized_subject(subject))


@dataclass(slots=True)
class PriorityLookup:
    by_entry_id: dict[str, int]
    by_key: dict[tuple[str, str], int]
    by_subject: dict[str, list[tuple[datetime | None, int]]]

    def resolve(
        self,
        entry_id: str | None,
        date_text: str | None,
        subject: str | None,
    ) -> int:
        stable_id = str(entry_id or "").strip()
        if stable_id and stable_id in self.by_entry_id:
            return self.by_entry_id[stable_id]

        key = _email_key(date_text, subject)
        if key in self.by_key:
            return self.by_key[key]

        # Legacy index/report rows may not contain EntryID. Match the same
        # normalized subject to the closest priority timestamp within 10 min.
        candidates = self.by_subject.get(key[1], [])
        target = _parse_date(key[0])
        if target and candidates:
            timed = [
                (abs((candidate_date - target).total_seconds()), value)
                for candidate_date, value in candidates
                if candidate_date is not None
            ]
            if timed:
                seconds, value = min(timed, key=lambda item: item[0])
                if seconds <= 600:
                    return value
        if len(candidates) == 1:
            return candidates[0][1]
        return 0


def load_ui_hidden_ids() -> set[str]:
    """IDs that the processing PC has identified as Outlook Other/Otros.

    Stored under OUTPUT_ROOT so viewer PCs apply exactly the same UI filter
    without opening the boss mailbox.
    """
    return _load_hidden_outlook_ids(OUTPUT_ROOT)


def _row_is_ui_hidden(row: dict[str, Any], hidden_ids: set[str] | None = None) -> bool:
    # Official DEHU notifications are business-critical. v1.25 could persist
    # their EntryID in _ui_hidden_outlook_ids.json when Outlook classified the
    # message as Other/Otros. Even after v1.26 stopped hiding new DEHU mail,
    # that stale shared ID could keep an already-processed DEHU row invisible
    # on every processor/viewer UI. The index sender is authoritative here, so
    # DEHU is always visible without requiring another Outlook COM lookup.
    sender = (
        row.get("Sender Email")
        or row.get("SenderEmail")
        or row.get("Sender/Recipient")
        or ""
    )
    if is_dehu_sender(sender):
        return False
    hidden = hidden_ids if hidden_ids is not None else load_ui_hidden_ids()
    entry_id = str(row.get("EntryID") or row.get("Entry ID") or "").strip()
    return bool(entry_id and entry_id in hidden)


def _missing_holding_pen_path(row: dict[str, Any]) -> bool:
    """Hide stale YY-000 MAILS rows whose archived folder no longer exists.

    This is a narrow safety net for legacy advertising/unsorted rows that were
    later removed or reclassified. Valid 26-000 MAILS entries remain visible.
    Project rows outside the holding pen are never hidden by this check.
    """
    raw = str(row.get("Folder Path") or row.get("Path") or row.get("Ruta") or "").strip()
    if not raw or not re.search(r"(?:^|[\\/])\d{2}-000\s+MAILS(?:[\\/]|$)", raw, re.IGNORECASE):
        return False
    return resolve_open_path(raw) is None


def _sync_ui_hidden_ids_from_outlook() -> None:
    """Refresh the shared Other/Otros exclusion list on the processor PC.

    This runs inside the existing background UI loader, so COM is initialized
    explicitly and the desktop stays responsive. Viewer installations never
    connect to Outlook and simply read the shared JSON generated here or by
    the scheduled pipeline.
    """
    if not IS_PROCESSOR or os.name != "nt":
        return
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return

    pythoncom.CoInitialize()
    try:
        namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        sync_hidden_ids_from_outlook(OUTPUT_ROOT, namespace)
    except Exception:
        # A temporary Outlook problem should never prevent the rest of the UI
        # from loading; the last successful shared exclusion list remains.
        pass
    finally:
        pythoncom.CoUninitialize()


def load_processed_ids() -> set[str]:
    try:
        if PROCESSED_IDS_PATH.exists():
            data = json.loads(PROCESSED_IDS_PATH.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return {str(item) for item in data if item}
    except (OSError, json.JSONDecodeError):
        pass
    return set()


def load_priority_lookup() -> PriorityLookup:
    by_entry_id: dict[str, int] = {}
    by_key: dict[tuple[str, str], int] = {}
    by_subject: dict[str, list[tuple[datetime | None, int]]] = {}
    for row in _read_csv(PRIORITY_LIST_PATH):
        try:
            priority = int(row.get("Priority") or row.get("Prioridad") or 0)
        except (TypeError, ValueError):
            continue
        if not 1 <= priority <= 5:
            continue

        entry_id = str(row.get("EntryID") or row.get("Entry ID") or "").strip()
        date_text = row.get("Date") or row.get("Fecha")
        subject = row.get("Subject") or row.get("Asunto")
        key = _email_key(date_text, subject)
        if entry_id:
            by_entry_id[entry_id] = priority
        by_key[key] = priority
        by_subject.setdefault(key[1], []).append((_parse_date(key[0]), priority))

    return PriorityLookup(by_entry_id, by_key, by_subject)


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


def sender_display(record: dict[str, Any]) -> str:
    """Format the actual sender as ``Name <email>`` for the Correos table.

    New v1.14 records persist the Outlook display name and SMTP address in
    ``index.csv``. Older rows fall back to the report/contact data so existing
    shared archives remain readable without migration.
    """
    name = str(
        record.get("SenderName")
        or record.get("Sender Name")
        or record.get("NombreRemitente")
        or ""
    ).strip()
    email = str(
        record.get("SenderEmail")
        or record.get("Sender Email")
        or record.get("EmailRemitente")
        or record.get("Sender/Recipient")
        or ""
    ).strip()

    # Some Outlook providers expose SenderName as an RFC mailbox string.
    # Normalize it so the Correos column always renders one clean
    # ``Display Name <smtp@address>`` value instead of duplicating the address.
    if name:
        parsed_name, parsed_email = parseaddr(name)
        if parsed_email:
            if not email:
                email = parsed_email
            if parsed_name:
                try:
                    name = str(make_header(decode_header(parsed_name))).strip()
                except Exception:
                    name = parsed_name.strip()
            elif name.casefold() == parsed_email.casefold():
                name = ""

    if name and email and name.casefold() != email.casefold():
        return f"{name} <{email}>"
    return name or email


def _is_saliente(row: dict[str, Any]) -> bool:
    """Compatibility helper for callers that need to identify outgoing rows."""
    return str(row.get("Direction") or row.get("Dirección") or "").strip().upper() == "SALIENTE"


def load_index_rows() -> list[dict[str, Any]]:
    hidden_ids = load_ui_hidden_ids()
    rows = [
        row for row in _read_csv(OUTPUT_ROOT / "index.csv")
        if (
            not _is_saliente(row)
            and not _row_is_ui_hidden(row, hidden_ids)
            and not _missing_holding_pen_path(row)
        )
    ]
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
        priority = priority_lookup.resolve(row.get("EntryID"), date_text, subject)
        row["Priority"] = priority if priority else ""
        row["_priority"] = priority
        report_row = report_lookup.get(key, {})
        processing_status = str(row.get("Processing Status") or "").strip().upper()
        if processing_status == "PENDING_AI":
            row["Summary"] = "Pendiente de procesamiento por IA"
            row["_summary"] = row["Summary"]
        elif processing_status == "AI_REVIEW":
            row["Summary"] = "Procesamiento IA fallido repetidamente; revisar"
            row["_summary"] = row["Summary"]
        else:
            row["Summary"] = report_row.get("Summary", "")
            row["_summary"] = report_row.get("Summary", "")
        # v1.14 stores the real Outlook sender in index.csv. For rows written
        # by older versions, fall back to the report and then to the existing
        # Sender/Recipient value so the new UI column is immediately useful.
        row["SenderName"] = (
            row.get("Sender Name")
            or report_row.get("SenderName")
            or ""
        )
        row["SenderEmail"] = (
            row.get("Sender Email")
            or report_row.get("SenderEmail")
            or (row.get("Sender/Recipient") if row.get("Direction") == "ENTRANTE" else "")
            or ""
        )
        row["SenderDisplay"] = sender_display(row)
        row["_pending"] = key in pending_lookup
        processing_status = str(row.get("Processing Status") or "").strip().upper()
        if processing_status in {"PENDING_AI", "AI_REVIEW"}:
            row["_status"] = processing_status
        else:
            row["_status"] = "REVISAR" if row.get("Project Folder") == "UNSORTED" else "PROCESADO"

    rows.sort(key=lambda row: row.get("_parsed_date") or datetime.min, reverse=True)
    return rows


def load_report_rows() -> list[dict[str, Any]]:
    priority_lookup = load_priority_lookup()
    pending_lookup = load_pending_lookup()
    hidden_ids = load_ui_hidden_ids()

    # v1.26.3 is incoming-only. Older report_log rows do not always carry a
    # Direction field, so use index.csv as the authority for identifying and
    # hiding any historical SALIENTE record as well.
    legacy_saliente_ids: set[str] = set()
    legacy_saliente_keys: set[tuple[str, str]] = set()
    for index_row in _read_csv(OUTPUT_ROOT / "index.csv"):
        if not _is_saliente(index_row):
            continue
        entry_id = str(index_row.get("EntryID") or "").strip()
        if entry_id:
            legacy_saliente_ids.add(entry_id)
        legacy_saliente_keys.add(
            _email_key(index_row.get("Date"), index_row.get("Subject"))
        )

    rows: list[dict[str, Any]] = []
    for row in _read_csv(REPORT_LOG_PATH):
        report_entry_id = str(row.get("EntryID") or "").strip()
        report_key = _email_key(row.get("Date"), row.get("Subject"))
        if (
            _is_saliente(row)
            or (report_entry_id and report_entry_id in legacy_saliente_ids)
            or report_key in legacy_saliente_keys
            or _row_is_ui_hidden(row, hidden_ids)
            or _missing_holding_pen_path(row)
        ):
            continue
        date_text = " ".join(
            part for part in (str(row.get("Date") or ""), str(row.get("Time") or "")) if part
        ).strip()
        subject = str(row.get("Subject") or "")
        key = _email_key(date_text, subject)
        priority = priority_lookup.resolve(row.get("EntryID"), date_text, subject)
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
    hidden_ids = load_ui_hidden_ids()
    rows = [
        row for row in _read_csv(OUTPUT_ROOT / "activity_log.csv")
        if not _row_is_ui_hidden(row, hidden_ids)
    ]
    if not rows:
        return load_index_rows()
    priority_lookup = load_priority_lookup()
    pending_lookup = load_pending_lookup()
    for row in rows:
        date_text = str(row.get("Date") or "")
        subject = str(row.get("Subject") or "")
        key = _email_key(date_text, subject)
        priority = priority_lookup.resolve(row.get("EntryID"), date_text, subject)
        row["_parsed_date"] = _parse_date(date_text)
        row["Priority"] = priority if priority else ""
        row["_priority"] = priority
        row["_pending"] = key in pending_lookup
    rows.sort(key=lambda row: row.get("_parsed_date") or datetime.min, reverse=True)
    return rows


def load_pending_rows() -> list[dict[str, Any]]:
    priority_lookup = load_priority_lookup()
    hidden_ids = load_ui_hidden_ids()
    normalized: list[dict[str, Any]] = []
    for row in _read_csv(PENDING_LIST_PATH):
        if _row_is_ui_hidden(row, hidden_ids):
            continue
        date = str(row.get("Date") or "")
        sender = str(row.get("Sender") or "")
        subject = str(row.get("Subject") or "")
        priority = priority_lookup.resolve(row.get("EntryID"), date, subject)
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
    rows = load_index_rows()
    # Count finalized EMAILS, not archive rows. One email can legitimately be
    # copied to multiple Plenergy/project destinations and therefore have more
    # than one index row with the same EntryID.
    finalized_ids = {
        str(row.get("EntryID") or "").strip()
        for row in rows
        if row.get("_status") not in {"PENDING_AI", "AI_REVIEW"}
        and str(row.get("EntryID") or "").strip()
    }
    return len(finalized_ids)


def load_attachment_rows() -> list[dict[str, Any]]:
    """Load attachments from folders referenced by the shared index.

    ARCHIVE_ROOT may be the whole P drive, so recursively scanning it would
    make every viewer slow.  Instead, use index.csv as the authoritative list
    of Email Assistant folders and inspect only those metadata.json files.
    """
    rows: list[dict[str, Any]] = []
    hidden_ids = load_ui_hidden_ids()
    for index_row in _read_csv(OUTPUT_ROOT / "index.csv"):
        if _is_saliente(index_row) or _row_is_ui_hidden(index_row, hidden_ids):
            continue
        raw_folder = str(index_row.get("Folder Path") or "").strip()
        if not raw_folder:
            continue
        folder = resolve_open_path(raw_folder)
        if folder is None:
            continue
        metadata_path = folder / "metadata.json"
        if not metadata_path.is_file():
            continue
        try:
            metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue

        for attachment in metadata.get("attachments", []):
            filename = str(attachment.get("filename", ""))
            status = str(attachment.get("status", "unknown"))
            reason = str(attachment.get("reason", ""))
            path = folder / filename
            date = metadata.get("timestamp", index_row.get("Date", ""))
            project = metadata.get("project_folder", index_row.get("Project Folder", ""))
            subject = metadata.get("subject", index_row.get("Subject", ""))
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



def get_report_excel_info(entries: int | None = None) -> dict[str, Any]:
    """Return status information for the generated Excel email report."""
    path = REPORT_XLSX_PATH
    info: dict[str, Any] = {
        "path": str(path),
        "name": path.name,
        "exists": False,
        "size_bytes": 0,
        "size_text": "—",
        "modified": "—",
        "entries": len(load_report_rows()) if entries is None else entries,
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


def load_ui_snapshot() -> dict[str, Any]:
    """Load all local/shared UI data in one background-friendly operation."""
    if IS_VIEWER and not OUTPUT_ROOT.is_dir():
        raise RuntimeError(
            "The shared Email Assistant data folder is unavailable. "
            f"Check the office network connection and access permissions:\n{OUTPUT_ROOT}"
        )
    # Processor refreshes the shared Other/Otros exclusion state before
    # loading tables; viewers consume that same state from OUTPUT_ROOT.
    rows = load_index_rows()
    activity_log = OUTPUT_ROOT / "activity_log.csv"
    activity_rows = load_activity_rows() if activity_log.exists() else rows
    attachments = load_attachment_rows()
    pending_rows = load_pending_rows()
    report_rows = load_report_rows()
    processed = processed_ids_count()
    excel_info = get_report_excel_info(entries=len(report_rows))
    return {
        "rows": rows,
        "activity_rows": activity_rows,
        "attachments": attachments,
        "pending_rows": pending_rows,
        "report_rows": report_rows,
        "processed": processed,
        "excel_info": excel_info,
    }

def get_scheduler_status() -> dict[str, Any] | None:
    if IS_VIEWER:
        return {"Mode": "VIEWER"}
    if os.name != "nt":
        return None
    try:
        command = (
            f"Get-ScheduledTaskInfo -TaskName '{TASK_NAME}' | "
            "Select-Object LastRunTime,LastTaskResult,NextRunTime | ConvertTo-Json"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
            capture_output=True,
            text=True,
            stdin=subprocess.DEVNULL,
            timeout=10,
            check=False,
            **_hidden_process_options(),
        )
        if result.returncode != 0 or not result.stdout.strip():
            return None
        return json.loads(result.stdout)
    except (OSError, subprocess.SubprocessError, json.JSONDecodeError):
        return None


def _shared_path_candidates(path: str | Path) -> list[Path]:
    r"""Return safe candidates for old and new saved folder paths.

    New processing records store the actual ARCHIVE_ROOT path.  Legacy records
    may contain ``Data\TRABAJOS...``, an old local ``...\app\Data`` path,
    or a mapped-drive path from the processor.  Viewer PCs translate the
    ``TRABAJOS <year>`` suffix to their configured ARCHIVE_ROOT.  As a final
    compatibility fallback, the old SHARED_DATA_PATH location is also tried.
    """
    raw = str(path or "").strip().strip('"').strip("'")
    if not raw:
        return []

    candidates: list[Path] = [Path(raw)]
    win_path = PureWindowsPath(raw.replace("/", "\\"))
    parts = list(win_path.parts)

    # Find the archive-relative suffix beginning at TRABAJOS <year>. This works
    # for relative Data paths, boss-local paths and mapped-drive paths alike.
    archive_index = None
    for i, part in enumerate(parts):
        if re.match(r"^TRABAJOS \d{4}$", str(part), flags=re.IGNORECASE):
            archive_index = i
            break
    if archive_index is not None:
        suffix = [str(part).strip("\\/") for part in parts[archive_index:] if str(part).strip("\\/")]
        if suffix:
            candidates.append(ARCHIVE_ROOT.joinpath(*suffix))
            # v1.7-v1.10 stored archive content inside the shared data folder.
            candidates.append(SHARED_DATA_PATH.joinpath(*suffix))

    # Legacy relative paths without a TRABAJOS segment still get the old shared
    # data fallback (for reports/config files rather than archived emails).
    if not win_path.is_absolute():
        clean = [str(part).strip("\\/") for part in parts if str(part).strip("\\/")]
        aliases = {"data", "output", SHARED_DATA_PATH.name.casefold()}
        if clean and clean[0].casefold() in aliases:
            clean = clean[1:]
        if clean:
            candidates.append(SHARED_DATA_PATH.joinpath(*clean))

    unique: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        key = str(candidate).casefold()
        if key not in seen:
            unique.append(candidate)
            seen.add(key)
    return unique


def resolve_open_path(path: str | Path) -> Path | None:
    for candidate in _shared_path_candidates(path):
        try:
            if candidate.exists():
                return candidate
        except OSError:
            continue
    return None


def open_path(path: str | Path) -> tuple[bool, str]:
    if not path:
        return False, "No path was configured / No se configuró ninguna ruta."

    target = resolve_open_path(path)
    if target is None:
        original = str(path)
        shared_hint = str(SHARED_DATA_PATH)
        archive_hint = str(ARCHIVE_ROOT)
        return False, (
            "Path does not exist / No existe la ruta:\n"
            f"{original}\n\n"
            "Project archive / Archivo de proyectos:\n"
            f"{archive_hint}\n\n"
            "Shared app data / Datos compartidos de la aplicación:\n"
            f"{shared_hint}"
        )
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



def open_email_record(record: dict[str, Any]) -> tuple[bool, str]:
    """Open the email's CURRENT shared-index folder, not a stale UI snapshot.

    The AI queue can move a captured email from ``Unprocessed`` to its final
    archive between two UI refreshes. If the row on screen still contains the
    old staging path, resolve the same EntryID again from index.csv and open the
    newest existing path instead of showing a false "path does not exist" error.
    """
    raw_path = str(record.get("Folder Path") or "").strip()
    ok, message = open_path(raw_path)
    if ok:
        return ok, message

    entry_id = str(record.get("EntryID") or "").strip()
    if not entry_id:
        return ok, message
    for current in _read_csv(OUTPUT_ROOT / "index.csv"):
        if str(current.get("EntryID") or "").strip() != entry_id:
            continue
        current_path = str(current.get("Folder Path") or "").strip()
        if not current_path or current_path == raw_path:
            continue
        current_ok, current_message = open_path(current_path)
        if current_ok:
            return current_ok, current_message
    return ok, message

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


def _load_shared_processed_summary(limit: int = 250) -> OutlookFlagSummary:
    """Build the viewer's processed-email view from the shared CSV records."""
    if not OUTPUT_ROOT.is_dir():
        return OutlookFlagSummary(
            0, [], error=f"Shared data folder is unavailable: {OUTPUT_ROOT}"
        )

    source_rows = load_index_rows()
    if not source_rows:
        source_rows = load_report_rows()

    rows: list[dict[str, Any]] = []
    for record in source_rows[:limit]:
        date_text = str(
            record.get("Date")
            or record.get("DateTime")
            or record.get("FechaHora")
            or ""
        )
        direction = str(record.get("Direction") or record.get("Dirección") or "")
        contact = str(
            record.get("Sender Email")
            or record.get("SenderEmail")
            or record.get("Sender")
            or record.get("Contact")
            or ""
        )
        subject = str(record.get("Subject") or record.get("Asunto") or "")
        rows.append(
            {
                "Fecha": date_text,
                "Dirección": direction,
                "Contacto": contact,
                "Asunto": subject,
                "Categoría": PROCESSED_CATEGORY_NAME,
                "Date": date_text,
                "Direction": direction,
                "Contact": contact,
                "Subject": subject,
                "Category": PROCESSED_CATEGORY_NAME,
                "EntryID": "",
            }
        )

    total = processed_ids_count()
    return OutlookFlagSummary(total=total, rows=rows, estimated=True)


def _target_mailbox_store_ids(namespace: Any) -> list[str]:
    """Return StoreID values for the explicitly configured boss mailbox.

    EntryID alone can fail when the mailbox is an additional/shared Outlook
    store. The flagged-email page must therefore resolve items using the same
    mailbox store that the processing pipeline used.
    """
    store_ids: list[str] = []
    try:
        from src.outlook_mailbox import get_target_folder
    except Exception:
        return store_ids

    for folder_id in (6,):  # Inbox only; Sent Items are intentionally ignored.
        try:
            folder = get_target_folder(namespace, folder_id, TARGET_MAILBOX)
            store_id = str(_safe_get(folder, "StoreID", "") or "").strip()
            if store_id and store_id not in store_ids:
                store_ids.append(store_id)
        except Exception:
            continue
    return store_ids


def _get_processed_outlook_item(namespace: Any, entry_id: str, store_ids: list[str]) -> Any:
    """Resolve one processed Outlook item across primary/shared stores."""
    for store_id in store_ids:
        try:
            return namespace.GetItemFromID(entry_id, store_id)
        except Exception:
            continue
    try:
        return namespace.GetItemFromID(entry_id)
    except Exception:
        return None


def load_outlook_flagged(limit: int = 250) -> OutlookFlagSummary:
    """Show AI-processed messages that are currently flagged in Outlook.

    Instead of calling GetItemFromID once for every processed email, query the
    Inbox for FlagStatus=2 first and intersect that much smaller set with the AI
    processed-ID set. Sent Items are intentionally never queried.
    """
    if IS_VIEWER:
        return _load_shared_processed_summary(limit)
    if os.name != "nt":
        return OutlookFlagSummary(0, [], error="Outlook is only available on Windows.")

    processed_ids = load_processed_ids()
    if not processed_ids:
        return OutlookFlagSummary(0, [])

    # The dedupe state also contains messages deliberately ignored by the app
    # (for example internal @ingevia.com mail). v1.26.3 can mark those messages
    # in Outlook, but they must stay invisible everywhere in the application.
    # Restrict the Flagged page to EntryIDs that actually have an app UI/report
    # record; internal/blocked/no-attachment ignores never receive such a row.
    ui_processed_ids: set[str] = set()
    for record in load_index_rows() + load_report_rows():
        entry_id = str(record.get("EntryID") or "").strip()
        if entry_id:
            ui_processed_ids.add(entry_id)
    processed_ids = processed_ids.intersection(ui_processed_ids)
    if not processed_ids:
        return OutlookFlagSummary(0, [])

    try:
        import pythoncom
        import win32com.client
        from src.outlook_mailbox import get_target_folder
    except ImportError as exc:
        return OutlookFlagSummary(0, [], error=f"pywin32 missing: {exc}")

    pythoncom.CoInitialize()
    namespace = None
    try:
        namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        rows: list[dict[str, Any]] = []
        total = 0
        successful_restrict_folders = 0

        for folder_id, direction, time_field in (
            (6, "ENTRANTE", "ReceivedTime"),
        ):
            try:
                folder = get_target_folder(namespace, folder_id, TARGET_MAILBOX)
                source_items = folder.Items
                flagged_items = source_items.Restrict("[FlagStatus] = 2")
                try:
                    flagged_items.Sort(f"[{time_field}]", True)
                except Exception:
                    pass
                successful_restrict_folders += 1
                count = int(_safe_get(flagged_items, "Count", 0) or 0)
                for index in range(1, count + 1):
                    item = None
                    try:
                        item = flagged_items.Item(index)
                        if _safe_get(item, "Class") != MAIL_ITEM_CLASS:
                            continue
                        entry_id = str(_safe_get(item, "EntryID", "") or "").strip()
                        if not entry_id or entry_id not in processed_ids:
                            continue
                        if is_hidden_outlook_item(item):
                            continue
                        total += 1
                        if len(rows) < limit:
                            rows.append(_outlook_item_row(item, direction))
                    finally:
                        if item is not None:
                            del item
                del flagged_items
                del source_items
                del folder
            except Exception:
                continue

        # Rare provider fallback: if Inbox does not accept the FlagStatus
        # Restrict filter, resolve processed EntryIDs directly. Sent messages
        # are explicitly discarded even in this fallback path.
        if successful_restrict_folders == 0:
            store_ids = _target_mailbox_store_ids(namespace)
            for entry_id in processed_ids:
                item = _get_processed_outlook_item(namespace, entry_id, store_ids)
                if item is None:
                    continue
                try:
                    if _safe_get(item, "Class") != MAIL_ITEM_CLASS:
                        continue
                    if is_hidden_outlook_item(item):
                        continue
                    categories = _categories(item)
                    flag_status = int(_safe_get(item, "FlagStatus", 0) or 0)
                    if PROCESSED_CATEGORY_NAME not in categories and flag_status != 2:
                        continue
                    if bool(_safe_get(item, "Sent", False)):
                        continue
                    total += 1
                    if len(rows) < limit:
                        rows.append(_outlook_item_row(item, "ENTRANTE"))
                finally:
                    del item

        rows.sort(key=lambda row: row.get("Fecha", ""), reverse=True)
        return OutlookFlagSummary(total=total, rows=rows[:limit], estimated=False)
    except Exception as exc:
        return OutlookFlagSummary(0, [], error=str(exc))
    finally:
        try:
            if namespace is not None:
                del namespace
        except Exception:
            pass
        gc.collect()
        pythoncom.CoUninitialize()
