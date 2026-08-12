import os
from pathlib import Path
from dotenv import load_dotenv

APP_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(APP_ROOT / ".env")

# Keep the desktop application launchable before the user has entered a key.
# Gemini calls will still fail clearly until GEMINI_API_KEY is configured.
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = "gemini-3.5-flash"

# Installation role. PROCESSOR is the only mode allowed to connect to Outlook,
# run AI classification, write files, and register the scheduled task. VIEWER
# reads the shared output folder only.
APP_MODE = (os.environ.get("APP_MODE") or "PROCESSOR").strip().upper()
if APP_MODE not in {"PROCESSOR", "VIEWER"}:
    APP_MODE = "PROCESSOR"
IS_PROCESSOR = APP_MODE == "PROCESSOR"
IS_VIEWER = APP_MODE == "VIEWER"

# Shared application state used by the desktop UI and every viewer PC.
# This contains index.csv, priorities.csv, reports and dedupe/run state.
# It is intentionally separate from ARCHIVE_ROOT so company emails and
# attachments can continue to be filed in the real engineering archive.
# Example:
#   OUTPUT_ROOT=\\OFFICE-SERVER\EmailAssistantData
#
# NOTE on all os.environ.get(...) calls below: python-dotenv sets a
# variable to an EMPTY STRING if the .env line is present but has
# nothing after the "=" (e.g. "OUTPUT_ROOT="). That's different from
# the line being absent -- os.environ.get("X", default) only returns
# `default` when the key is missing entirely, not when it's "". So a
# blank value in your real .env would NOT fall back to the default --
# it would silently become "" (or False, for the on/off flag below).
# Using `os.environ.get("X") or default` instead makes an empty
# string behave the same as an unset/missing key.

OUTPUT_ROOT = os.environ.get("OUTPUT_ROOT") or r"C:\EmailAssistant\Output"

# Real engineering/project archive. The processing pipeline reads existing
# project folders from here and saves each email, PDF, MSG and attachment
# here. For INGEVIA this is normally the company P drive (preferably its
# UNC equivalent for scheduled processing), e.g. P:\ or \\SERVER\Projects.
# The existing structure is preserved: ARCHIVE_ROOT/TRABAJOS 2026/<project>.
ARCHIVE_ROOT = os.environ.get("ARCHIVE_ROOT") or OUTPUT_ROOT

# When true, every email the pipeline looks at gets stamped back in
# Outlook itself (a category + a completed flag) so the boss can see,
# right in Outlook, that the system has processed it -- whether it got
# filed into a project or was skipped as noise. Set to "false" in .env
# to turn this off without touching code.
FLAG_PROCESSED_EMAILS = (os.environ.get("FLAG_PROCESSED_EMAILS") or "true").lower() == "true"

# The Outlook category name used for the stamp above. Change it in
# .env if you'd rather it match an existing category the boss already
# uses. The color itself has to be assigned once inside Outlook
# (right-click any email > Categorize > All Categories) -- Outlook
# doesn't let code pick a color, only the category name.
PROCESSED_CATEGORY_NAME = os.environ.get("PROCESSED_CATEGORY_NAME") or "IA - PROCESADO"

# When true, emails classified as noise (social media, marketing,
# automated mail) get MOVED out of the Inbox into a dedicated
# subfolder instead of just being flagged in place -- keeps the real
# Inbox clean. Set to "false" in .env to go back to just flagging
# junk in place (no move).
ARCHIVE_JUNK_EMAILS = (os.environ.get("ARCHIVE_JUNK_EMAILS") or "true").lower() == "true"

# The subfolder junk gets moved into, created automatically under the
# Inbox the first time it's needed. Change it in .env if you'd rather
# reuse an existing folder name.
JUNK_ARCHIVE_FOLDER_NAME = os.environ.get("JUNK_ARCHIVE_FOLDER_NAME") or "Archivo"

# When true, every relevant, incoming (ENTRANTE) email additionally
# gets a cheap check for whether it's waiting on a reply. If so, it's
# moved into a dedicated Outlook subfolder and logged to pendientes.csv.
# Set to "false" in .env to turn this off.
CHECK_PENDING_RESPONSES = (os.environ.get("CHECK_PENDING_RESPONSES") or "true").lower() == "true"

# The subfolder emails needing a reply get moved into, created
# automatically under the Inbox the first time it's needed.
PENDING_FOLDER_NAME = os.environ.get("PENDING_FOLDER_NAME") or "PENDIENTE DE RESPUESTA"

# Path to the boss's reference spreadsheet (company + subfolder
# descriptions -- see project_descriptions.py). Leave unset to disable
# this entirely (classification just falls back to bare folder names,
# like before). Safe to point at a file that grows every year -- only
# the current email's own year ever gets read into a prompt, so
# adding historical years back to 2008 does not increase per-email
# token cost. e.g.:
#   DESCRIPTIONS_XLSX_PATH=P:\Trabajo IA.xlsx
DESCRIPTIONS_XLSX_PATH = os.environ.get("DESCRIPTIONS_XLSX_PATH") or ""

# Max characters kept per company/subfolder description once loaded
# from the spreadsheet -- this is the real lever on token cost, since
# it applies per-company regardless of how many years of history the
# spreadsheet holds. Cut at a sentence boundary when possible.
DESCRIPTION_MAX_CHARS = int(os.environ.get("DESCRIPTION_MAX_CHARS") or "220")

# Folder holding every classification agent's editable prompt file
# (project_prompt.md, address_prompt.md, billing_prompt.md,
# department_prompt.md, plenergy_address_prompt.md,
# post_filing_prompt.md). Defaults to this app's own src/prompts
# folder -- every agent below builds its own PROMPT_PATH from this
# single constant, so the desktop app's "Edit AI prompts" screen
# (see desktop_app/prompt_editor.py) always edits the exact files the
# pipeline actually reads, never a copy that silently falls out of
# sync. Only override this in .env if you specifically want prompts
# served from somewhere else.
PROMPTS_DIR = Path(os.environ.get("PROMPTS_DIR") or str(APP_ROOT / "src" / "prompts"))

BOSS_EMAIL = os.environ.get("BOSS_EMAIL") or "m.vera@ingevia.com"

# Outlook mailbox explicitly selected by processor installations. This avoids
# accidentally processing the Windows user's personal/default mailbox.
TARGET_MAILBOX = os.environ.get("TARGET_MAILBOX") or BOSS_EMAIL

# Human-readable alias for the shared application-data location. This is
# the same as OUTPUT_ROOT, while ARCHIVE_ROOT can point at P:\ separately.
SHARED_DATA_PATH = os.environ.get("SHARED_DATA_PATH") or OUTPUT_ROOT

ADMINISTRACION_EMAIL = os.environ.get("ADMINISTRACION_EMAIL") or "administracion@ingevia.com"

INTERNAL_DOMAIN = os.environ.get("INTERNAL_DOMAIN") or "ingevia.com"

BILLING_OUTPUT_ROOT = os.environ.get("BILLING_OUTPUT_ROOT") or r"X:\EMAILS"

IGNORE_DOMAINS = os.environ.get("IGNORE_DOMAINS") or "ingevia.com"

IGNORE_SENDERS = os.environ.get("IGNORE_SENDERS") or "fmunoz@munozbosch.com"

PLENERGY_SENDER_DOMAINS = os.environ.get("PLENERGY_SENDER_DOMAINS") or "plenergy.es,plainco.es"

# The MINIMUM the pipeline ever looks back for emails on each run.
# This MUST be longer than however often the Windows Scheduled Task
# actually runs the pipeline (see setup_scheduler.ps1 / Task
# Scheduler), or an email could land in the gap between two runs and
# be silently skipped. Defaults to 35 to safely cover the normal
# 30-minute schedule with a 5-minute margin for a run that starts a
# bit late. If you ever change how often the scheduled task runs
# during the day, update this to match (interval + 5-10 min margin) in
# .env, e.g.:
#   LOOKBACK_MINUTES=35
#
# Note: this is only a floor, not the actual window used on every run.
# src/pipeline.py (via src/output/run_state.py) remembers when it last
# ran successfully and automatically looks back further than this when
# the gap since that last run is bigger -- e.g. the first run of the
# day at 05:30 automatically covers the ~9.5-hour overnight gap since
# the previous day's last run at 20:00, without needing a separate
# setting for that.
LOOKBACK_MINUTES = int(os.environ.get("LOOKBACK_MINUTES") or "35")

def ensure_output_root():
    os.makedirs(OUTPUT_ROOT, exist_ok=True)