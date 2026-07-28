import os
from dotenv import load_dotenv

load_dotenv()  # reads .env and loads it into environment variables

GEMINI_API_KEY = os.environ["GEMINI_API_KEY"]
GEMINI_MODEL = "gemini-3.5-flash"

# The live pipeline reads existing project folders from here and
# writes new ones here. Override OUTPUT_ROOT in .env to point at a
# safe test copy while testing -- e.g.:
#   OUTPUT_ROOT=C:\EmailAssistant\TestOutput
# Once testing checks out, point it at the real archive for production:
#   OUTPUT_ROOT=P:\
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

# Used by the read-only tools (eval_project_agent.py, build_archive_index.py)
# that only ever read .msg files and never write into the archive.
# Safe to point straight at the real archive even while OUTPUT_ROOT is
# still pointed at a test copy. Defaults to OUTPUT_ROOT if not set
# separately in .env.
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

BOSS_EMAIL = os.environ.get("BOSS_EMAIL") or "m.vera@ingevia.com"

ADMINISTRACION_EMAIL = os.environ.get("ADMINISTRACION_EMAIL") or "administracion@ingevia.com"

INTERNAL_DOMAIN = os.environ.get("INTERNAL_DOMAIN") or "ingevia.com"

BILLING_OUTPUT_ROOT = os.environ.get("BILLING_OUTPUT_ROOT") or r"X:\EMAILS"

IGNORE_DOMAINS = os.environ.get("IGNORE_DOMAINS") or "ingevia.com"

IGNORE_SENDERS = os.environ.get("IGNORE_SENDERS") or "fmunoz@munozbosch.com"

PLENERGY_SENDER_DOMAINS = os.environ.get("PLENERGY_SENDER_DOMAINS") or "plenergy.es,plainco.es"

def ensure_output_root():
    os.makedirs(OUTPUT_ROOT, exist_ok=True)