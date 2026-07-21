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
OUTPUT_ROOT = os.environ.get("OUTPUT_ROOT", r"C:\EmailAssistant\Output")

# Used by the read-only tools (eval_project_agent.py, build_archive_index.py)
# that only ever read .msg files and never write into the archive.
# Safe to point straight at the real archive even while OUTPUT_ROOT is
# still pointed at a test copy. Defaults to OUTPUT_ROOT if not set
# separately in .env.
ARCHIVE_ROOT = os.environ.get("ARCHIVE_ROOT", OUTPUT_ROOT)

# When true, every email the pipeline looks at gets stamped back in
# Outlook itself (a category + a completed flag) so the boss can see,
# right in Outlook, that the system has processed it -- whether it got
# filed into a project or was skipped as noise. Set to "false" in .env
# to turn this off without touching code.
FLAG_PROCESSED_EMAILS = os.environ.get("FLAG_PROCESSED_EMAILS", "true").lower() == "true"

# The Outlook category name used for the stamp above. Change it in
# .env if you'd rather it match an existing category the boss already
# uses. The color itself has to be assigned once inside Outlook
# (right-click any email > Categorize > All Categories) -- Outlook
# doesn't let code pick a color, only the category name.
PROCESSED_CATEGORY_NAME = os.environ.get("PROCESSED_CATEGORY_NAME", "IA - PROCESADO")


def ensure_output_root():
    os.makedirs(OUTPUT_ROOT, exist_ok=True)