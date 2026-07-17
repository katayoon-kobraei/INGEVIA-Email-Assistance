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


def ensure_output_root():
    os.makedirs(OUTPUT_ROOT, exist_ok=True)