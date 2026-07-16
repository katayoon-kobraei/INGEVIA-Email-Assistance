import os
from dotenv import load_dotenv

load_dotenv()  # reads .env and loads it into environment variables

GEMINI_API_KEY = os.environ["GEMINI_API_KEY"]
OUTPUT_ROOT = r"C:\EmailAssistant\Output"
GEMINI_MODEL = "gemini-3.5-flash"

def ensure_output_root():
    os.makedirs(OUTPUT_ROOT, exist_ok=True)