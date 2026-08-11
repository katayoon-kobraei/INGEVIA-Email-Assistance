import time

from google import genai
from src.config import GEMINI_API_KEY

client = genai.Client(api_key=GEMINI_API_KEY)

MAX_RETRIES = 5
RETRY_DELAY_SECONDS = 10


def _is_transient(exc):
    """True for a temporary server-side overload (503 UNAVAILABLE) --
    worth waiting and trying again. False for anything else (bad
    request, auth failure, etc.), which won't fix itself by retrying."""
    if getattr(exc, "code", None) == 503:
        return True
    text = str(exc)
    return "503" in text or "UNAVAILABLE" in text or "overloaded" in text.lower()


def generate_content_with_retry(**kwargs):
    """Every agent should call this instead of client.models.generate_content
    directly. Retries with an increasing wait (10s, 20s, 30s, 40s...)
    when Gemini is overloaded, instead of failing the whole
    classification the moment a 503 shows up."""
    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return client.models.generate_content(**kwargs)
        except Exception as e:
            last_error = e
            if not _is_transient(e) or attempt == MAX_RETRIES:
                raise
            wait = RETRY_DELAY_SECONDS * attempt
            print(f"  Gemini overloaded (attempt {attempt}/{MAX_RETRIES}) -- retrying in {wait}s...")
            time.sleep(wait)
    raise last_error