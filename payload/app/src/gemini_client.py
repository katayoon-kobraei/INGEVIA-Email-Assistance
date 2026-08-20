from __future__ import annotations

import time

from google import genai
from src.config import GEMINI_API_KEY, GEMINI_MIN_REQUEST_INTERVAL_SECONDS

client = genai.Client(api_key=GEMINI_API_KEY)

# Only small in-call retries are used for transient server/network problems.
# 429 is intentionally NOT retried here: the staged queue/circuit breaker owns
# that recovery so one rate-limit event cannot explode into a burst of calls.
MAX_TRANSIENT_RETRIES = 2
TRANSIENT_RETRY_DELAY_SECONDS = 10
_last_request_started = 0.0


class GeminiTemporaryUnavailable(RuntimeError):
    pass


class GeminiRateLimited(GeminiTemporaryUnavailable):
    pass


def _status_code(exc):
    for attr in ("code", "status_code"):
        value = getattr(exc, attr, None)
        try:
            if callable(value):
                value = value()
            if value is not None:
                return int(value)
        except Exception:
            pass
    return None


def is_rate_limit_error(exc) -> bool:
    code = _status_code(exc)
    if code == 429:
        return True
    text = str(exc).casefold()
    return "429" in text or "toomanyrequests" in text or "resource_exhausted" in text or "rate limit" in text


def is_transient_error(exc) -> bool:
    code = _status_code(exc)
    if code in {408, 429, 500, 502, 503, 504}:
        return True
    text = str(exc).casefold()
    markers = (
        "503", "unavailable", "overloaded", "timeout", "timed out",
        "connection reset", "temporarily unavailable", "resource_exhausted",
        "too many requests", "429",
    )
    return any(marker in text for marker in markers)


def _throttle() -> None:
    global _last_request_started
    interval = float(GEMINI_MIN_REQUEST_INTERVAL_SECONDS)
    if interval <= 0:
        _last_request_started = time.monotonic()
        return
    now = time.monotonic()
    remaining = interval - (now - _last_request_started)
    if _last_request_started > 0 and remaining > 0:
        time.sleep(remaining)
    _last_request_started = time.monotonic()


def generate_content_with_retry(**kwargs):
    """Central Gemini request gate used by every classifier.

    * serializes calls with a configurable minimum interval;
    * surfaces 429 immediately to the queue circuit breaker;
    * retries a small number of genuine transient 5xx/network failures;
    * never silently converts API failures into a classification result.
    """
    last_error = None
    for attempt in range(1, MAX_TRANSIENT_RETRIES + 1):
        _throttle()
        try:
            return client.models.generate_content(**kwargs)
        except Exception as exc:
            last_error = exc
            if is_rate_limit_error(exc):
                raise GeminiRateLimited(str(exc)) from exc
            if not is_transient_error(exc):
                raise
            if attempt >= MAX_TRANSIENT_RETRIES:
                raise GeminiTemporaryUnavailable(str(exc)) from exc
            wait = TRANSIENT_RETRY_DELAY_SECONDS * attempt
            print(f"  Gemini temporary error (attempt {attempt}/{MAX_TRANSIENT_RETRIES}) -- retrying in {wait}s...")
            time.sleep(wait)
    raise GeminiTemporaryUnavailable(str(last_error)) from last_error
