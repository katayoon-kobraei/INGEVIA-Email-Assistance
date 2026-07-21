from pathlib import Path

from google.genai import types

from src.gemini_client import client
from src.classification.schemas import RelevanceResult
from src.config import GEMINI_MODEL

PROMPT_PATH = Path(__file__).resolve().parent.parent / "prompts" / "relevance_prompt.md"
RELEVANCE_RUBRIC = PROMPT_PATH.read_text(encoding="utf-8")


def classify_relevance(email):
    """Cheap first-pass filter: a short prompt with no folder
    structure, no matching rules, no naming conventions -- just
    "is this real correspondence or noise?". Meant to be called on
    every email BEFORE the expensive classify_project call, so junk
    (social media notifications, marketing, automated mail) never
    pays for the full classification prompt at all."""
    contact = email.get("sender") or email.get("recipient")
    prompt = (
        f"{RELEVANCE_RUBRIC}\n\n"
        f"Direction: {email['direction']}\nEmail subject: {email['subject']}\nContact: {contact}\n\n{email['body']}"
    )
    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=RelevanceResult,
        ),
    )
    return RelevanceResult.model_validate_json(response.text)