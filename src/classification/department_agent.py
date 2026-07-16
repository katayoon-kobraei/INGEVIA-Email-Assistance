from pathlib import Path
from google.genai import types
from src.gemini_client import client
from src.classification.schemas import DepartmentResult
from src.config import GEMINI_MODEL

PROMPT_PATH = Path(__file__).resolve().parent.parent / "prompts" / "department_prompt.md"
DEPARTMENT_RUBRIC = PROMPT_PATH.read_text(encoding="utf-8")

def classify_department(email):
    prompt = f"{DEPARTMENT_RUBRIC}\n\nSubject: {email['subject']}\nFrom: {email['sender']}\n\n{email['body']}"
    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=DepartmentResult,
        ),
    )
    return DepartmentResult.model_validate_json(response.text)