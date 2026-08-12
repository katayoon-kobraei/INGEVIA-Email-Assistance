from google.genai import types
from src.gemini_client import generate_content_with_retry
from src.classification.schemas import DepartmentResult
from src.config import GEMINI_MODEL, PROMPTS_DIR

PROMPT_PATH = PROMPTS_DIR / "department_prompt.md"

def classify_department(email):
    # Read fresh on every call -- see the matching comment in
    # project_agent.py for why.
    department_rubric = PROMPT_PATH.read_text(encoding="utf-8")
    prompt = f"{department_rubric}\n\nSubject: {email['subject']}\nFrom: {email['sender']}\n\n{email['body']}"
    response = generate_content_with_retry(
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=DepartmentResult,
        ),
    )
    return DepartmentResult.model_validate_json(response.text)