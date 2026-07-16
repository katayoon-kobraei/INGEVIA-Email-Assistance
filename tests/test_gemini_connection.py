from src.gemini_client import client
from src.config import GEMINI_MODEL

response = client.models.generate_content(
    model=GEMINI_MODEL,
    contents="Say hello in one word.",
)
print(response.text)