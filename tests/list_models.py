from src.gemini_client import client

for model in client.models.list():
    print(model.name)