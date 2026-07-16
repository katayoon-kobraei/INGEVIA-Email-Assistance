from src.classification.project_agent import classify_project

fake_email = {
    "subject": "Consulta sobre plan urbanístico",
    "sender": "cliente.prueba@example.com",
    "recipient": None,
    "direction": "ENTRANTE",
    "body": "Buenos días, quería preguntar sobre el estado del plan parcial de nuestro proyecto.",
}

result = classify_project(fake_email, existing_projects=[])
print(result)