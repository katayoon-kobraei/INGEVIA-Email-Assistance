import re

from src.config import DEPARTMENT_SECRETARIA_DOMAIN, DEPARTMENT_SECRETARIA_FOLDER_NAME

DEPARTMENT_ROUTES = []
if DEPARTMENT_SECRETARIA_DOMAIN:
    DEPARTMENT_ROUTES.append((DEPARTMENT_SECRETARIA_DOMAIN.strip().lower(), DEPARTMENT_SECRETARIA_FOLDER_NAME))


def _sender_domain(email):
    sender = email.get("sender") or ""
    match = re.search(r"@([\w.-]+)$", sender.strip())
    return match.group(1).lower() if match else None


def match_department(email):
    """Returns the department folder name if this email's sender
    domain matches a configured route, else None. Incoming mail
    only -- outgoing "sender" is the firm itself."""
    if email.get("direction") != "ENTRANTE":
        return None
    domain = _sender_domain(email)
    if not domain:
        return None
    for route_domain, folder_name in DEPARTMENT_ROUTES:
        if domain == route_domain:
            return folder_name
    return None