def build_conversation_folder_name(email, contact_label, topic_label):
    date_str = email["timestamp"].strftime("%y-%m-%d")
    return f"{date_str} {contact_label}_{topic_label}"


def build_holding_pen_folder_name(email, project_folder_name, contact_label, topic_label):
    """Flat, fully self-descriptive folder name for a not-yet-formal
    project's email, filed directly under that year's holding pen
    (e.g. '26-000 MAILS') -- no company subfolder, no 03.-CORREO/
    ENTRANTE-SALIENTE nesting. Every email gets its own folder here, so
    the name has to carry everything needed to identify it from the
    listing alone: date, time, company/project name, sender, and a
    short subject hint.

    e.g. '26-07-29_11-05 ENCARNA BOSCH - E.BOSCH - PRESUPUESTO OBRA'
    """
    date_str = email["timestamp"].strftime("%y-%m-%d")
    time_str = email["timestamp"].strftime("%H-%M")
    return f"{date_str}_{time_str} {project_folder_name} - {contact_label} - {topic_label}"


def build_datetime_folder_name(email):
    """
    Builds the innermost folder name for one specific email's exact send/receive time,
    e.g. '2026-06-10_142305'. This is where the email text, attachments, and metadata.json
    actually live.
    """
    return email["timestamp"].strftime("%Y-%m-%d_%H%M%S")