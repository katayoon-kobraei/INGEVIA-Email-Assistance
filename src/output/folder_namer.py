def build_conversation_folder_name(email, contact_label, topic_label):
    date_str = email["timestamp"].strftime("%y-%m-%d")
    return f"{date_str} {contact_label}_{topic_label}"


def build_datetime_folder_name(email):
    """
    Builds the innermost folder name for one specific email's exact send/receive time,
    e.g. '2026-06-10_142305'. This is where the email text, attachments, and metadata.json
    actually live.
    """
    return email["timestamp"].strftime("%Y-%m-%d_%H%M%S")

