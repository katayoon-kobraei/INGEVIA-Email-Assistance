

def build_folder_name(email):
    date_str = email["received"].strftime("%Y-%m-%d_%H%M")
    safe_subject = "".join(c for c in email["subject"] if c.isalnum() or c in " _-")[:50]
    return f"{date_str}_{safe_subject}"