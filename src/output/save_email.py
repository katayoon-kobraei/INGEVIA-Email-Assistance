from src.output.folder_namer import build_folder_name

import os
import json

def save_email(email, output_root):
    folder_path = os.path.join(output_root, build_folder_name(email))
    os.makedirs(folder_path, exist_ok=True)

    # Human-readable — for the boss/staff to open directly
    with open(os.path.join(folder_path, "email.txt"), "w", encoding="utf-8") as f:
        f.write(f"From: {email['sender']}\n")
        f.write(f"Subject: {email['subject']}\n")
        f.write(f"Received: {email['received']}\n\n")
        f.write(email["body"])

    # Structured — for the program itself: duplicate checks, future tooling, reporting
    metadata = {
        "id": email["id"],
        "sender": email["sender"],
        "subject": email["subject"],
        "received": email["received"].isoformat(),  # datetime isn't JSON-serializable directly, convert to string
        "attachment_count": email["attachments"].Count,
        # once classification exists, add here:
        # "priority": ..., "department": ..., "safety_verdict": ..., "reasoning": {...}
    }
    with open(os.path.join(folder_path, "metadata.json"), "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=2, ensure_ascii=False)

    attachments = email["attachments"]
    for i in range(1, attachments.Count + 1):
        attachment = attachments.Item(i)
        attachment.SaveAsFile(os.path.join(folder_path, attachment.FileName))

    return folder_path