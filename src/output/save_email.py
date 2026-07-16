import os
import json
import shutil
import tempfile
from src.output.folder_namer import build_folder_name
from src.safety.attachment_scanner import check_attachment

QUARANTINE_ROOT = r"C:\EmailAssistant\Quarantine"

def save_email(email, department, output_root):
    folder_path = os.path.join(output_root, department, build_folder_name(email))
    os.makedirs(folder_path, exist_ok=True)

    with open(os.path.join(folder_path, "email.txt"), "w", encoding="utf-8") as f:
        f.write(f"From: {email['sender']}\nSubject: {email['subject']}\nReceived: {email['received']}\n\n")
        f.write(email["body"])

    attachment_results = []
    attachments = email["attachments"]

    with tempfile.TemporaryDirectory() as tmp_dir:
        for i in range(1, attachments.Count + 1):
            attachment = attachments.Item(i)
            tmp_path = os.path.join(tmp_dir, attachment.FileName)
            attachment.SaveAsFile(tmp_path)
            is_safe, reason = check_attachment(tmp_path)

            if is_safe:
                shutil.move(tmp_path, os.path.join(folder_path, attachment.FileName))
                attachment_results.append({"filename": attachment.FileName, "status": "saved", "reason": reason})
            else:
                os.makedirs(QUARANTINE_ROOT, exist_ok=True)
                shutil.move(tmp_path, os.path.join(QUARANTINE_ROOT, f"{email['id']}_{attachment.FileName}"))
                attachment_results.append({"filename": attachment.FileName, "status": "quarantined", "reason": reason})

    metadata = {
        "id": email["id"], "sender": email["sender"], "subject": email["subject"],
        "received": email["received"].isoformat(),
        "department": department,
        "attachments": attachment_results,
    }
    with open(os.path.join(folder_path, "metadata.json"), "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=2, ensure_ascii=False)

    return folder_path