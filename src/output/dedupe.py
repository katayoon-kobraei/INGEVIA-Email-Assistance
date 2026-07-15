# src/output/dedupe.py
import json
import os

def load_processed_ids(output_root):
    path = os.path.join(output_root, "_processed_ids.json")
    if not os.path.exists(path):
        return set()
    with open(path, "r", encoding="utf-8") as f:
        return set(json.load(f))

def mark_processed(email_id, output_root):
    ids = load_processed_ids(output_root)
    ids.add(email_id)
    path = os.path.join(output_root, "_processed_ids.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(list(ids), f)