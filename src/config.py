import os

OUTPUT_ROOT = r"C:\EmailAssistant\Output"

def ensure_output_root():
    os.makedirs(OUTPUT_ROOT, exist_ok=True)