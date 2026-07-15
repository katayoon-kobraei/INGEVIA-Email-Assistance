# dashboard/app.py
from flask import Flask, render_template
import os, json, subprocess

app = Flask(__name__)
OUTPUT_ROOT = r"C:\EmailAssistant\Output"

def load_emails():
    emails = []
    for folder in os.listdir(OUTPUT_ROOT):
        meta_path = os.path.join(OUTPUT_ROOT, folder, "metadata.json")
        if os.path.exists(meta_path):
            with open(meta_path, encoding="utf-8") as f:
                data = json.load(f)
                data["folder"] = folder
                emails.append(data)
    emails.sort(key=lambda e: e["received"], reverse=True)
    return emails

@app.route("/")
def dashboard():
    return render_template("dashboard.html", emails=load_emails())

@app.route("/open/<folder>")
def open_folder(folder):
    subprocess.Popen(["explorer", os.path.join(OUTPUT_ROOT, folder)])  # opens the real folder in File Explorer
    return "", 204

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000)