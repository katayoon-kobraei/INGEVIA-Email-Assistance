import csv
import os
import shutil
import tempfile
import webbrowser
from datetime import datetime, timedelta

import src.output.status_page as status_page

TEST_ROOT = os.path.join(tempfile.gettempdir(), "email_assistant_status_test")


def write_fake_index(output_root):
    os.makedirs(output_root, exist_ok=True)
    index_path = os.path.join(output_root, "index.csv")
    now = datetime.now()

    rows = [
        [2026, "26-011 MEETPACK BENASAU", "ENTRANTE", "JOSE ANTONIO", "DOC ASISTENCIA",
         (now - timedelta(hours=1)).strftime("%Y-%m-%d %H:%M"),
         "Documentacion asistencia", "jose@meetpack.com", 2, r"C:\fake\path\26-011"],

        [2026, "26-012 UE ALBAL MANISES", "SALIENTE", "CONCEPCION SALMERON", "PRESUPUESTO",
         (now - timedelta(hours=3)).strftime("%Y-%m-%d %H:%M"),
         "RE: Presupuesto tecnico", "concepcion@example.com", 0, r"C:\fake\path\26-012"],

        [2026, "26-013 NACHO GADEA MONCADA", "ENTRANTE", "NACHO GADEA", "PLANO CAD",
         (now - timedelta(days=2)).strftime("%Y-%m-%d %H:%M"),
         "Plano CAD actualizado", "nacho@example.com", 1, r"C:\fake\path\26-013"],

        [2026, "UNSORTED", "ENTRANTE", "DESCONOCIDO", "SIN CLASIFICAR",
         (now - timedelta(hours=2)).strftime("%Y-%m-%d %H:%M"),
         "Consulta general sin identificar", "alguien@desconocido.com", 1, r"C:\fake\path\unsorted1"],
    ]

    with open(index_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow([
            "Year", "Project Folder", "Direction", "Contact", "Topic",
            "Date", "Subject", "Sender/Recipient", "Attachments", "Folder Path",
        ])
        writer.writerows(rows)


def main():
    if os.path.exists(TEST_ROOT):
        shutil.rmtree(TEST_ROOT)
    write_fake_index(TEST_ROOT)

    # Simulate 2 quarantined attachments without touching the real
    # C:\EmailAssistant\Quarantine folder.
    original_quarantine_root = status_page.QUARANTINE_ROOT
    fake_quarantine = os.path.join(TEST_ROOT, "FakeQuarantine")
    os.makedirs(fake_quarantine, exist_ok=True)
    open(os.path.join(fake_quarantine, "factura_sospechosa.exe"), "w").close()
    open(os.path.join(fake_quarantine, "adjunto_bloqueado.docm"), "w").close()
    status_page.QUARANTINE_ROOT = fake_quarantine

    output_path = status_page.generate_status_page(TEST_ROOT)
    status_page.QUARANTINE_ROOT = original_quarantine_root

    print(f"Test status page generated at: {output_path}")
    webbrowser.open(f"file:///{output_path.replace(os.sep, '/')}")


if __name__ == "__main__":
    main()