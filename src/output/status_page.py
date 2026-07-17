import csv
import os
import urllib.parse
from datetime import datetime, timedelta

try:
    from src.output.save_email import QUARANTINE_ROOT
except Exception:
    QUARANTINE_ROOT = None

STATUS_PAGE_NAME = "Estado.html"


def _to_file_uri(path):
    if not path:
        return "#"
    normalized = path.replace("\\", "/")
    return "file:///" + urllib.parse.quote(normalized, safe=":/")


def _load_index_rows(output_root):
    index_path = os.path.join(output_root, "index.csv")
    if not os.path.exists(index_path):
        return []
    with open(index_path, "r", newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def _parse_date(value):
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt)
        except (ValueError, TypeError):
            continue
    return None


def _count_quarantine_files():
    if not QUARANTINE_ROOT or not os.path.isdir(QUARANTINE_ROOT):
        return 0
    return len([
        f for f in os.listdir(QUARANTINE_ROOT)
        if os.path.isfile(os.path.join(QUARANTINE_ROOT, f))
    ])


def _esc(value):
    return (
        (str(value) if value is not None else "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def generate_status_page(output_root):
    rows = _load_index_rows(output_root)
    now = datetime.now()

    for row in rows:
        row["_parsed_date"] = _parse_date(row.get("Date", ""))

    rows_with_dates = [r for r in rows if r["_parsed_date"] is not None]
    rows_with_dates.sort(key=lambda r: r["_parsed_date"], reverse=True)

    total = len(rows)
    today_count = sum(1 for r in rows_with_dates if r["_parsed_date"].date() == now.date())
    week_count = sum(1 for r in rows_with_dates if r["_parsed_date"] >= now - timedelta(days=7))
    unsorted_rows = [r for r in rows if r.get("Project Folder") == "UNSORTED"]
    recent_rows = rows_with_dates[:15]
    quarantine_count = _count_quarantine_files()
    last_processed = rows_with_dates[0]["_parsed_date"] if rows_with_dates else None

    html = _render_html(
        total=total,
        today_count=today_count,
        week_count=week_count,
        unsorted_rows=unsorted_rows,
        recent_rows=recent_rows,
        quarantine_count=quarantine_count,
        last_processed=last_processed,
        now=now,
    )

    output_path = os.path.join(output_root, STATUS_PAGE_NAME)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    return output_path


def _render_html(total, today_count, week_count, unsorted_rows, recent_rows,
                  quarantine_count, last_processed, now):
    if last_processed and (now - last_processed) < timedelta(hours=24):
        status_color, status_text = "#2e7d32", "El sistema está funcionando correctamente."
    elif total == 0:
        status_color, status_text = "#757575", "Todavía no se ha procesado ningún correo."
    else:
        status_color, status_text = "#c77700", "No se ha procesado ningún correo nuevo en las últimas 24 horas."

    last_processed_str = last_processed.strftime("%d/%m/%Y %H:%M") if last_processed else "-"

    alert_html = ""
    if unsorted_rows:
        items = "".join(
            f"<li><b>{_esc(r.get('Subject'))}</b> - {_esc(r.get('Date'))} "
            f"(<a href=\"{_to_file_uri(r.get('Folder Path', ''))}\">abrir carpeta</a>)</li>"
            for r in unsorted_rows[:20]
        )
        alert_html = f"""
        <div class="alert">
            <h3>&#9888; {len(unsorted_rows)} correo(s) necesitan una revisión rápida</h3>
            <p>No se pudieron clasificar automáticamente en la carpeta de un cliente/proyecto.</p>
            <ul>{items}</ul>
        </div>
        """

    recent_rows_html = "".join(
        f"""<tr>
            <td>{_esc(r.get('Date'))}</td>
            <td>{'Entrante' if r.get('Direction') == 'ENTRANTE' else 'Saliente'}</td>
            <td>{_esc(r.get('Project Folder'))}</td>
            <td>{_esc(r.get('Contact'))}</td>
            <td>{_esc(r.get('Subject'))}</td>
            <td><a href="{_to_file_uri(r.get('Folder Path', ''))}">abrir carpeta</a></td>
        </tr>"""
        for r in recent_rows
    )

    quarantine_html = ""
    if quarantine_count:
        quarantine_html = f"""
        <div class="notice">
            {quarantine_count} archivo(s) adjunto(s) fueron bloqueados por seguridad y no se guardaron.
            Consulta con el equipo técnico si esperabas alguno de ellos.
        </div>
        """

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Estado del Asistente de Correo</title>
<style>
    body {{ font-family: Segoe UI, Arial, sans-serif; background: #f5f5f5; margin: 0; padding: 40px; color: #222; }}
    .container {{ max-width: 900px; margin: 0 auto; background: white; border-radius: 10px; padding: 32px; box-shadow: 0 1px 4px rgba(0,0,0,0.1); }}
    h1 {{ margin-top: 0; font-size: 24px; }}
    .status-banner {{ padding: 16px 20px; border-radius: 8px; color: white; background: {status_color}; font-size: 16px; margin-bottom: 24px; }}
    .metrics {{ display: flex; gap: 16px; margin-bottom: 24px; flex-wrap: wrap; }}
    .metric {{ flex: 1; min-width: 140px; background: #f0f4f8; border-radius: 8px; padding: 16px; text-align: center; }}
    .metric .num {{ font-size: 28px; font-weight: bold; }}
    .metric .label {{ font-size: 13px; color: #555; margin-top: 4px; }}
    .alert {{ background: #fff3e0; border-left: 4px solid #c77700; border-radius: 6px; padding: 16px 20px; margin-bottom: 24px; }}
    .notice {{ background: #fef9e7; border-left: 4px solid #d4a017; border-radius: 6px; padding: 12px 16px; margin-bottom: 24px; font-size: 14px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
    th, td {{ text-align: left; padding: 8px 10px; border-bottom: 1px solid #eee; }}
    th {{ background: #fafafa; }}
    a {{ color: #1565c0; text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
    .updated {{ color: #888; font-size: 12px; margin-top: 24px; }}
</style>
</head>
<body>
<div class="container">
    <h1>Estado del Asistente de Correo</h1>
    <div class="status-banner">{status_text}<br>Último correo procesado: {last_processed_str}</div>

    <div class="metrics">
        <div class="metric"><div class="num">{total}</div><div class="label">Correos procesados en total</div></div>
        <div class="metric"><div class="num">{today_count}</div><div class="label">Procesados hoy</div></div>
        <div class="metric"><div class="num">{week_count}</div><div class="label">Procesados esta semana</div></div>
        <div class="metric"><div class="num">{len(unsorted_rows)}</div><div class="label">Necesitan revisión</div></div>
    </div>

    {alert_html}
    {quarantine_html}

    <h3>Actividad reciente</h3>
    <table>
        <tr><th>Fecha</th><th>Dirección</th><th>Proyecto</th><th>Contacto</th><th>Asunto</th><th></th></tr>
        {recent_rows_html if recent_rows else '<tr><td colspan="6">Sin actividad todavía.</td></tr>'}
    </table>

    <div class="updated">Página generada automáticamente el {now.strftime('%d/%m/%Y %H:%M')}.</div>
</div>
</body>
</html>
"""