# Email summary UI update

The **Correos / Filed emails** page now shows the same one-line summary used by `Informe de Emails.xlsx`.

Changes:

- Added a **Resumen / Summary** column for every filed email.
- Added a highlighted summary panel that updates when a row is selected.
- Added summary text to the page search.
- Added tooltips containing the full summary.
- Kept the existing summary in the email detail drawer.

Modified files:

- `desktop_app/main_window.py`
- `desktop_app/main_window_en.py`
- `desktop_app/styles.py`

No backend, Outlook, Gemini, routing, priority, attachment, or Excel generation logic was changed.
