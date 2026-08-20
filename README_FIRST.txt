INGEVIA EMAIL AI ASSISTANT - PORTABLE INSTALLER v1.26.9
========================================================

IMPORTANT - INCOMING EMAILS ONLY
--------------------------------
This build processes only the configured boss mailbox Inbox (ENTRANTE).

It does NOT open, scan, classify, archive, flag, display, or otherwise process
Sent Items / SALIENTE. Existing legacy SALIENTE rows are hidden from the app.

v1.26.9 keeps all previous fixes:
- Correos/Emails table columns are now fully user-resizable: drag any header separator to expand/shrink a column; double-click a separator to auto-fit it.
- Outlook capture is separated from Gemini retry processing.
- AI queue runs every 2 minutes from the durable Unprocessed folder without
  reopening Outlook.
- 429/temporary Gemini errors remain Pending AI instead of DESCONOCIDO.
- Official noreply.dehu@correos.gob.es messages are recognized as DEHU when they are in the normal Inbox view; messages in Outlook Other/Otros are ignored by explicit requirement.
- One narrow 10-day DEHU recovery scan runs after upgrade to recover missing
  official DEHU notices without duplicating already captured EntryIDs.
- Known v1.25 path-length AI_REVIEW failures are recovered once from their
  durable Unprocessed copies.
- Final archive paths are shortened safely when a deep Windows path would be too
  long, without changing the company/project routing rules.
- AI_REVIEW is shown distinctly from Pending AI.
- Temporary boss-mailbox access failures get one fresh worker retry and the
  normal catch-up lookback prevents missed Inbox emails.

NORMAL FLOW
-----------
Every 30 minutes the processor checks only Inbox:

    Inbox email
        -> deterministic filters
        -> save durable staging copy under SHARED_DATA_PATH\Unprocessed
        -> flag/category the original Inbox item
        -> show Pending AI in the shared UI

Every 2 minutes the AI queue runs without Outlook:

    Unprocessed
        -> Gemini / routing
        -> existing company + project if matched
        -> otherwise 26-000 MAILS
        -> update shared index.csv / UI
        -> remove successful staging item

NO-ATTACHMENT / DEHU RULE
-------------------------
- Normal email with attachment: process normally.
- Normal email without attachment: ignore before Gemini.
- Official DEHU sender noreply.dehu@correos.gob.es: process even with no attachment, provided the message is not in Outlook Other/Otros.
- Forwarded DEHU notices can also be recognized from DEHU text in subject/body.

DEHU FOLDER NAMING
------------------
    YY-MM-DD Notificacion correo.gob.es Organismo Emisor

The organismo is extracted deterministically from the captured body, with no
extra Gemini request.

ROUTING - INCOMING ONLY
-----------------------
Company exists + project exists:
    P:\TRABAJOS 2026\<COMPANY>\<PROJECT>\03.-CORREO\ENTRANTE\<EMAIL FOLDER>

Company missing OR project missing:
    P:\TRABAJOS 2026\26-000 MAILS\<EMAIL FOLDER>

The application never invents company/project folders.

Normal incoming folder naming:
    YY-MM-DD Sender_Name Subject Company

STORAGE
-------
Shared application data:
    P:\Email Assistant Data

Real engineering archive:
    P:\TRABAJOS 2026

Do not use P:\Email Assistant Data\TRABAJOS 2026 as the engineering archive.

PROCESSING COMPUTER
-------------------
Run:
    INSTALL_AS_PROCESSING_COMPUTER.bat

Do not uninstall the previous version first. The installer preserves the
existing .env, Gemini API key, shared state, archive paths and scheduler setup.

VIEWER COMPUTER
---------------
Run:
    INSTALL_AS_VIEWER_COMPUTER.bat

Viewer PCs never access Outlook or Gemini. They only read shared application
state and open archived folders.

SCHEDULED TASKS
---------------
INGEVIA Email AI Assistant
    - Inbox capture every 30 minutes during the configured work period.

INGEVIA Email AI Assistant - AI Queue
    - Every 2 minutes.
    - Processes up to 2 staged Inbox emails.
    - Never opens Outlook.

DEFAULT AI QUEUE SETTINGS
-------------------------
    AI_QUEUE_BATCH_SIZE=2
    AI_QUEUE_RETRY_INITIAL_MINUTES=8
    AI_QUEUE_MAX_RETRY_ATTEMPTS=6
    GEMINI_MIN_REQUEST_INTERVAL_SECONDS=5
    GEMINI_429_COOLDOWN_MINUTES=6


v1.26.3 UI VISIBILITY FIX
--------------------------
- Official noreply.dehu@correos.gob.es records are never hidden by stale Outlook Other/Otros state.
- Existing DEHU EntryIDs accidentally stored in _ui_hidden_outlook_ids.json are automatically purged from the shared visibility state.
- This restores already-processed DEHU rows on both processor and viewer UIs without reprocessing or duplicating the archived email.

v1.26.3 INTERNAL EMAIL OUTLOOK MARKING
--------------------------------------
- Incoming @ingevia.com messages remain ignored by Gemini and are not archived or added to the application UI.
- They are now marked in Outlook with the same red follow-up flag / IA - PROCESADO category used to indicate that the assistant has inspected a message.
- The internal-email check runs before the dedupe skip, so a recent internal message recorded by v1.26.2 without the visual Outlook mark is repaired on the next Process New Emails / scheduled overlap scan.
- Internal messages remain excluded from the application's Flagged page even though Outlook itself shows the follow-up flag.
- Sent Items / SALIENTE remains disabled.


v1.26.4 CORREOS TABLE WIDTH FIX
--------------------------------
- The Correos/Emails table no longer lets long project names expand the Project column and push Contact, Sender, Subject and the remaining fields off-screen.
- Date, Direction, Project, Contact, Sender, Attachments, Priority and Status now start at compact, controlled widths.
- Subject and Summary automatically share the remaining window width.
- Long Project, Contact, Sender, Subject and Summary values remain available through tooltips and the columns can still be resized manually.
- This UI-only change keeps all v1.26.3 processing, DEHU, incoming-only and internal-email Outlook marking behavior unchanged.

v1.26.5 OUTLOOK STABILITY + UTF-8 HARDENING
----------------------------------------------
- Outlook Application and MAPI Namespace lifetimes are explicit and released in
  deterministic order at the end of every short-lived capture worker.
- Garbage collection now runs while COM is still initialized, with additional
  cleanup checkpoints between Outlook side-effect, recovery and visibility phases.
- Other/Otros legacy visibility reconciliation now reopens far fewer archived
  EntryIDs per 30-minute run. New Other/Otros messages are still filtered during
  normal Inbox ingestion, so routing/UI behavior is unchanged while MAPI pressure
  is reduced.
- The COM errors observed when Outlook entered the bad resource state
  ("Operación no disponible" / "Error en la ejecución de servidor") are treated
  as transient Outlook resource failures and get the clean-worker retry path.
- Worker stdout/stderr is forced to UTF-8. Emoji and other non-cp1252 characters
  in email subjects can no longer crash the Outlook capture log.
- Outlook is never closed by the assistant; the user's interactive Outlook session
  remains open.


v1.26.7 CRASH-SAFE INBOX CHECKPOINT
------------------------------------
- The last-successful Outlook scan timestamp is no longer advanced immediately
  after Inbox enumeration. It is committed only after the complete capture batch
  has finished without item-level capture failures.
- If a subject/logging/file/Outlook item unexpectedly fails halfway through a
  batch, the previous checkpoint is preserved. The next run therefore expands its
  lookback automatically and retries the missed interval.
- Existing EntryID deduplication means already captured emails are skipped, so the
  recovery does not create duplicate staged/archive records.
- This directly prevents the failure mode seen on 18/08 where a mid-batch exception
  could leave older emails unmarked and absent from the UI.


v1.26.7 OUTLOOK MARKING + UI PATH UPDATE
----------------------------------------
- External Inbox emails without actual attachments are still excluded from Gemini/archive/UI,
  but are now visibly marked in Outlook as inspected/processed.
- Older deduped messages with missing visual marks are repaired when they are seen again in the lookback window; no Gemini/archive duplication occurs.
- Outlook Focused Inbox Other/Otros is a hard exclusion again: normal Other/Otros messages are ignored completely (not marked, not staged, not sent to Gemini, and not shown in the UI).
- There are no sender exceptions: even DEHU is ignored when Outlook places it in Other/Otros.
- Correos/Emails UI removes Direction/Dirección and Status/Estado columns.
- Project/Proyecto display column is replaced by Saved in/Guardado en and shows the complete
  Folder Path stored in index.csv (including the final email folder path).
- Dashboard Recent activity uses the same simplified Date / Saved in / Subject / Priority view.


v1.26.9 OTHER/OTROS EXCLUSION RESTORED
----------------------------------------
- Outlook Focused Inbox Other/Otros is again a hard exclusion by explicit user requirement.
- Normal Other/Otros messages are ignored completely: no Outlook processed mark, no staging, no Gemini, no archive and no UI row.
- A real Outlook folder named Otros/Other is also excluded.
- There are no exceptions: DEHU is also ignored when it is in Other/Otros.
- The v1.26.7 no-attachment marking rule still applies to normal Inbox messages outside Other/Otros.
- Direction/Dirección and Status/Estado remain removed from the UI, and Saved in/Guardado en continues to show the complete saved folder path.
