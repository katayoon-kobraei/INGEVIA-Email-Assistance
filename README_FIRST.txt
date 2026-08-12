INGEVIA EMAIL AI ASSISTANT - PORTABLE INSTALLER v1.20
======================================================

This is the complete processor/viewer package based on the working v1.17 build.
Everything from v1.17 is preserved. v1.18 makes the boss-approved routing rule
explicit and identical for both incoming (ENTRANTE) and outgoing (SALIENTE) mail.


v1.20 - OUTLOOK FLAGGED VIEW RESTORED
--------------------------------------
The original intended workflow is restored:

    AI processes email -> red Outlook follow-up flag/category is applied
    -> "Marcados en Outlook" shows that processed flagged email.

The UI now resolves processed EntryIDs against the configured boss mailbox
StoreID, so additional/shared Outlook stores work correctly. Incoming and
outgoing processed emails are both included again. Automatic flag retries also
retain StoreID after transient Outlook sync conflicts.

IMPORTANT: TWO DIFFERENT STORAGE LOCATIONS
------------------------------------------

1) SHARED APP DATA FOLDER
   Used only for Email Assistant state and team-wide UI data such as index.csv,
   priorities.csv, reports, processed IDs and pending-response state.

   Example:
       P:\Email Assistant Data
   or its UNC equivalent.

2) REAL ENGINEERING TRABAJOS FOLDER
   Actual processed email copies and attachments are filed in the existing
   company/project hierarchy under:

       P:\TRABAJOS 2026

   The application must NOT create/use:
       P:\Email Assistant Data\TRABAJOS 2026
   as the engineering archive.

v1.18 BOSS-APPROVED ROUTING - ENTRANTE AND SALIENTE
----------------------------------------------------

The routing is COMPANY FIRST, then PROJECT, and both matches must correspond to
real folders already existing on the server. The same validation applies to
incoming and outgoing correspondence.

A) COMPANY FOUND + PROJECT FOUND - INCOMING

       P:\TRABAJOS 2026\<COMPANY>\<PROJECT>\03.-CORREO\ENTRANTE\<EMAIL FOLDER>

B) COMPANY FOUND + PROJECT FOUND - OUTGOING

       P:\TRABAJOS 2026\<COMPANY>\<PROJECT>\03.-CORREO\SALIENTE\<EMAIL FOLDER>

C) COMPANY FOUND + PROJECT NOT FOUND

   For either ENTRANTE or SALIENTE, the application does NOT invent/create a
   project. The email goes to:

       P:\TRABAJOS 2026\26-000 MAILS\<EMAIL FOLDER>

D) COMPANY NOT FOUND

   For either ENTRANTE or SALIENTE, the application does NOT invent/create a
   company. The email also goes to:

       P:\TRABAJOS 2026\26-000 MAILS\<EMAIL FOLDER>

EMAIL FOLDER NAMING
-------------------

Project/client email folders keep the existing human-readable convention:

    YY-MM-DD Contact_Name Subject Company

For incoming mail the contact is the sender. For outgoing mail the contact is
the recipient, so the folder remains useful to the engineering team. Example:

    26-08-07 Vanessa_Vierness Solicitud documentacion PLENERGY

The Outlook display name is preferred when available. Invalid Windows filename
characters are cleaned automatically and long subjects are trimmed.

SAFETY RULE
-----------

The AI may classify a company/project, but only an exact folder name that
physically exists on the server is treated as a match. This prevents the AI
from creating a fake company/project hierarchy.

PROCESSING COMPUTER
-------------------
Run:
    INSTALL_AS_PROCESSING_COMPUTER.bat

Use:
    Shared data folder:
        P:\Email Assistant Data
        (or the same UNC path used by the team)

    TRABAJOS 2026 folder:
        P:\TRABAJOS 2026
        (or its exact UNC equivalent)

    Boss Outlook mailbox:
        the mailbox that should be processed

Only ONE processing computer should be active. The processor reads both the
boss Inbox and Sent Items.

ENGINEER / VIEWER COMPUTERS
---------------------------
Run:
    INSTALL_AS_VIEWER_COMPUTER.bat

Use the SAME shared app-data folder and the SAME real TRABAJOS 2026 path used
by the processing computer. Viewer PCs do not process Outlook and do not run
the email scheduler.

UPDATING FROM v1.17
-------------------
You do not need to uninstall first. Run the appropriate v1.18 installer over
the existing installation. The real .env, Gemini key, shared paths, Outlook
mailbox, processed IDs, priorities and existing archived emails are preserved.

Existing archived email folders are not moved or renamed automatically. The
new explicit direction-safe routing applies to newly processed emails after the
v1.18 upgrade.
