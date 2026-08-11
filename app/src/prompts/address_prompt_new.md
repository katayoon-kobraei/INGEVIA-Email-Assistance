You are deciding which EXISTING PROJECT/SITE folder inside one already-matched company an email
belongs to. All output folder names must be copied exactly from the candidate list when matched.

## Context

The company itself has already been verified as an existing top-level server folder. It may have
several project/site subfolders, for example:

  26-003-01 CAMÍ DE FAITANAR 2 - PICAÑA
  26-003-02 AV RAMON MENENDEZ PIDAL 58 - ALBACETE
  26-003-03 SAN RAMON - CAMPELLO

These subfolders can represent a site, address, store, station, commission, or other project.
Use the full email subject/body and the supplied folder descriptions to decide which existing
project/site it belongs to.

## Decision

1. If the email clearly belongs to one existing project/site candidate, set
   matched_existing = true and return that exact folder name as address_folder_name.

2. If none of the existing project/site folders fits, set matched_existing = false. You may
   return a short bare description of the new/unknown project/site in address_folder_name, but
   the application will NOT create it automatically. For BOTH incoming and outgoing mail, the
   application will file the email folder in the year's YY-000 MAILS holding folder for manual
   review.

3. Do not force a match merely because the company is correct. A wrong project match is worse
   than correctly returning matched_existing = false.
