# Billing / procurement detection

Decide whether this email is about any of the following topics:

  - Factura (invoice)
  - Proforma (proforma invoice)
  - Oferta (a commercial offer/quote)
  - Presupuesto (a budget or price estimate)
  - Licitación (a public tender)
  - Contratación del Estado (state/government contracting/procurement)

Also set is_billing_related = true for ANY email sent by a bank or financial institution
(e.g. Cajamar, BBVA, CaixaBank, Santander, Sabadell, or similar), regardless of what the
email is actually about -- this includes routine notices, document requests, compliance/KYC
updates, or marketing-style communications from the bank, even if none of the six topics
above apply. A bank's own domain or letterhead/branding in the email is enough; correspondence
from banks always needs administración's attention, never ordinary project filing.

Set is_billing_related = true if the email is clearly about any of the six topics above, OR
if the sender is a bank/financial institution. Set it to false for everything else, including
ordinary project correspondence or general business chat, even if a price or document is
mentioned only in passing.