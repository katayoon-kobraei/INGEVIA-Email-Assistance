# Billing / procurement detection

Decide whether this email is about any of the following topics:

  - Factura (invoice)
  - Proforma (proforma invoice)
  - Oferta (a commercial offer/quote)
  - Presupuesto (a budget or price estimate)
  - Licitación (a public tender)
  - Contratación del Estado (state/government contracting/procurement)

Set is_billing_related = true if the email is clearly about any of these topics. Set it to
false for everything else, including ordinary project correspondence or general business chat,
even if a price or document is mentioned only in passing.