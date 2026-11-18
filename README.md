# Nova Geo — Finance ERP

## ⚠ Sign-in and security (do this first)
Earlier versions had no login, and the database accepted reads and writes from anyone holding the
site address. This version puts every screen behind a sign-in and locks the database to approved users.

1. Supabase > SQL editor: run **`supabase/001_security.sql`** (safe to re-run).
2. Open the app and **Create account**. The first account becomes **Admin**.
3. Other people create their own accounts; they wait as *Pending* until an admin approves them
   under **Admin › Users** and picks a role:
   - **Viewer**: sees everything, changes nothing (good for the CA / auditor)
   - **Accountant**: adds and edits entries
   - **Admin**: also manages users and reads the audit trail
4. **Admin › Audit Trail** lists every add, change and delete with who did it and the before/after
   values. Nobody (including admins) can edit or delete audit rows.
5. Supabase > Authentication > URL configuration: set *Site URL* to your Render URL so password-reset
   emails link back to the app.

The ops app's `finance-link` function uses the service key, so it keeps working after the lockdown.

## Database scripts (run in order, all safe to re-run)
| Script | What it adds |
|---|---|
| `supabase/001_security.sql` | Sign-in roles, locked tables, private bill storage, audit trail |
| `supabase/002_compliance.sql` | Customers, GST sales invoices & receipts, ITC register, TDS register & rates, compliance calendar |
| `supabase/003_capital_projects.sql` | Tables for Capital & Infra › Projects & Dues (this screen could not save before) |

## Sales & GST
- **Sales Invoices**: GST tax invoices, credit and debit notes. Numbers are consecutive per financial
  year (`NGR/26-27/0001`, max 16 characters as GST requires) and are given by the database, so two
  people can't get the same number. Issued invoices can't be deleted, only cancelled (the number stays
  used and is reported in GSTR-1 table 13). Once a month's GSTR-1 is marked filed, amounts on that
  month's invoices are locked; correct them with a credit/debit note.
- Place of supply decides IGST vs CGST+SGST automatically. GSTINs are checked (format + check digit).
- PDF in three copies (Original for recipient / Duplicate for transporter / Triplicate for supplier).
- Warnings for a missing e-way bill (over ₹50,000) and missing IRN (when e-invoicing is switched on).
- **Receipts** per invoice (can be matched to a bank credit; TDS deducted by the buyer counts as settled)
  and a **receivables ageing** report.
- **GST Returns**: GSTR-1 working with an Excel of the b2b / b2cl / b2cs / cdnr / hsn (B2B and B2C) /
  docs tables; GSTR-3B working with ITC set-off in the legal order; ITC register that pulls bills from
  Purchases and matches them against the GSTR-2B Excel from the portal (flags bills your supplier hasn't
  reported, which you can't claim yet).

## Compliance
- **Compliance Calendar**: GST (monthly or QRMP, Telangana dates), TDS deposits and quarterly returns,
  advance tax, LLP Form 11 and Form 8, ITR / tax audit, PF/ESI and Telangana professional tax (if
  switched on), plus your own items (licence renewals, insurance…). Mark each one filed with the ARN.
  Overdue items also show on the Dashboard. Use "Start tracking from today" the first time so returns
  already filed by your CA don't show as overdue.
- **TDS**: record deductions (rate suggested from the payee type and PAN; 20% without PAN), record the
  challan for several entries at once (late interest at 1.5% per month is worked out), quarterly
  deductee export for the Form 140 (old 26Q) return, and a **threshold check** that flags payees who
  got more than ₹1 lakh in the year from the bank statement with no TDS recorded.
  From 1 April 2026 TDS is under the Income-tax Act 2025 (Section 393). The seeded rates are a starting
  point; confirm them with your CA under TDS › Rates.
- **Tax Settings** (admin): GSTIN, TAN, bank details, invoice prefix and terms, and which compliances apply.
- **Suppliers** now hold GSTIN, PAN, payee type, usual TDS nature and MSME number.

## Other fixes
- "Today" used UTC, so between midnight and 5:30 am it gave yesterday's date. Now uses India time.
- Pop-up alerts replaced by small notices; failed database loads now show a message instead of failing silently.
- Each screen has its own address (`#gst`, `#tds`…), so the browser Back button and bookmarks work.
- Light / dark theme switch, print-friendly pages, better small-screen layout.

## Tests
`node tools/test-tax.mjs` checks the GST maths, GSTIN check digit, set-off order, TDS due dates and
interest, and the calendar dates.


Finance-only version of the Nova Geo ERP. All plant-operations modules (weighment, boulder in,
production, plant reconciliation, stock ledger, shift log, cost sheet / cost config, and the
vehicles / customers / grades masters) have been removed.

## What's in it
Dashboard · **Vendor Projects** · Purchases · Monthly Overheads · Expenses & Vendor Ledger ·
Capital & Infra · Bank Statement · Partners & Capital · Balance Sheet · Suppliers

## Vendor Projects (new)
Create a project for each job a vendor is doing (optional contract value). Then link that vendor's
payments to it:
- **Link Payment** — pick from the vendor's bank debits, expenses and partner-paid cash. Link the full
  amount or only part of it, so one payment can be split across projects.
- **Unassigned Payments** tab — every vendor payment not yet fully assigned, with one-click assign.
- **Manual Payment** — project-only payment that isn't recorded anywhere else.
- Per-project paid / due, progress bar, PDF statement, Excel export.

Linking only tags a payment; it never creates or changes bank / expense entries, so nothing is double-counted.

### One-time setup
Run `supabase/vendor_projects.sql` in the Supabase SQL editor, then reload the app.
Until then the app works normally and the Vendor Projects tab shows a setup notice.

## Balance sheet note
With operations gone there is no stock valuation or weighment revenue. "Revenue" is now the bank
credits from counterparties tagged **Customer / Buyer** on the Bank Statement screen.

## Run
```
npm install
npm run dev      # needs VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY
npm run build
```


## Purchase Orders (new)
The Purchase Orders screen creates supplier POs as PDF files, with vendor and billing details,
multiple line items, GST calculations, advance/balance summaries, PO history, duplicate and
cancel actions. It is separate from the existing Purchases register so it does not change current
expense or balance-sheet records.

### One-time setup
Open the Purchase Orders tab in the app, copy the displayed SQL, and run it in the Supabase SQL
Editor. Refresh the app afterward. The SQL creates `purchase_orders` and the `next_nova_geo_po_number()`
database function. Number format: `NGR/PO/YYYY-YY/001`, with a financial-year sequence.

### PO PDF appearance and advance
The PO PDF follows the supplied Nova Geo template: navy header and divider, pale-blue metadata/table headings and delivery band, bordered vendor/buyer columns, payment summary, terms, and authorised-signatory area. The "Less: Advance Paid (optional)" field may be left blank when no advance has been paid; the PDF calculates the balance payable from the order total minus the entered advance. Advance cannot be negative or exceed the PO total.
