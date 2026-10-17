# Nova Geo — Finance ERP

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
