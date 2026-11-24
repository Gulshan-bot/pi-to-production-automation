# Changelog

## 1.2.0 (2026-11-24)

- Repositioned around the production sheet split: one confirmation now routes every PI item to the right product tab (PC to PC, PP to PP, and so on), so nobody distributes orders by hand and machine operators work straight from their own sheet.
- README rewritten: what one click creates, who saves time, where each item goes, setup including the planning sheet, and how to adapt the tabs and column layouts.
- Code comments updated to match. No change to the script logic.
- Project renamed from `order-to-delivery-pi-splitter` to `pi-to-production-automation`, because it now does more than split the PI: it also plans production. The menu in Google Sheets is still called "Order to Delivery".

## 1.1.0 (2026-11-24)

- New: optional planning sheet split (`PLAN.ENABLED`, off by default). One row per PI item goes to its product tab, rows the plant has not started on are replaced when a PI is confirmed again, revisions and possible duplicates are flagged, and items that match no tab are kept in `Unsorted PI items`.
- New menu items (shown when the split is on): check planning split, send this PI to planning sheet, remove a PI from the planning sheet.
- The confirmation dialog warns when the PI number is not in a form the planning split accepts.
- Fix: the backfill's clean-up of leftover temporary tabs now matches the temporary tab name (`PI_<timestamp>`) that the plant copy actually uses.

## 1.0.0 (2026-11-16)

- First public release: one-click PI split, customer IDs, revisions, PI register, plant mirror of the FMS tab, leak tests and backfill.
