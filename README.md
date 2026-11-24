# RMS Procurement Shipment Tracker

Push this folder to git and deploy as usual. **Nothing else to do** - on first start-up the app
merges the Excel tracker data into your live database by itself.

## What happens automatically on the first start after deploy (one time)
1. A full backup of the live database is saved first (`backups/pre_migration_<time>.db`).
2. New columns are added (Import license stage, Missing cargo document, Cargo dispatched, Clearing agent,
   Arrival notice ready, Actual arrival in Kigali, Remarks, Reason for delay, Responsible staff).
3. The tracker data in `seed/seed_data.json` is merged in:
   - **Existing shipments** are updated with the tracker's values - but a value somebody edited in the app
     since the last export is **kept** (the tracker's value is written to that shipment's Edit Trail instead).
   - **New shipments** from the tracker are added; **older sheets** are loaded into 2025-2026 and 2024-2025.
   - **Nothing is deleted.** Shipments you deleted stay deleted. Rows that exist only in the app stay untouched.
   - Every change is recorded in the Edit Trail with its old value.
4. If anything fails, the whole merge is rolled back (no partial data) and the app shows the reason.
5. It is recorded in the database and never runs twice.

A green message at the top of the app summarises what was merged.

## Also fixed in this version
- Bulk upload can no longer wipe data (blank cells never overwrite; unknown statuses and "TBA" values are handled).
- Smarter matching on upload (PO + BL, then PO + invoice; "TBA" -> real BL updates the row, no duplicate).
- **Delays had the wrong sign** (now arrival - deadline: positive = late, same as your Excel). Delays and warnings
  refresh every day.
- New look: banner, icon KPI cards, shipment-pipeline bar, rounded tabs and buttons.

## Backup button
In the tracker tab, **💾 Backup database → Prepare backup file → Download** saves a full copy of the live database
(all years, edit trail, emails) to your computer.

## Notes
- The database file itself is never part of git (`*.db` is ignored), so deploying cannot overwrite live data.
- `seed/Review_and_changes.xlsx` lists every change and the items that need a human check (17 tracker rows have
  cells in the wrong columns and are only partly imported; compound POs; rows with no PO are stored as `NO PO`).
- `original_data/` holds your original Excel files. `scripts/build_seed.py` rebuilds `seed/seed_data.json` from them.
- The seed contains your shipment data - keep the repository **private**.
