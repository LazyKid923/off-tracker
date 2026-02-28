# Off Tracker (Google Sheets + Apps Script)

Off Tracker is a Google Sheets + Apps Script tool for tracking time off (“offs”) per person: what was granted, what was used, and what remains. It builds a ready-to-use set of sheets, provides menu-driven workflows (modal dialogs) for adding/using/editing records, and generates a month view calendar.

## What It Can Do

- Track offs per person (switchable via a dropdown on `Dashboard`).
- Record **offs granted** (full day or half day) with a reason and metadata.
- Record **offs used** (AM / PM / full day), automatically deducting from available granted balances.
- Support **partial usage** and **splitting/combining** grants (e.g., two 0.5 grants can be used as 1 full day).
- Show totals and remaining balance on `Dashboard`, plus a color-coded monthly `Calendar`.
- Edit, delete, and undo actions with an audit trail in `Edit Logs`.
- Protect tracker-managed sheets from manual edits and recover if sheet structure is tampered with.

## Key Concepts

- **Offs (Granted)**: each grant is a record with an ID like `G-0001`, a granted date, duration (1 or 0.5), and a computed status (`Unused`, `Partial`, `Used`) based on remaining balance.
- **Offs (Used)**: each usage is a record with a Use ID like `U-0001`, intended date, session (AM/PM/Full), duration used, and a breakdown of which grant IDs were consumed (with amounts).

## Features (Detailed)

- **One-click setup**: `Build Layout` clears and recreates all tracker sheets with headers, formatting, and protections.
- **Reset & rebuild**: `Reset All & Rebuild` deletes existing tracker sheets and rebuilds from scratch (with confirmation).
- **Personnel management**:
  - Add personnel names and select the active person from the `Dashboard` dropdown.
  - Filter views in `Offs (Granted)`, `Offs (Used)`, and `Edit Logs` to the currently selected personnel.
  - Delete a personnel entry (optionally also delete all their related records).
- **Grant workflow (`Add Off Day`)**:
  - Duration: Full Day (1.0) or Half Day (0.5).
  - Reasons supported:
    - `Ops`: requires a weekend duty date (Saturday/Sunday) and auto-generates reason details.
    - `Others`: requires “provided by” and freeform details.
- **Use workflow (`Use Off Day`)**:
  - Session: AM (0.5), PM (0.5), or Full Day (1.0).
  - Pick one or more available `G-` IDs with remaining balance; the script allocates amounts until the needed duration is satisfied.
  - Automatically updates remaining balances and statuses in `Offs (Granted)`.
- **Edit/undo workflows**:
  - Edit granted records (`Edit Off Granted`).
  - Delete granted records in batch (`Delete Off Granted`) when they have not been used.
  - Edit used records or undo a used record (`Edit/Undo Off Used`); undo restores balances back to the affected `G-` IDs.
- **Dashboard + Calendar**:
  - Dashboard totals per selected person: total granted, total used, and remaining balance.
  - Calendar month selector (next 24 months) with a Monday-first grid and chips indicating granted (`+`) and used (`-`) on each day; includes a legend.
- **Audit trail (`Edit Logs`)**:
  - Logs edits/deletes/undo actions with before/after snapshots, a human-readable summary, timestamp, and editor identity (where available).
- **Protection & recovery**:
  - Managed sheets are protected against manual edits; edits are reverted automatically.
  - Structural changes (row/column/sheet insert/remove) trigger restore from internal backup sheets.

## Sheets Created

- `Dashboard`: select personnel and view totals/balance.
- `Offs (Granted)`: grants with remaining balance and status.
- `Offs (Used)`: usage records and which grants were consumed.
- `Calendar`: monthly view with a month/year dropdown.
- `Personnel`: list of personnel names.
- `Edit Logs`: audit trail of tracker actions.
- Hidden backups: `__BKP_GRANTED__`, `__BKP_USED__`, `__BKP_CALENDAR__`, `__BKP_LOGS__`.

## Getting Started

1. Create a Google Sheet (or upload `Off Tracker.xlsx` and open it with Google Sheets).
2. In the sheet, open **Extensions → Apps Script**.
3. Replace the script with the contents of `Code.js`.
4. Save, then reload the Google Sheet.
5. Use the **Off Day Tracker** menu → `Build Layout`.

Important:
- `Build Layout` clears and recreates the tracker-managed sheets. Run it once for setup, and avoid running it again unless you intend to rebuild the tracker.
- The tracker protects managed sheets from manual edits (the only intended manual input is selecting personnel on `Dashboard` and selecting the month on `Calendar`).

Optional:
- Add drawing buttons in `Dashboard` and assign scripts (e.g. `addOffDay`, `useOffDay`, `manageOffUsed`) if you prefer click buttons over the custom menu.

## Repo Contents

- `Code.js`: the full Apps Script implementation (menu, dialogs, sheet logic, protections, audit logging).
- `Off Tracker.xlsx`: a starter spreadsheet layout you can upload into Google Sheets.
- `appsscript.json`: Apps Script manifest.
- `.clasp.json`: clasp configuration (may contain IDs; treat as sensitive).
