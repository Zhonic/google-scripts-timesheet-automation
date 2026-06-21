# Timesheet Automation (Google Apps Script)

A self-managing timesheet system built on Google Sheets + Apps Script. One
**Controller** spreadsheet drives a fleet of per-employee timesheet files that
are generated from a template, shared automatically, locked on a weekly cycle,
and submitted (or auto-submitted) without any manual admin work.

It was built for a real service business and has been sanitised into a reusable
template: every deployment-specific value (Drive folder, template file, admin
emails, schedule) lives in configuration, and the per-file layout (which columns
hold dates, checkboxes and status) lives in a hidden `_Meta` sheet, so the logic
layer carries no hardcoded assumptions.

---

## What it does

- **Provisions files automatically.** When an employee row is marked `Active`,
  the script copies the template into a Drive folder, names it `<CODE>-Timesheet`,
  and applies admin access. It re-links an existing file by name instead of
  creating duplicates.
- **Manages access by employment status.** `Active` employees get edit access
  (from one day before their commencement date); `Left` employees keep edit
  access until the end of the current week, then drop to view-only; `Pick status`
  removes access. Permission changes are cached so the same change is never
  reapplied — this is what prevents Google "shared with you" notification spam.
- **Generates month tabs on demand** from a `_MonthTemplate`, populates the
  correct dates per week, writes the month title (e.g. `January 2026`), and
  pre-creates next month near month-end. Backdated commencement dates backfill
  all missing months.
- **Enforces weekly locking.** Only the current week is editable. Past weeks lock
  automatically; an admin can grant a time-boxed unlock window. Protection is
  rebuilt on every run, so tampering self-heals.
- **Handles employee actions** via two checkboxes per week — *Submit Timesheet*
  and *Request Unlock*. Clicks are validated, logged to a `Requests` sheet, and
  emailed to admins. Disallowed clicks are reverted instantly.
- **Auto-submits** any past week that was never submitted, flags blank weeks, and
  emails both the admins and the employee.
- **Sends weekly reminders** to anyone with an incomplete current week.

---

## Architecture

```
Controller spreadsheet  (you own this, script is bound to it)
├── Employees   ← one row per person; you edit this
├── Requests    ← audit log of submits/unlocks; admins approve here
└── Config      ← all deployment settings

Drive folder
├── ABCD-Timesheet   ← generated per employee from the template
│   ├── Jan-2026 / Feb-2026 / ...   ← created on demand
│   ├── _MonthTemplate  (hidden, locked)
│   ├── Time-Lists      (hidden, locked) — dropdown values for time pickers
│   └── _Meta           (hidden, locked) — per-file layout config
└── ...
```

The script is **bound to the Controller spreadsheet** (Extensions → Apps Script).
It opens each employee file by ID to enforce locks, ingest actions and send mail.

---

## Prerequisites

1. A Google account (Workspace recommended for domain sharing controls).
2. A Drive folder to hold generated timesheets.
3. A timesheet template file (see **Template setup** — read this carefully).
4. The **Advanced Drive Service** enabled in the Apps Script project (defaults to
   **v3**), used to suppress Drive share-notification emails. The code is
   version-tolerant — it tries Drive API v3, then v2, then a `DriveApp` fallback —
   so it still works without it, but the fallback can't suppress notifications.

---

## Template setup (read this — the format matters)

The timesheet template uses Google Sheets features that **do not survive an
`.xlsx` round-trip** — specifically checkboxes (the Submit / Request-Unlock
controls) and sheet protections. So there are two ways to get the template, and
the order matters:

**Recommended — copy the live Google Sheet (checkboxes intact):**

> **[Make a copy of the template](https://docs.google.com/spreadsheets/d/1iGzPMUqlWP9cSYoj-FV93UdBkNp2pKSMS1yjNNrIu4A/copy)**

This link forces Google to create your own copy with all checkboxes, dropdowns
and formatting preserved. Drop the copy into your timesheets folder and use its
file ID as `TEMPLATE_FILE_ID`. Nothing else to fix.

**Reference — the `.xlsx` in this repo:**

`Timesheet_Template.xlsx` is included so the layout is visible directly on
GitHub. **Do not use it as the live template without repairing it**, because the
checkbox cells import as the literal text `FALSE`. If you must start from the
xlsx (e.g. no access to the copy link): upload it, **File → Save as Google
Sheets**, then select columns **P** and **Q** on each week row (rows 3, 5, 7, 9,
11 — ten cells) and apply **Insert → Checkbox**. The script detects action rows
by looking for genuine boolean cells, so this step is mandatory. The time
dropdowns (columns B–O sourced from `Time-Lists`) are list validations and do
round-trip, so they survive the import.

The template layout (don't change without updating `_Meta`):

| Element            | Location                                      |
| ------------------ | --------------------------------------------- |
| Month title        | `A1` (written by the script)                  |
| Day headers        | Row 2: Mon/Tue/.../Sun across B, D, F, H, J, L, N |
| Week label rows    | A3 `Week 1`, A5 `Week 2`, … A11 `Week 5`      |
| Time entry rows    | The row directly below each week row (offset 1) |
| Start/end times    | Columns B–O on the time row (paired per day)  |
| Request Unlock box | Column **P** on the week row                  |
| Submit box         | Column **Q** on the week row                  |
| Week status        | Column **R** (written by the script)          |

The `Time-Lists` tab holds the dropdown options for the start/end time pickers
(column A = start times, column B = end times). Alongside the time values it
includes the placeholders `Pick start time` / `Pick end time` and a sentinel
value **`Didn't work`**. Selecting `Didn't work` is how an employee records a day
with no shift, as distinct from leaving the cell blank (forgot to fill in) — the
script treats these two cases differently on auto-submit (see **No-shift weeks**).

---

## Installation

The Controller is built from an **empty** spreadsheet — there's no Controller
template to copy. Setup creates and labels all three sheets (`Config`,
`Requests`, `Employees`) for you.

1. Create a **new blank** Google Sheet — this is your Controller. Open
   **Extensions → Apps Script**, delete the stub, and paste `Code.gs`. If using
   the manifest, also paste `appsscript.json` (enable "Show appsscript.json
   manifest file" under the editor's project settings first).
2. Enable the Advanced Drive Service: in the Apps Script editor, **Services (＋)**
   → **Drive API** → Add. This defaults to **v3** (the default since Dec 2023),
   which is what the script targets. If you use the included `appsscript.json`
   it's already declared as v3, but you may still need to toggle it on once.
3. Reload the Controller spreadsheet. A **Timesheets** menu appears.
4. **Timesheets → Setup (safe) + triggers.** This creates and labels the
   `Config`, `Requests` and `Employees` sheets (writing the Employees header
   row and a Status dropdown), then installs the time-based triggers. It is
   idempotent and will **not** overwrite sheets or data that already exist.
5. Fill in **Config** (see below). At minimum set `FOLDER_ID`,
   `TEMPLATE_FILE_ID` and `ADMIN_EMAILS`.
6. Add rows to **Employees**, set `Status` to `Active`, then run
   **Timesheets → Run Sync Now** once to verify.

On first run you'll be asked to authorise scopes (Sheets, Drive, Gmail send).

---

## Configuration reference (`Config` sheet)

| Key | Default | Purpose |
| --- | --- | --- |
| `FOLDER_ID` | *(paste)* | Drive folder ID where employee files are created. |
| `TEMPLATE_FILE_ID` | *(paste)* | File ID of the Google Sheets template. |
| `ADMIN_EMAILS` | `admin1@…,…` | Comma-separated admins; get edit access + notifications. |
| `UNLOCK_HOURS` | `24` | How long an approved unlock window stays open. |
| `TICK_MINUTES` | `5` | Sync cadence. Allowed: 1, 5, 10, 15, 30. |
| `REMINDER_WEEKDAY` | `SATURDAY` | Day the weekly reminder runs. |
| `REMINDER_HOUR` | `12` | Hour (0–23) the reminder runs. |
| `SEND_DRIVE_SHARE_NOTIFICATIONS` | `TRUE` | If `FALSE`, suppress Drive's share emails (needs Advanced Drive Service). |
| `SEND_CUSTOM_ACCESS_EMAILS` | `FALSE` | If `TRUE` *and* notifications are off, send one branded access email instead. |

### `Employees` sheet headers (row 1)

```
Employee_Code | First_Name | Middle_Name | Last_Name | Email | Commencement_Date | Timesheet_File_ID | Status | Notes
```

- `Employee_Code` is auto-generated if blank (3 letters of last name + a letter
  of the first name, made unique).
- `Commencement_Date` accepts AU `dd/mm/yyyy` or a real date cell.
- `Timesheet_File_ID` is filled in by the script — leave blank for new staff.
- `Status` is a dropdown: `Pick status`, `Active`, `Left`.
- `Notes` is written to by the script (status messages, warnings).

### `_Meta` sheet (inside each employee file)

Layout config, auto-created and back-filled with defaults. Override here if your
template differs.

| Key | Default | Meaning |
| --- | --- | --- |
| `TITLE_CELL` | `A1` | Where the month title is written. |
| `TITLE_FORMAT` | `MMMM yyyy` | Month title format. |
| `DATE_COLS` | `B,D,F,H,J,L,N` | Columns holding each weekday's date. |
| `TIME_ROW_OFFSET` | `1` | Rows from week label to the time row. |
| `TIME_EDIT_RANGE` | `B:O` | Columns employees fill in. |
| `ACTION_COL_REQUEST` | `P` | Request-Unlock checkbox column. |
| `ACTION_COL_SUBMIT` | `Q` | Submit checkbox column. |
| `ACTION_COL_STATUS` | `R` | Week status column. |
| `MONTH_SHEET_NAME_FORMAT` | `MMM-yyyy` | Month tab naming (e.g. `Jan-2026`). Detection derives from this. |
| `NO_SHIFT_VALUE` | `Didn't work` | Dropdown sentinel meaning "no shift this day". |

---

## How locking works

| Week state | Time cells | Submit (Q) | Request Unlock (P) |
| --- | --- | --- | --- |
| Future | locked | – | – |
| Current | editable | available | – |
| Past, locked | locked | – | available |
| Past, unlock approved (within window) | editable | available | – |
| Submitted | locked | – | available (to request a re-open) |

Admins approve unlocks in the `Requests` sheet by setting **Status → Approved**;
the next `processRequests` run opens a window of `UNLOCK_HOURS` and emails the
employee. Expired windows auto-relock.

---

## No-shift weeks

When the auto-submitter sweeps a past week that was never submitted, it
classifies the week into one of three states and reports each differently:

| State | What it means | Status flag | Email tone |
| --- | --- | --- | --- |
| Has entries | At least one real time was entered | *(none)* | neutral |
| No shifts | Every day is the `NO_SHIFT_VALUE` sentinel (`Didn't work`) | `(NO SHIFTS)` | informational |
| Blank | Nothing filled in (placeholders only) | `(NO TIME ENTRIES)` | red warning |

The distinction matters: a **blank** week usually means the employee forgot and
is worth chasing, while a **no-shifts** week is a deliberate "I had no shifts"
and shouldn't read as an error. The sentinel is configurable via `NO_SHIFT_VALUE`
in `_Meta`, so it stays in sync with whatever you list in the `Time-Lists`
dropdown.

---

## Testing safely

- Use a throwaway folder and one test employee row with **your own** email.
- Run **Run Sync Now** manually rather than waiting for the trigger.
- Check the **Executions** panel in the Apps Script editor for logs (the script
  logs per-employee errors without halting the whole run).
- To test locking without waiting a week, temporarily set a week's status to
  `Locked` and watch it become editable only in the current week.

---

## Adapting to a different layout or service

The logic layer reads everything from `_Meta`, so most changes are
configuration, not code:

- **Different columns?** Update `DATE_COLS`, `TIME_EDIT_RANGE`,
  `ACTION_COL_*` in `_Meta`.
- **Checkboxes on the time row instead of the week row?** Supported
  automatically — `getActionRowForBlock_` detects which row holds the booleans.
- **Different month tab names?** Just change `MONTH_SHEET_NAME_FORMAT`.
  Detection and name-parsing are derived from this value at runtime
  (`buildMonthSheetMatcher_`), so common formats like `MMM-yyyy`, `MMMM yyyy`
  and `yyyy-MM` work without code changes. Month names are matched in English;
  a non-English locale would need a localized month map in `MONTH3_INDEX_`.
- **Re-skin the tool name?** Edit the `BRAND` constant at the top of `Code.gs`
  (menu label, dialog title, protection namespace).

---

## Notes / limitations

- Sheet protections deter, but Google editors can still *unhide* system sheets;
  protection prevents edits, not visibility. Sensitive logic stays in the bound
  script, not in the sheets.
- Email sending is subject to the daily Gmail/Apps Script quota.
- The script assumes a Monday-start week.

---

*Author: Harnish Patel. Sanitised for portfolio use — all client identifiers,
folder/template IDs and admin emails have been replaced with placeholders.*
