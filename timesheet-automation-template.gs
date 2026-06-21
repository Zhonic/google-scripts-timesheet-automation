/**
 * Timesheet Automation (Controller Script)
 * ===========================================
 *
 * Author: Harnish Patel
 * Last Edited: 21/06/2026
 * Template Version: 3
 *
 * WHAT THIS SCRIPT DOES
 * ---------------------
 * A single "Controller" Google Sheet drives a fleet of per-employee timesheet
 * Sheets that are generated from a template. Per employee, the script:
 *   - Creates (or re-links) a personal timesheet file from a template.
 *   - Grants/revokes Drive access based on employment Status (Active/Left/Pick).
 *   - Generates month tabs on demand and pre-creates the next month near month-end.
 *   - Enforces weekly locking: only the current week (or an admin-approved unlock
 *     window) is editable; everything else is protected.
 *   - Handles employee-initiated "Submit" and "Request Unlock" actions (checkboxes)
 *     and logs them to a Requests sheet, emailing admins.
 *   - Auto-submits past weeks that were never submitted, and emails all parties.
 *   - Sends weekly reminders for incomplete timesheets.
 *
 * All deployment-specific values (folder, template, admins, cadence) live in a
 * "Config" sheet on the Controller. Per-file layout (which columns hold dates,
 * checkboxes, status, etc.) lives in a hidden "_Meta" sheet inside each employee
 * file, so the logic layer carries no hardcoded layout assumptions.
 *
 * COSMETIC BRANDING
 * -----------------
 * Menu labels, dialog titles and the internal protection namespace are centralised
 * in the BRAND constant below. Change them once to re-skin the tool.
 *
 * EMPLOYEES SHEET HEADERS (row 1)
 * ---------------------------------------
 * Employee_Code | First_Name | Middle_Name | Last_Name | Email | Commencement_Date | Timesheet_File_ID | Status | Notes
 *
 * Commencement_Date:
 * - AU format dd/mm/yyyy (e.g., 02/01/2026) OR an actual Date cell.
 */

/* ===========================
   Branding / cosmetic labels
   (the only place the tool name appears)
   =========================== */
const BRAND = {
  MENU_NAME: 'Timesheets',        // label shown in the spreadsheet's custom menu
  UI_TITLE: 'Timesheets',         // title used in alert/toast dialogs
  LOCK_NAMESPACE: 'TIMESHEET',    // namespace tagged onto sheet protections this script owns
};

/* ===========================
   Controller sheet names
   =========================== */
const SHEET_EMPLOYEES = 'Employees';
const SHEET_REQUESTS  = 'Requests';
const SHEET_CONFIG    = 'Config';

/* ===========================
   Script Properties keys
   =========================== */
const PROP_CONTROLLER_SSID = 'CONTROLLER_SPREADSHEET_ID';

/**
 * Permission cache key prefix:
 * Stores last applied mode per fileId + email.
 *
 * Example key:
 *   PERM_MODE|<fileId>|employee@domain.com => "editor" / "viewer" / "none"
 */
const PROP_PERM_MODE_PREFIX = 'PERM_MODE|';

/* ===========================
   Status values
   =========================== */
const STATUS_PICK   = 'Pick status';
const STATUS_ACTIVE = 'Active';
const STATUS_LEFT   = 'Left';

/* ===========================
   Defaults (used if Config not filled)
   =========================== */
const DEFAULTS = {
  FOLDER_ID: 'PASTE_FOLDER_ID_HERE',
  TEMPLATE_FILE_ID: 'PASTE_TEMPLATE_FILE_ID_HERE',
  ADMIN_EMAILS: 'admin1@example.com,admin2@example.com,admin3@example.com',
  UNLOCK_HOURS: '24',
  TICK_MINUTES: '5',
  REMINDER_WEEKDAY: 'SATURDAY',
  REMINDER_HOUR: '12',

  // Recommended TRUE to prevent Drive "shared with you" notification spam being suppressed.
  // (Set FALSE only if you want Drive's own share notifications sent.)
  SEND_DRIVE_SHARE_NOTIFICATIONS: 'TRUE',

  // If Drive notifications are disabled, send ONE custom email when access first granted.
  SEND_CUSTOM_ACCESS_EMAILS: 'FALSE',
};

/* ============================================================
   Controller Spreadsheet Resolver (robust for triggers)
   ============================================================ */
function getControllerSs_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  const id = PropertiesService.getScriptProperties().getProperty(PROP_CONTROLLER_SSID);
  if (!id) {
    throw new Error(
      `Controller spreadsheet not found.\n` +
      `Run setupControllerSafe() from inside the Controller spreadsheet:\n` +
      `Controller Sheet → Extensions → Apps Script → Run setupControllerSafe`
    );
  }
  return SpreadsheetApp.openById(id);
}

/* ===========================
   UI menu
   =========================== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(BRAND.MENU_NAME)
    .addItem('Setup (safe) + triggers', 'setupControllerSafe')
    .addSeparator()
    .addItem('Run Sync Now', 'tick')
    .addItem('Process Requests Now', 'processRequests')
    .addItem('Send Reminders Now', 'sendReminders')
    .addToUi();
}

/* ===========================
   Timezone helper
   =========================== */
function tz_() {
  return getControllerSs_().getSpreadsheetTimeZone();
}

/* ===========================
   Config handling
   =========================== */
function getConfig_() {
  const ss = getControllerSs_();
  const sh = ss.getSheetByName(SHEET_CONFIG);
  const cfg = { ...DEFAULTS };

  if (!sh) return cfg;

  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || '').trim();
    const val = String(data[i][1] || '').trim();
    if (key) cfg[key] = val;
  }
  return cfg;
}

function parseEmails_(s) {
  return String(s || '')
    .split(',')
    .map(x => x.trim())
    .filter(x => x && x.includes('@'));
}

function clampToAllowedMinutes_(n) {
  const allowed = [1, 5, 10, 15, 30];
  return allowed.includes(n) ? n : 5;
}

function cfgBool_(v, fallback) {
  if (v === undefined || v === null || v === '') return fallback;
  const s = String(v).trim().toUpperCase();
  if (['TRUE', 'YES', 'Y', '1'].includes(s)) return true;
  if (['FALSE', 'NO', 'N', '0'].includes(s)) return false;
  return fallback;
}

/* ===========================
   Header helpers (Employees sheet)
   =========================== */
function normHeader_(h) {
  return String(h || '')
    .trim()
    .toUpperCase()
    .replace(/[\s_]+/g, ''); // remove spaces/underscores
}

function buildHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = new Map();
  headers.forEach((h, i) => {
    const key = normHeader_(h);
    if (key) map.set(key, i + 1); // 1-based
  });
  return map;
}

function requireCol_(headerMap, headerName) {
  const col = headerMap.get(normHeader_(headerName));
  if (!col) throw new Error(`Missing required column header in Employees sheet: "${headerName}"`);
  return col;
}

function optionalCol_(headerMap, headerName) {
  return headerMap.get(normHeader_(headerName)) || null;
}

/* ===========================
   AU date parsing
   =========================== */
function parseDateAU_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;

  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;

  const dd = Number(m[1]);
  const mm = Number(m[2]);
  const yyyy = Number(m[3]);

  const d = new Date(yyyy, mm - 1, dd);
  if (Number.isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}

/* ===========================
   Column helpers
   =========================== */
function colToIndex_(col) {
  const c = String(col).toUpperCase().replace(/[^A-Z]/g, '');
  let n = 0;
  for (let i = 0; i < c.length; i++) n = n * 26 + (c.charCodeAt(i) - 64);
  return n;
}

function startOfDay_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function addDays_(d, days) {
  const x = new Date(d);
  x.setDate(x.getDate() + days);
  return x;
}

function monthStart_(d) {
  return new Date(d.getFullYear(), d.getMonth(), 1);
}

function compareMonth_(a, b) {
  const aa = monthStart_(a).getTime();
  const bb = monthStart_(b).getTime();
  return aa < bb ? -1 : aa > bb ? 1 : 0;
}

/**
 * End of current week (Sunday 23:59:59.999), with week starting Monday.
 */
function endOfCurrentWeek_(now) {
  const d = new Date(now);
  const dowMon0 = (d.getDay() + 6) % 7; // Mon=0..Sun=6
  const monday = new Date(d);
  monday.setDate(monday.getDate() - dowMon0);
  monday.setHours(0,0,0,0);

  const sunday = new Date(monday);
  sunday.setDate(sunday.getDate() + 6);
  sunday.setHours(23,59,59,999);
  return sunday;
}

/* ===========================
   SAFE setup (doesn't wipe your sheets)
   =========================== */
function setupControllerSafe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open your Controller spreadsheet, then run setupControllerSafe() again.');

  PropertiesService.getScriptProperties().setProperty(PROP_CONTROLLER_SSID, ss.getId());

  ensureConfigKeys_(ss);
  ensureRequestsHeaders_(ss);
  ensureEmployeesHeaders_(ss);
  ensureEmployeesValidations_(ss);

  deleteAllProjectTriggers_();
  createTriggers_();

  uiAlertSafe_(
    'Setup complete.\n\nNext:\n1) Fill Config values\n2) Add employee rows\n3) Set Status to Active when ready\n4) Run "Run Sync Now" once to test'
  );
}

function ensureConfigKeys_(ss) {
  let sh = ss.getSheetByName(SHEET_CONFIG);
  if (!sh) sh = ss.insertSheet(SHEET_CONFIG);

  if (sh.getLastRow() === 0) sh.getRange(1,1,1,2).setValues([['Key','Value']]);

  const data = sh.getDataRange().getValues();
  const existing = new Set(data.slice(1).map(r => String(r[0] || '').trim()).filter(Boolean));

  const rows = [
    ['FOLDER_ID', DEFAULTS.FOLDER_ID],
    ['TEMPLATE_FILE_ID', DEFAULTS.TEMPLATE_FILE_ID],
    ['ADMIN_EMAILS', DEFAULTS.ADMIN_EMAILS],
    ['UNLOCK_HOURS', DEFAULTS.UNLOCK_HOURS],
    ['TICK_MINUTES', DEFAULTS.TICK_MINUTES],
    ['REMINDER_WEEKDAY', DEFAULTS.REMINDER_WEEKDAY],
    ['REMINDER_HOUR', DEFAULTS.REMINDER_HOUR],

    ['SEND_DRIVE_SHARE_NOTIFICATIONS', DEFAULTS.SEND_DRIVE_SHARE_NOTIFICATIONS],
    ['SEND_CUSTOM_ACCESS_EMAILS', DEFAULTS.SEND_CUSTOM_ACCESS_EMAILS],
  ];

  rows.forEach(r => {
    if (!existing.has(r[0])) sh.appendRow(r);
  });
}

function ensureRequestsHeaders_(ss) {
  let sh = ss.getSheetByName(SHEET_REQUESTS);
  if (!sh) sh = ss.insertSheet(SHEET_REQUESTS);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 12).setValues([[
      'Request ID', 'Timestamp', 'Employee Code', 'Employee Email',
      'Month Sheet', 'Week Index', 'Week Start', 'Week End',
      'Type', 'Status', 'Unlock Until', 'Notes'
    ]]);
  }
}

/**
 * Canonical Employees header row. The script is header-driven (see buildHeaderMap_),
 * so these are the names loadEmployees_/processRequests look up.
 */
const EMPLOYEE_HEADERS = [
  'Employee_Code', 'First_Name', 'Middle_Name', 'Last_Name', 'Email',
  'Commencement_Date', 'Timesheet_File_ID', 'Status', 'Notes'
];

/**
 * Ensures the Employees sheet exists AND has a header row.
 *
 * SAFE: only writes headers when row 1 is empty (a fresh/blank sheet). If the
 * sheet already has content in row 1, it's left untouched — a genuinely
 * mislabelled header is surfaced later by requireCol_ rather than silently
 * overwritten.
 */
function ensureEmployeesHeaders_(ss) {
  let sh = ss.getSheetByName(SHEET_EMPLOYEES);
  if (!sh) sh = ss.insertSheet(SHEET_EMPLOYEES);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  let row1Blank = true;
  if (lastRow >= 1 && lastCol >= 1) {
    const vals = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    row1Blank = vals.every(v => String(v || '').trim() === '');
  }

  if (lastRow === 0 || row1Blank) {
    sh.getRange(1, 1, 1, EMPLOYEE_HEADERS.length).setValues([EMPLOYEE_HEADERS]);
    try { sh.setFrozenRows(1); } catch (e) {}
    try { sh.getRange(1, 1, 1, EMPLOYEE_HEADERS.length).setFontWeight('bold'); } catch (e) {}
  }

  return sh;
}

function ensureEmployeesValidations_(ss) {
  let sh = ss.getSheetByName(SHEET_EMPLOYEES);
  if (!sh) sh = ss.insertSheet(SHEET_EMPLOYEES);

  const headerMap = buildHeaderMap_(sh);
  const statusCol = requireCol_(headerMap, 'Status');

  const dv = SpreadsheetApp.newDataValidation()
    .requireValueInList([STATUS_PICK, STATUS_ACTIVE, STATUS_LEFT], true)
    .setAllowInvalid(true)
    .build();

  sh.getRange(2, statusCol, sh.getMaxRows() - 1, 1).setDataValidation(dv);
}

function deleteAllProjectTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}

function createTriggers_() {
  const cfg = getConfig_();
  const tickMins = clampToAllowedMinutes_(Number(cfg.TICK_MINUTES || 5));

  ScriptApp.newTrigger('tick').timeBased().everyMinutes(tickMins).create();
  ScriptApp.newTrigger('processRequests').timeBased().everyMinutes(tickMins).create();

  const reminderDay = String(cfg.REMINDER_WEEKDAY || 'SATURDAY').toUpperCase();
  const reminderHour = Number(cfg.REMINDER_HOUR || 12);

  const weekdayMap = {
    'SUNDAY': ScriptApp.WeekDay.SUNDAY,
    'MONDAY': ScriptApp.WeekDay.MONDAY,
    'TUESDAY': ScriptApp.WeekDay.TUESDAY,
    'WEDNESDAY': ScriptApp.WeekDay.WEDNESDAY,
    'THURSDAY': ScriptApp.WeekDay.THURSDAY,
    'FRIDAY': ScriptApp.WeekDay.FRIDAY,
    'SATURDAY': ScriptApp.WeekDay.SATURDAY,
  };

  ScriptApp.newTrigger('sendReminders')
    .timeBased()
    .onWeekDay(weekdayMap[reminderDay] || ScriptApp.WeekDay.SATURDAY)
    .atHour(reminderHour)
    .create();
}

/* ===========================
   Employee row parsing (HEADER-DRIVEN + BULLETPROOF)
   =========================== */
function loadEmployees_() {
  const ss = getControllerSs_();
  const sh = ss.getSheetByName(SHEET_EMPLOYEES);
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const headerMap = buildHeaderMap_(sh);

  const colCode  = requireCol_(headerMap, 'Employee_Code');
  const colFirst = requireCol_(headerMap, 'First_Name');
  const colMid   = optionalCol_(headerMap, 'Middle_Name');
  const colLast  = requireCol_(headerMap, 'Last_Name');
  const colEmail = requireCol_(headerMap, 'Email');
  const colComm  = requireCol_(headerMap, 'Commencement_Date');
  const colFile  = requireCol_(headerMap, 'Timesheet_File_ID');
  const colStat  = requireCol_(headerMap, 'Status');
  const colNotes = requireCol_(headerMap, 'Notes');

  // Reserve all existing codes to prevent collisions
  const existingCodes = sh.getRange(2, colCode, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(v => String(v || '').trim())
    .filter(Boolean);

  const reservedCodes = new Set(existingCodes);

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const out = [];

  for (let i = 0; i < data.length; i++) {
    const rowNum = i + 2;

    const codeRaw  = data[i][colCode  - 1];
    const firstRaw = data[i][colFirst - 1];
    const midRaw   = colMid ? data[i][colMid - 1] : '';
    const lastRaw  = data[i][colLast  - 1];
    const emailRaw = data[i][colEmail - 1];
    const commRaw  = data[i][colComm  - 1];
    const fileRaw  = data[i][colFile  - 1];
    const statRaw  = data[i][colStat  - 1];
    const notesRaw = data[i][colNotes - 1];

    const hasAny = codeRaw || firstRaw || midRaw || lastRaw || emailRaw || commRaw || fileRaw || statRaw || notesRaw;
    if (!hasAny) continue;

    let code   = String(codeRaw || '').trim();
    let first  = String(firstRaw || '').trim();
    let middle = String(midRaw || '').trim();
    let last   = String(lastRaw || '').trim();
    let email  = String(emailRaw || '').trim();
    let status = String(statRaw || '').trim();
    let notes  = String(notesRaw || '').trim();
    let fileId = String(fileRaw || '').trim();

    if (!status) {
      status = STATUS_PICK;
      sh.getRange(rowNum, colStat).setValue(STATUS_PICK);
    }

    // require first+last
    if (!first || !last) {
      sh.getRange(rowNum, colNotes).setValue('Cannot run automation: First_Name and Last_Name are required.');
      continue;
    }

    // Don't treat this row's existing code as a collision while evaluating
    if (code) reservedCodes.delete(code);

    const expectedCode = makeEmployeeCodeUnique_(last, first, reservedCodes);

    if (!code) {
      code = expectedCode;
      sh.getRange(rowNum, colCode).setValue(code);
      reservedCodes.add(code);
      sh.getRange(rowNum, colNotes).setValue('Employee_Code generated automatically.');
    } else if (code !== expectedCode) {
      if (!fileId) {
        sh.getRange(rowNum, colCode).setValue(expectedCode);
        sh.getRange(rowNum, colNotes).setValue(`Employee_Code corrected from "${code}" to "${expectedCode}".`);
        code = expectedCode;
      } else {
        sh.getRange(rowNum, colNotes).setValue(
          `WARNING: Employee_Code "${code}" != expected "${expectedCode}". File already exists; not changing code.`
        );
      }
    }

    // Re-reserve the final code for subsequent rows
    if (code) reservedCodes.add(code);

    const commencementDate = parseDateAU_(commRaw);

    notes = String(sh.getRange(rowNum, colNotes).getValue() || '').trim();

    out.push({
      row: rowNum,
      code,
      first,
      middle,
      last,
      email,
      commencementDate,
      fileId,
      status,
      notes,
    });
  }

  return out;
}

/**
 * Creates a 4-char Employee Code:
 * - first 3 letters of LAST name
 * - then tries 1st/2nd/3rd/... letter of FIRST name until unique
 * - if first name runs out, falls back to digits 0-9
 */
function makeEmployeeCodeUnique_(lastName, firstName, reservedSet) {
  const ln = String(lastName || '').trim().toUpperCase().replace(/[^A-Z]/g, '');
  const fn = String(firstName || '').trim().toUpperCase().replace(/[^A-Z]/g, '');

  const base = (ln.length >= 3) ? ln.slice(0, 3) : ln.padEnd(3, 'X');

  // Try letters of first name: 1st, 2nd, 3rd...
  for (let i = 0; i < fn.length; i++) {
    const candidate = base + fn[i];
    if (!reservedSet.has(candidate)) return candidate;
  }

  // Fallback: digits
  for (let d = 0; d <= 9; d++) {
    const candidate = base + String(d);
    if (!reservedSet.has(candidate)) return candidate;
  }

  // Last resort
  return base + 'X';
}

/**
 * Ensures exactly ONE non-admin collaborator exists on the file: the current employee email.
 * - Removes any other non-admin editors/viewers (covers email changes cleanly)
 * - Then applies desiredMode to the current employee email.
 */
function reconcileEmployeeFileAccess_(fileId, employeeEmail, desiredMode) {
  if (!fileId) return;

  const cfg = getConfig_();
  const admins = parseEmails_(cfg.ADMIN_EMAILS).map(e => e.toLowerCase());
  const adminSet = new Set(admins);

  const target = String(employeeEmail || '').trim();
  const targetLower = target.toLowerCase();

  const file = DriveApp.getFileById(fileId);

  // Remove any non-admin collaborators that are NOT the current employee email
  const editors = file.getEditors().map(u => u.getEmail()).filter(Boolean);
  const viewers = file.getViewers().map(u => u.getEmail()).filter(Boolean);

  editors.forEach(e => {
    const el = e.toLowerCase();
    if (!adminSet.has(el) && el !== targetLower) {
      try { file.removeEditor(e); } catch (err) {}
    }
  });

  viewers.forEach(e => {
    const el = e.toLowerCase();
    if (!adminSet.has(el) && el !== targetLower) {
      try { file.removeViewer(e); } catch (err) {}
    }
  });

  // Apply desired access to current employee email
  if (!target) return;

  const isEditorNow = file.getEditors().some(u => u.getEmail().toLowerCase() === targetLower);
  const isViewerNow = file.getViewers().some(u => u.getEmail().toLowerCase() === targetLower);

  if (desiredMode === 'editor') {
    if (isViewerNow) { try { file.removeViewer(target); } catch (e) {} }
    if (!isEditorNow) { file.addEditor(target); }
  } else if (desiredMode === 'viewer') {
    if (isEditorNow) { try { file.removeEditor(target); } catch (e) {} }
    if (!isViewerNow) { file.addViewer(target); }
  } else if (desiredMode === 'none') {
    if (isEditorNow) { try { file.removeEditor(target); } catch (e) {} }
    if (isViewerNow) { try { file.removeViewer(target); } catch (e) {} }
  }
}

function uiAlertSafe_(msg, title) {
  title = title || BRAND.UI_TITLE;
  try {
    SpreadsheetApp.getUi().alert(title, msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // No UI context (running from Apps Script editor / trigger) — fallback
    try {
      const ss = getControllerSs_();
      ss.toast(msg, title, 10);
    } catch (e2) {}
    Logger.log(`${title}: ${msg}`);
  }
}

/* ===========================
   Employee system sheets hidden
   =========================== */
function enforceEmployeeHiddenSheets_(empSs) {
  ['Time-Lists', '_MonthTemplate', '_Meta'].forEach(name => {
    const sh = empSs.getSheetByName(name);
    if (sh) { try { sh.hideSheet(); } catch (e) {} }
  });
}

/* ===========================
   Permission spam-proofing
   =========================== */
function permCacheKey_(fileId, email) {
  return `${PROP_PERM_MODE_PREFIX}${fileId}|${String(email || '').trim().toLowerCase()}`;
}

function getCachedPermMode_(fileId, email) {
  return PropertiesService.getScriptProperties().getProperty(permCacheKey_(fileId, email)) || '';
}

function setCachedPermMode_(fileId, email, mode) {
  PropertiesService.getScriptProperties().setProperty(permCacheKey_(fileId, email), mode);
}

/**
 * Adds permission using the Advanced Drive service, with notification suppression.
 *
 * Version-tolerant: works whether the project has Drive API v3 OR v2 enabled.
 * Since Dec 2023, enabling "Drive API" in Advanced Services defaults to v3, where:
 *   - the method is Permissions.create (not insert),
 *   - the email field is `emailAddress` (not `value`),
 *   - the flag is `sendNotificationEmail` (singular, not `sendNotificationEmails`).
 * We try v3 first, then v2, then fall back to DriveApp (which cannot suppress
 * the "shared with you" notification).
 *
 * @returns {string} 'v3' | 'v2' | 'driveapp' — which path was used (handy for logs/tests).
 */
function driveAddPermission_(fileId, email, role /* 'writer'|'reader' */, sendNotify) {
  // --- Try Drive API v3: Drive.Permissions.create ---
  try {
    if (typeof Drive !== 'undefined' && Drive.Permissions && Drive.Permissions.create) {
      Drive.Permissions.create(
        { type: 'user', role: role, emailAddress: email },
        fileId,
        { sendNotificationEmail: !!sendNotify }
      );
      return 'v3';
    }
  } catch (e) {
    // fall through to v2 / DriveApp
  }

  // --- Try Drive API v2: Drive.Permissions.insert ---
  try {
    if (typeof Drive !== 'undefined' && Drive.Permissions && Drive.Permissions.insert) {
      Drive.Permissions.insert(
        { type: 'user', role: role, value: email },
        fileId,
        { sendNotificationEmails: !!sendNotify }
      );
      return 'v2';
    }
  } catch (e) {
    // fall through to DriveApp
  }

  // --- Fallback: built-in DriveApp (cannot suppress the share notification) ---
  const f = DriveApp.getFileById(fileId);
  if (role === 'writer') f.addEditor(email);
  else f.addViewer(email);
  return 'driveapp';
}

/**
 * Sets employee access in a spam-proof way:
 * - Uses cache to ensure we only apply changes ONCE per mode change.
 * - Avoids touching permissions if the employee is the file OWNER.
 * - If SEND_DRIVE_SHARE_NOTIFICATIONS is FALSE, uses Drive API no-notify (if enabled).
 *
 * Returns: true if a change was applied, false if no-op.
 */
function setEmployeeAccess_(fileId, email, mode /* 'editor'|'viewer'|'none' */, cfg) {
  if (!fileId || !email) return false;

  const target = String(email).trim();
  if (!target) return false;

  const desired = mode;
  const cached = getCachedPermMode_(fileId, target);

  // If we've already applied this mode, do nothing. (THIS is the core spam stopper.)
  if (cached === desired) return false;

  const file = DriveApp.getFileById(fileId);

  // If permissions already match desired, just sync cache and exit (prevents repeat work/spam)
  const targetLower = target.toLowerCase();
  const isEditorNow = file.getEditors().some(u => (u.getEmail() || '').toLowerCase() === targetLower);
  const isViewerNow = file.getViewers().some(u => (u.getEmail() || '').toLowerCase() === targetLower);

  if (desired === 'editor' && isEditorNow) {
    setCachedPermMode_(fileId, target, desired);
    return false;
  }
  if (desired === 'viewer' && isViewerNow && !isEditorNow) {
    setCachedPermMode_(fileId, target, desired);
    return false;
  }
  if (desired === 'none' && !isEditorNow && !isViewerNow) {
    setCachedPermMode_(fileId, target, desired);
    return false;
  }

  // Owner-safe: if employee email is the owner, do NOT try to add them as editor.
  // (Owners are not listed as editors; re-adding can cause repeated "shared with you".)
  try {
    const ownerEmail = (file.getOwner && file.getOwner()) ? String(file.getOwner().getEmail() || '').toLowerCase() : '';
    if (ownerEmail && ownerEmail === target.toLowerCase()) {
      setCachedPermMode_(fileId, target, desired);
      return false;
    }
  } catch (e) {}

  const sendNotify = cfgBool_(cfg.SEND_DRIVE_SHARE_NOTIFICATIONS, false);

  // Apply desired mode
  try {
    if (desired === 'none') {
      try { file.removeEditor(target); } catch (e) {}
      try { file.removeViewer(target); } catch (e) {}
    }

    if (desired === 'viewer') {
      try { file.removeEditor(target); } catch (e) {}
      // add viewer (prefer Drive API)
      driveAddPermission_(fileId, target, 'reader', sendNotify);
    }

    if (desired === 'editor') {
      try { file.removeViewer(target); } catch (e) {}
      // add editor (prefer Drive API)
      driveAddPermission_(fileId, target, 'writer', sendNotify);
    }

    // Update cache ONLY after successful attempt
    setCachedPermMode_(fileId, target, desired);
    return true;
  } catch (err) {
    console.error('setEmployeeAccess_ failed:', err);
    return false;
  }
}

function maybeSendCustomAccessEmail_(emp, fileId, mode, cfg) {
  const sendCustom = cfgBool_(cfg.SEND_CUSTOM_ACCESS_EMAILS, true);
  const sendNotify = cfgBool_(cfg.SEND_DRIVE_SHARE_NOTIFICATIONS, false);

  // If Drive notifications are ON, don't duplicate with custom emails.
  if (!sendCustom || sendNotify) return;

  if (mode !== 'editor' && mode !== 'viewer') return;
  if (!emp.email) return;

  const link = `https://docs.google.com/spreadsheets/d/${fileId}/edit`;
  const roleWord = (mode === 'editor') ? 'edit' : 'view';

  MailApp.sendEmail({
    to: emp.email,
    subject: `[Timesheet] Access granted`,
    htmlBody:
      `<p>Hi ${emp.first || ''},</p>` +
      `<p>Your timesheet is now available. You can ${roleWord} it here:</p>` +
      `<p><a href="${link}">${link}</a></p>` +
      `<p>Thanks!</p>`
  });
}

/* ===========================
   Template copy + anti-duplicate linking
   =========================== */
function findExistingEmployeeFileId_(empCode, cfg) {
  try {
    const folder = DriveApp.getFolderById(cfg.FOLDER_ID);
    const name = `${empCode}-Timesheet`;
    const files = folder.getFilesByName(name);
    if (files.hasNext()) return files.next().getId();
  } catch (e) {}
  return null;
}

function ensureEmployeeFile_(emp) {
  if (emp.fileId) return emp.fileId;

  const cfg = getConfig_();

  const ctrlSs = getControllerSs_();
  const sh = ctrlSs.getSheetByName(SHEET_EMPLOYEES);
  const headerMap = buildHeaderMap_(sh);
  const colFile = requireCol_(headerMap, 'Timesheet_File_ID');
  const colNotes = requireCol_(headerMap, 'Notes');

  // SAFETY: try to re-link existing file by name to avoid duplicates
  const existingId = findExistingEmployeeFileId_(emp.code, cfg);
  if (existingId) {
    sh.getRange(emp.row, colFile).setValue(existingId);
    sh.getRange(emp.row, colNotes).setValue('Re-linked existing timesheet file by name (avoided duplicate creation).');
    SpreadsheetApp.flush();
    return existingId;
  }

  // Otherwise create a copy from template
  const folder = DriveApp.getFolderById(cfg.FOLDER_ID);
  const templateFile = DriveApp.getFileById(cfg.TEMPLATE_FILE_ID);

  const newName = `${emp.code}-Timesheet`;
  const copy = templateFile.makeCopy(newName, folder);
  const newId = copy.getId();

  const empSs = SpreadsheetApp.openById(newId);

  // Ensure _Meta exists BEFORE we try to lock it
  ensureEmployeeMeta_(empSs);

  // Now lock/hide system sheets (incl _Meta)
  lockAndHideSystemSheets_(empSs);

  sh.getRange(emp.row, colFile).setValue(newId);
  SpreadsheetApp.flush();

  // Admin editors
  const admins = parseEmails_(cfg.ADMIN_EMAILS);
  const file = DriveApp.getFileById(newId);
  admins.forEach(a => { try { file.addEditor(a); } catch (e) {} });

  // Prevent employee editors from re-sharing / changing permissions
  try { file.setShareableByEditors(false); } catch (e) {}

  // Hide system sheets in the employee file
  try {
    enforceEmployeeHiddenSheets_(empSs);
  } catch (e) {}

  sh.getRange(emp.row, colNotes).setValue('Timesheet file created and admin access applied.');
  return newId;
}

/* ===========================
   Month creation gating rules
   =========================== */
function computeShareStart_(emp, now) {
  if (!emp.commencementDate) return now;
  const cd = startOfDay_(emp.commencementDate);
  return startOfDay_(addDays_(cd, -1));
}

function shouldCreateMonth_(emp, monthDate) {
  if (!emp.commencementDate) return true;
  return compareMonth_(monthDate, emp.commencementDate) >= 0;
}

function isThirdLastDayOrLater_(now) {
  const y = now.getFullYear();
  const m = now.getMonth();
  const last = new Date(y, m + 1, 0);
  const thirdLast = new Date(y, m + 1, last.getDate() - 2);
  thirdLast.setHours(0, 0, 0, 0);

  const today = new Date(now);
  today.setHours(0, 0, 0, 0);

  return today.getTime() >= thirdLast.getTime();
}

/* ===========================
   Main periodic runner (SAFE)
   =========================== */
function tick() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;

  try {
    const cfg = getConfig_();
    const now = new Date();
    const employees = loadEmployees_();

    employees.forEach(emp => {
      try {
        const status = emp.status || STATUS_PICK;

        // Only create/link file when Active/Left
        if (status === STATUS_ACTIVE || status === STATUS_LEFT) {
          emp.fileId = ensureEmployeeFile_(emp);
        }

        if (!emp.fileId) return;

        // Open employee file once for enforcement actions
        const empSs = SpreadsheetApp.openById(emp.fileId);

        // Keep _Meta present + locked forever
        ensureEmployeeMeta_(empSs);
        lockAndHideSystemSheets_(empSs);
        enforceEmployeeHiddenSheets_(empSs);


        if (status === STATUS_PICK) {
          reconcileEmployeeFileAccess_(emp.fileId, emp.email, 'none'); // cleanup old emails
          setEmployeeAccess_(emp.fileId, emp.email, 'none', cfg);      // updates cache too
          return;
        }

        if (status === STATUS_ACTIVE) {
          const shareStart = computeShareStart_(emp, now);

          if (now >= shareStart) {
            const changed = setEmployeeAccess_(emp.fileId, emp.email, 'editor', cfg);
            if (changed) maybeSendCustomAccessEmail_(emp, emp.fileId, 'editor', cfg);
          } else {
            reconcileEmployeeFileAccess_(emp.fileId, emp.email, 'none');
            setEmployeeAccess_(emp.fileId, emp.email, 'none', cfg);
          }

          if (now >= shareStart) {
            // ---- Backfill: create any missing past months from commencement to now ----
            // This covers employees added with a commencement date in a prior month.
            if (emp.commencementDate) {
              let cursor = monthStart_(emp.commencementDate);
              const currentMonth = monthStart_(now);
              while (cursor.getTime() < currentMonth.getTime()) {
                if (shouldCreateMonth_(emp, cursor)) ensureMonthSheet_(empSs, cursor);
                cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
              }
            }

            // Current month
            if (shouldCreateMonth_(emp, now)) ensureMonthSheet_(empSs, now);

            // Pre-create next month if near end of current month
            if (isThirdLastDayOrLater_(now)) {
              const next = new Date(now.getFullYear(), now.getMonth() + 1, 1);
              if (shouldCreateMonth_(emp, next)) ensureMonthSheet_(empSs, next);
            }

            repairMonthTitles_(empSs);
            enforceLocksForEmployee_(empSs, now);
            ingestEmployeeActions_(empSs, emp, now);
            autoSubmitPastWeeks_(empSs, emp, now);
          }
          return;
        }

        if (status === STATUS_LEFT) {
          const weekEnd = endOfCurrentWeek_(now);

          if (now <= weekEnd) {
            const changed = setEmployeeAccess_(emp.fileId, emp.email, 'editor', cfg);
            if (changed) maybeSendCustomAccessEmail_(emp, emp.fileId, 'editor', cfg);
          } else {
            reconcileEmployeeFileAccess_(emp.fileId, emp.email, 'viewer');
            const changed = setEmployeeAccess_(emp.fileId, emp.email, 'viewer', cfg);
            if (changed) maybeSendCustomAccessEmail_(emp, emp.fileId, 'viewer', cfg);
          }

          repairMonthTitles_(empSs);
          enforceLocksForEmployee_(empSs, now);
          ingestEmployeeActions_(empSs, emp, now);
          autoSubmitPastWeeks_(empSs, emp, now);
        }
      } catch (empErr) {
        console.error(`tick(): employee ${emp.code || '(no code)'} failed:`, empErr);
      }
    });
  } catch (err) {
    console.error('tick() fatal error:', err);
  } finally {
    lock.releaseLock();
  }
}

/* ===========================
   Month sheet creation (template-driven)
   =========================== */
function ensureMonthSheet_(empSs, dateInMonth) {
  ensureEmployeeMeta_(empSs);

  const meta = readMeta_(empSs);
  const fmt = meta.MONTH_SHEET_NAME_FORMAT || 'MMM-yyyy';
  const monthName = Utilities.formatDate(monthStart_(dateInMonth), tz_(), fmt);

  if (empSs.getSheetByName(monthName)) return;

  const tpl = empSs.getSheetByName('_MonthTemplate');
  if (!tpl) throw new Error('Employee file missing "_MonthTemplate".');

  tpl.copyTo(empSs).setName(monthName).showSheet();

  const sheet = empSs.getSheetByName(monthName);
  populateMonthDates_(sheet, monthStart_(dateInMonth), meta);
  updateMonthTitle_(sheet, monthStart_(dateInMonth), meta);
  ensureSheetProtection_(sheet);

  enforceEmployeeHiddenSheets_(empSs);
}

/**
 * Updates the month title cell (e.g. A1) with the correct month/year.
 * Uses TITLE_FORMAT from _Meta (default: "MMMM yyyy" → "January 2026").
 */
function updateMonthTitle_(sheet, monthStartDate, meta) {
  const titleCell = meta.TITLE_CELL || 'A1';
  const titleFmt  = meta.TITLE_FORMAT || 'MMMM yyyy';

  const titleText = Utilities.formatDate(monthStartDate, tz_(), titleFmt);
  sheet.getRange(titleCell).setValue(titleText);
}

/**
 * Repairs titles on ALL existing month sheets (fixes sheets created before the title fix).
 * Lightweight: only writes if the current title doesn't match the expected value.
 */
function repairMonthTitles_(empSs) {
  const meta = readMeta_(empSs);
  const titleCell = meta.TITLE_CELL || 'A1';
  const titleFmt  = meta.TITLE_FORMAT || 'MMMM yyyy';

  const monthSheets = empSs.getSheets().filter(s => isMonthSheetName_(s.getName(), meta));

  monthSheets.forEach(sh => {
    // Parse the month back from the sheet name (format-aware)
    const parsed = parseMonthSheetName_(sh.getName(), meta);
    if (!parsed) return;

    const expected = Utilities.formatDate(parsed, tz_(), titleFmt);
    const current  = String(sh.getRange(titleCell).getValue() || '');

    if (current !== expected) {
      sh.getRange(titleCell).setValue(expected);
    }
  });
}

/* ===========================
   Month-sheet name detection (DERIVED from MONTH_SHEET_NAME_FORMAT)
   ---------------------------------------------------------------
   Instead of hardcoding /^[A-Za-z]{3}-\d{4}$/ (which silently breaks if the
   _Meta format changes), we build a matcher from the configured format so the
   same single source of truth drives both naming and detection.
   Supported tokens: yyyy, yy, MMMM, MMM, MM, M, dd, d. Everything else is a
   literal. Month names are matched in English (matching Utilities.formatDate
   defaults); other locales would need a localized month map.
   =========================== */
const MONTH3_INDEX_ = {
  jan:0, feb:1, mar:2, apr:3, may:4, jun:5,
  jul:6, aug:7, sep:8, oct:9, nov:10, dec:11
};

function escapeRegExp_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Builds an anchored matcher from a date format like "MMM-yyyy" or "MMMM yyyy".
 * Returns { re, monthKind } where monthKind is 'monthText' | 'monthNum' | null.
 * Year and month are captured via named groups (?<year>) and (?<month>).
 */
function buildMonthSheetMatcher_(fmt) {
  const tokens = [
    ['yyyy', '(?<year>\\d{4})',         'year'],
    ['yy',   '(?<year>\\d{2})',         'year'],
    ['MMMM', '(?<month>[A-Za-z]+)',     'monthText'],
    ['MMM',  '(?<month>[A-Za-z]{3})',   'monthText'],
    ['MM',   '(?<month>0[1-9]|1[0-2])',   'monthNum'],
    ['M',    '(?<month>1[0-2]|[1-9])',     'monthNum'],
    ['dd',   '\\d{2}',                  'day'],
    ['d',    '\\d{1,2}',                'day'],
  ];

  let pat = '';
  let monthKind = null;
  let i = 0;

  while (i < fmt.length) {
    let matched = false;
    for (const [tok, frag, kind] of tokens) {
      if (fmt.startsWith(tok, i)) {
        pat += frag;
        if (kind === 'monthText' || kind === 'monthNum') monthKind = kind;
        i += tok.length;
        matched = true;
        break;
      }
    }
    if (!matched) {
      pat += escapeRegExp_(fmt[i]);
      i++;
    }
  }

  return { re: new RegExp('^' + pat + '$'), monthKind };
}

/**
 * True if a sheet name matches the configured month-sheet format.
 * System sheets are always excluded.
 */
function isMonthSheetName_(name, meta) {
  const SYSTEM = ['_Meta', '_MonthTemplate', 'Time-Lists'];
  if (SYSTEM.includes(name)) return false;

  const fmt = (meta && meta.MONTH_SHEET_NAME_FORMAT) || 'MMM-yyyy';
  const { re } = buildMonthSheetMatcher_(fmt);
  return re.test(String(name || ''));
}

/**
 * Parses a month sheet name back into a Date (1st of that month), format-aware.
 * Handles "Jan-2026", "January 2026", "2026-01", etc. depending on the format.
 */
function parseMonthSheetName_(name, meta) {
  const fmt = (meta && meta.MONTH_SHEET_NAME_FORMAT) || 'MMM-yyyy';
  const { re, monthKind } = buildMonthSheetMatcher_(fmt);

  const m = re.exec(String(name || ''));
  if (!m || !m.groups) return null;

  // Month
  let monthIndex;
  if (monthKind === 'monthText') {
    const key = String(m.groups.month || '').slice(0, 3).toLowerCase();
    monthIndex = MONTH3_INDEX_[key];
  } else if (monthKind === 'monthNum') {
    monthIndex = Number(m.groups.month) - 1;
  }
  if (monthIndex === undefined || monthIndex < 0 || monthIndex > 11) return null;

  // Year (expand 2-digit years to 20xx)
  let year = Number(m.groups.year);
  if (m.groups.year && m.groups.year.length === 2) year = 2000 + year;
  if (!year) return null;

  return new Date(year, monthIndex, 1);
}

/* ===========================
   Employee _Meta (layout config)
   =========================== */
function ensureEmployeeMeta_(empSs) {
  let meta = empSs.getSheetByName('_Meta');
  if (!meta) meta = empSs.insertSheet('_Meta');

  if (meta.getLastRow() === 0) meta.getRange(1,1,1,2).setValues([['Key','Value']]);

  const defaults = [
    ['TITLE_CELL', 'A1'],
    ['TITLE_FORMAT', 'MMMM yyyy'],
    ['DATE_COLS', 'B,D,F,H,J,L,N'],
    ['TIME_ROW_OFFSET', '1'],
    ['TIME_EDIT_RANGE', 'B:O'],
    ['ACTION_COL_REQUEST', 'P'],
    ['ACTION_COL_SUBMIT', 'Q'],
    ['ACTION_COL_STATUS', 'R'],
    ['MONTH_SHEET_NAME_FORMAT', 'MMM-yyyy'],
    ['NO_SHIFT_VALUE', "Didn't work"],
  ];

  const data = meta.getDataRange().getValues();
  const existing = new Set(data.slice(1).map(r => String(r[0]||'').trim()).filter(Boolean));

  defaults.forEach(([k,v]) => {
    if (!existing.has(k)) meta.appendRow([k,v]);
  });

  meta.hideSheet();
}

function readMeta_(empSs) {
  const meta = empSs.getSheetByName('_Meta');
  const out = {};
  if (!meta) return out;

  const data = meta.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const k = String(data[i][0] || '').trim();
    const v = String(data[i][1] || '').trim();
    if (k) out[k] = v;
  }
  return out;
}

/* ===========================
   Date population (Week rows)
   =========================== */
function populateMonthDates_(sheet, monthStart, meta) {
  const year = monthStart.getFullYear();
  const month = monthStart.getMonth();
  const lastDay = new Date(year, month + 1, 0);

  const dateColsLetters = (meta.DATE_COLS || 'B,D,F,H,J,L,N').split(',').map(s => s.trim());
  const dateCols = dateColsLetters.map(colToIndex_);

  const lastRow = sheet.getLastRow();
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues().map(r => String(r[0] || ''));

  const weekRows = [];
  for (let r = 1; r <= lastRow; r++) {
    if (/^Week\s+\d+/i.test(colA[r - 1])) weekRows.push(r);
  }

  weekRows.forEach(wr => dateCols.forEach(c => sheet.getRange(wr, c).clearContent()));

  const firstDowMon0 = (monthStart.getDay() + 6) % 7;
  for (let d = 1; d <= lastDay.getDate(); d++) {
    const cur = new Date(year, month, d);
    const offset = (d - 1) + firstDowMon0;
    const weekIndex = Math.floor(offset / 7);
    const dow = offset % 7;

    if (weekIndex >= weekRows.length) break;
    sheet.getRange(weekRows[weekIndex], dateCols[dow]).setValue(cur);
  }
}

/* ===========================
   Protection & weekly locking
   =========================== */
/**
 * Ensures the month sheet is protected and only admins can edit protected cells.
 * Everyone else (spreadsheet editors) can only edit ranges in setUnprotectedRanges().
 *
 * IMPORTANT:
 * - Protection does NOT support setEditors(). Use addEditors/removeEditors instead.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {GoogleAppsScript.Spreadsheet.Protection}
 */
function ensureSheetProtection_(sheet) {
  const cfg = getConfig_();
  const admins = parseEmails_(cfg.ADMIN_EMAILS);

  // Always keep the effective user as an allowed editor of the protection (prevents lockouts).
  const me = String(Session.getEffectiveUser().getEmail() || '').trim();
  const allowed = [...new Set([...admins, me].filter(Boolean))];

  // IMPORTANT: Template leftovers can include RANGE protections that block checkbox clicks
  // even when the sheet protection has unprotected ranges.
  // Remove them so sheet-level protection rules are the single source of truth.
  try {
    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(rp => rp.remove());
  } catch (e) {}


  const DESC = `${BRAND.LOCK_NAMESPACE}_LOCK`;

  // Remove any sheet protections not ours, and de-duplicate ours
  let ours = [];
  try {
    const all = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    all.forEach(sp => {
      if (sp.getDescription() === DESC) ours.push(sp);
      else sp.remove();
    });
  } catch (e) {}

  // Keep exactly ONE managed protection (multiple protections can "stack" and block edits)
  let p = ours[0];
  if (ours.length > 1) {
    ours.slice(1).forEach(sp => { try { sp.remove(); } catch (e) {} });
  }

  if (!p) {
    p = sheet.protect();
    p.setDescription(DESC);
  }

  // Enforce real protection (not warning-only)
  p.setWarningOnly(false);

  // Turn off domain-wide edit (if applicable)
  try { p.setDomainEdit(false); } catch (e) {}

  // Remove anyone not in allowed list from protection editors
  try {
    const currentEditors = p.getEditors().map(u => u.getEmail()).filter(Boolean);
    const toRemove = currentEditors.filter(e => !allowed.includes(e));
    if (toRemove.length) p.removeEditors(toRemove);
  } catch (e) {}

  // Add missing allowed editors
  try {
    const currentEditors = p.getEditors().map(u => u.getEmail()).filter(Boolean);
    const toAdd = allowed.filter(e => !currentEditors.includes(e));
    if (toAdd.length) p.addEditors(toAdd);
  } catch (e) {}

  return p;
}

function getWeekBlocks_(sheet, meta) {
  const lastRow = sheet.getLastRow();
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues().map(r => String(r[0] || ''));
  const offset = Number(meta.TIME_ROW_OFFSET || 1);

  const blocks = [];
  for (let r = 1; r <= lastRow; r++) {
    const m = colA[r - 1].match(/^Week\s+(\d+)/i);
    if (!m) continue;
    blocks.push({ weekIndex: Number(m[1]), weekRow: r, timeRow: r + offset });
  }
  return blocks;
}

/**
 * Detect which row actually contains the action checkboxes (P/Q):
 * - Some templates place them on the Week row
 * - Others place them on the Time row
 */
function getActionRowForBlock_(sheet, b, colReq, colSub) {
  const wReq = sheet.getRange(b.weekRow, colReq).getValue();
  const wSub = sheet.getRange(b.weekRow, colSub).getValue();
  const tReq = sheet.getRange(b.timeRow, colReq).getValue();
  const tSub = sheet.getRange(b.timeRow, colSub).getValue();

  const weekHas = (typeof wReq === 'boolean') || (typeof wSub === 'boolean');
  const timeHas = (typeof tReq === 'boolean') || (typeof tSub === 'boolean');

  if (weekHas && !timeHas) return b.weekRow;
  if (timeHas && !weekHas) return b.timeRow;
  if (weekHas) return b.weekRow;     // if both, prefer Week row
  return b.timeRow;                  // fallback
}

function readStatusForBlock_(sheet, b, colSta, preferredRow) {
  const otherRow = (preferredRow === b.weekRow) ? b.timeRow : b.weekRow;

  const pref = String(sheet.getRange(preferredRow, colSta).getValue() || '');
  const other = String(sheet.getRange(otherRow, colSta).getValue() || '');

  const rp = statusRank_(pref);
  const ro = statusRank_(other);

  // If neither is "known", stick to preferred row
  if (rp === 0 && ro === 0) return { text: pref, row: preferredRow };

  // Prefer whichever row has the "stronger" state (e.g., Unlocked until beats Locked)
  return (rp >= ro)
    ? { text: pref, row: preferredRow }
    : { text: other, row: otherRow };
}

function statusRank_(s) {
  s = String(s || '');
  if (s.startsWith('Unlocked until')) return 50;
  if (s.startsWith('Submitted')) return 40;
  if (s.startsWith('Auto-submitted')) return 40;
  if (s.includes('Requested')) return 30;
  if (s.startsWith('Unlocked (current week)')) return 20;
  if (s.startsWith('Locked')) return 10;
  return 0;
}

// Always write status to ONE row and clear the other row if it contains another "known" status.
// This prevents the lock logic reading a different row later.
function setBlockStatus_(sheet, b, colSta, rowToUse, text) {
  const otherRow = (rowToUse === b.weekRow) ? b.timeRow : b.weekRow;

  sheet.getRange(rowToUse, colSta).setValue(text);

  try {
    const otherVal = String(sheet.getRange(otherRow, colSta).getValue() || '');
    if (statusRank_(otherVal) > 0) sheet.getRange(otherRow, colSta).clearContent();
  } catch (e) {}
}

function computeWeekStartEnd_(sheet, weekRow, meta) {
  const dateColsLetters = (meta.DATE_COLS || 'B,D,F,H,J,L,N').split(',').map(s => s.trim());
  const dateCols = dateColsLetters.map(colToIndex_);

  const dates = dateCols
    .map(c => sheet.getRange(weekRow, c).getValue())
    .filter(v => v instanceof Date);

  if (dates.length === 0) return { start: null, end: null };

  const earliest = new Date(Math.min.apply(null, dates.map(d => d.getTime())));
  const dowMon0 = (earliest.getDay() + 6) % 7;

  const start = new Date(earliest);
  start.setDate(start.getDate() - dowMon0);
  start.setHours(0,0,0,0);

  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  end.setHours(23,59,59,999);

  return { start, end };
}

function enforceLocksForEmployee_(empSs, now) {
  const meta = readMeta_(empSs);
  const monthSheets = empSs.getSheets().filter(s => isMonthSheetName_(s.getName(), meta));
  monthSheets.forEach(sh => enforceLocksOnMonthSheet_(sh, now, meta));
}

/**
 * Parses the date embedded in "Unlocked until dd-MMM HH:mm" status text.
 * Returns a Date if parseable, or null.
 */
function parseUnlockUntilDate_(statusText, referenceDate) {
  const match = String(statusText || '').match(
    /Unlocked until (\d{1,2})-([A-Za-z]{3})\s+(\d{1,2}):(\d{2})/
  );
  if (!match) return null;

  const day   = Number(match[1]);
  const monStr = match[2];
  const hour  = Number(match[3]);
  const min   = Number(match[4]);

  const months = {
    'Jan':0,'Feb':1,'Mar':2,'Apr':3,'May':4,'Jun':5,
    'Jul':6,'Aug':7,'Sep':8,'Oct':9,'Nov':10,'Dec':11
  };
  const monthNum = months[monStr];
  if (monthNum === undefined) return null;

  // Use the reference year; the unlock is always near "now"
  const year = referenceDate.getFullYear();
  return new Date(year, monthNum, day, hour, min, 0, 0);
}

function enforceLocksOnMonthSheet_(sheet, now, meta) {
  const p = ensureSheetProtection_(sheet);
  const blocks = getWeekBlocks_(sheet, meta);

  const colReq = colToIndex_(meta.ACTION_COL_REQUEST || 'P');
  const colSub = colToIndex_(meta.ACTION_COL_SUBMIT || 'Q');
  const colSta = colToIndex_(meta.ACTION_COL_STATUS || 'R');

  const timeParts = String(meta.TIME_EDIT_RANGE || 'B:O').split(':');
  const timeStart = colToIndex_(timeParts[0]);
  const timeEnd   = colToIndex_(timeParts[1]);
  const timeWidth = (timeEnd - timeStart) + 1;

  const unprotected = [];

  blocks.forEach(b => {
    const { start, end } = computeWeekStartEnd_(sheet, b.weekRow, meta);
    if (!start || !end) return;

    const actionRow = getActionRowForBlock_(sheet, b, colReq, colSub);
    const statusInfo = readStatusForBlock_(sheet, b, colSta, actionRow);
    const status = String(statusInfo.text || '');
    const statusRow = statusInfo.row; // where status currently "lives" / should be written

    const inWeek = now >= start && now <= end;
    const isPast = now > end;
    const isFuture = now < start;

    const isSubmitted = status.startsWith('Submitted') || status.startsWith('Auto-submitted');
    const isRequested = status.includes('Requested');

    // Validate "Unlocked until" — if the date has passed, treat it as expired/locked
    let isUnlockedUntil = false;
    if (status.startsWith('Unlocked until')) {
      const untilDate = parseUnlockUntilDate_(status, now);
      if (untilDate && now.getTime() <= untilDate.getTime()) {
        isUnlockedUntil = true;  // still valid
      } else {
        // Expired — auto-relock so the employee can't edit between processRequests ticks
        setBlockStatus_(sheet, b, colSta, statusRow, 'Locked');
      }
    }

    // ---- 1) Submitted = lock time edits, BUT allow "Request Unlock" (P)
    if (isSubmitted) {
      // Allow request unlock for submitted weeks (current or past), unless already requested/unlocked
      if (!isFuture && !isUnlockedUntil && !isRequested) {
        unprotected.push(sheet.getRange(actionRow, colReq, 1, 1));
      }
      return; // keep time locked + no submit while submitted
    }

    // ---- 2) Time entry editability
    // Allow time edits if current week OR unlocked-until
    if (inWeek || isUnlockedUntil) {
      unprotected.push(sheet.getRange(b.timeRow, timeStart, 1, timeWidth));
      if (inWeek && !isUnlockedUntil && !isRequested) {
        setBlockStatus_(sheet, b, colSta, statusRow, 'Unlocked (current week)');
      }
    } else {
      if (!isRequested) {
        setBlockStatus_(sheet, b, colSta, statusRow, 'Locked');
      }
    }

    // ---- 3) Actions (P/Q) editability — ONLY for the relevant week row
    // Future weeks: nothing
    if (isFuture) return;

    // Current week: Submit only (Q)
    if (inWeek) {
      unprotected.push(sheet.getRange(actionRow, colSub, 1, 1));
      return;
    }

    // Past week:
    // - If unlocked-until => allow Submit (Q)
    // - Else => allow Request Unlock (P) (but not if already requested)
    if (isPast) {
      if (isUnlockedUntil) {
        unprotected.push(sheet.getRange(actionRow, colSub, 1, 1));
      } else {
        if (!isRequested) unprotected.push(sheet.getRange(actionRow, colReq, 1, 1));
      }
    }
  });

  p.setUnprotectedRanges(unprotected);
}

/**
 * Hides and protects internal/system sheets so employees can't edit them.
 * Note: Editors may still be able to unhide, but protection prevents edits.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} empSs
 */
function lockAndHideSystemSheets_(empSs) {
  const cfg = getConfig_();
  const admins = parseEmails_(cfg.ADMIN_EMAILS).map(e => e.toLowerCase());
  const me = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  const allowed = [...new Set([...admins, me].filter(Boolean))]; // only these can edit system sheets

  const systemNames = ['_Meta', '_MonthTemplate', 'Time-Lists'];

  systemNames.forEach(name => {
    const sh = empSs.getSheetByName(name);
    if (!sh) return;

    // Always hide
    try { sh.hideSheet(); } catch (e) {}

    const DESC = `${BRAND.LOCK_NAMESPACE}_SYSTEM_LOCK:${name}`;

    // Check if our protection already exists — if so, skip recreation (saves API quota)
    let existingProtection = null;
    try {
      const allSheet = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (const sp of allSheet) {
        if (sp.getDescription() === DESC) {
          existingProtection = sp;
          break;
        }
      }
    } catch (e) {}

    if (existingProtection) return; // already locked by us — no work needed

    // Remove any existing protections coming from the template or old runs
    try {
      sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    } catch (e) {}
    try {
      sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
    } catch (e) {}

    // Create a fresh sheet protection that locks EVERYTHING
    const p = sh.protect().setDescription(DESC);
    p.setWarningOnly(false);

    // No domain-wide editing (Workspace)
    try { p.setDomainEdit(false); } catch (e) {}

    // Explicitly: nothing is editable for normal editors
    try { p.setUnprotectedRanges([]); } catch (e) {}

    // Remove all current editors from the protection, then add only allowed
    try {
      const current = p.getEditors()
        .map(u => (u.getEmail() || '').toLowerCase())
        .filter(Boolean);

      if (current.length) p.removeEditors(current);
    } catch (e) {}

    try { p.addEditors(allowed); } catch (e) {}
  });
}

/* ===========================
   Requests / Submissions ingestion
   =========================== */
function ingestEmployeeActions_(empSs, emp, now) {
  const cfg = getConfig_();
  const admins = parseEmails_(cfg.ADMIN_EMAILS);

  const ctrl = getControllerSs_();
  const shReq = ctrl.getSheetByName(SHEET_REQUESTS);
  if (!shReq) return;

  const meta = readMeta_(empSs);
  const monthSheets = empSs.getSheets().filter(s => isMonthSheetName_(s.getName(), meta));

  const colReq = colToIndex_(meta.ACTION_COL_REQUEST || 'P');
  const colSub = colToIndex_(meta.ACTION_COL_SUBMIT || 'Q');
  const colSta = colToIndex_(meta.ACTION_COL_STATUS || 'R');

  monthSheets.forEach(ms => {
    const blocks = getWeekBlocks_(ms, meta);

    blocks.forEach(b => {
      const { start, end } = computeWeekStartEnd_(ms, b.weekRow, meta);
      if (!start || !end) return;

      const actionRow = getActionRowForBlock_(ms, b, colReq, colSub);
      const statusInfo = readStatusForBlock_(ms, b, colSta, actionRow);
      const status = String(statusInfo.text || '');
      const statusRow = statusInfo.row;

      const inWeek = now >= start && now <= end;
      const isPast = now > end;
      const isFuture = now < start;

      const isSubmitted = status.startsWith('Submitted') || status.startsWith('Auto-submitted');
      const isRequested = status.includes('Requested');

      // Validate "Unlocked until" expiry (same logic as enforceLocksOnMonthSheet_)
      let isUnlockedUntil = false;
      if (status.startsWith('Unlocked until')) {
        const untilDate = parseUnlockUntilDate_(status, now);
        if (untilDate && now.getTime() <= untilDate.getTime()) {
          isUnlockedUntil = true;
        }
        // If expired, enforceLocksOnMonthSheet_ will handle relocking
      }

      const reqCell = ms.getRange(actionRow, colReq);
      const subCell = ms.getRange(actionRow, colSub);
      const stCell  = ms.getRange(statusRow, colSta);

      const req = reqCell.getValue() === true;
      const sub = subCell.getValue() === true;

      // If already submitted:
      // - ignore Submit clicks
      // - BUT allow Request Unlock (handled by allowReq below)
      if (isSubmitted) {
        if (sub) subCell.setValue(false);
      }

      // Allow Request Unlock if:
      // - week is submitted (current or past) OR week is past+locked
      // - and it's not future, not already unlocked, not already requested
      const allowReq = (!isFuture) && !isUnlockedUntil && !isRequested && (isPast || isSubmitted);

      // Allow Submit only if:
      // - current week OR unlocked-until
      // - and not already submitted
      const allowSub = (!isFuture) && (inWeek || isUnlockedUntil) && !isSubmitted;


      // Clear disallowed clicks immediately
      if (req && !allowReq) reqCell.setValue(false);
      if (sub && !allowSub) subCell.setValue(false);

      // Only proceed if allowed
      const doReq = req && allowReq;
      const doSub = sub && allowSub;

      if (doReq) {
        const baseKey = makeUnlockBaseKey_(emp.code, ms.getName(), b.weekIndex);

        // If one is already pending/approved/unlocked, don't create another row/email
        if (hasOpenUnlockRequest_(shReq, baseKey)) {
          // still reflect "requested" on the sheet (optional)
          if (!status.includes('Requested')) {
            stCell.setValue(isSubmitted ? `${status}\nRequested (Pending)` : 'Requested (Pending)');
          }
          reqCell.setValue(false);
          enforceLocksOnMonthSheet_(ms, now, meta);
          return;
        }

        const requestId = makeUnlockRequestId_(emp.code, ms.getName(), b.weekIndex, now);

        // Always show status + lock immediately
        if (!status.includes('Requested')) {
          stCell.setValue(isSubmitted ? `${status}\nRequested (Pending)` : 'Requested (Pending)');
        }
        enforceLocksOnMonthSheet_(ms, now, meta);

        // Create row + email (every new request is unique)
        appendRequest_(shReq, requestId, emp, ms.getName(), b.weekIndex, start, end, 'Unlock', 'Pending', '');

        if (admins.length) {
          MailApp.sendEmail({
            to: admins.join(','),
            subject: `[Timesheet] Unlock request: ${emp.code} ${ms.getName()} Week ${b.weekIndex}`,
            htmlBody: `<p>Employee <b>${emp.code}</b> requested an unlock.</p>
                      <p><b>Month:</b> ${ms.getName()}<br/>
                      <b>Week:</b> ${b.weekIndex}</p>
                      <p>Approved/Denied in Controller → Requests sheet.</p>`
          });
        }

        // Clear checkbox
        reqCell.setValue(false);
      }


      if (doSub) {
        // Unique ID each time, so re-submissions always create a new row + email
        const requestId = makeSubmitRequestId_(emp.code, ms.getName(), b.weekIndex, now);

        // Update status + lock immediately
        stCell.setValue(`Submitted (${Utilities.formatDate(now, tz_(), 'dd-MMM HH:mm')})`);
        enforceLocksOnMonthSheet_(ms, now, meta);

        // Add a row EVERY time it's submitted
        const note = isUnlockedUntil ? 'Re-submission (after unlock)' : '';
        appendRequest_(shReq, requestId, emp, ms.getName(), b.weekIndex, start, end, 'Submit', 'Submitted', note);

        // Email admins EVERY time it's submitted
        if (admins.length) {
          MailApp.sendEmail({
            to: admins.join(','),
            subject: isUnlockedUntil
              ? `[Timesheet] Re-submitted: ${emp.code} ${ms.getName()} Week ${b.weekIndex}`
              : `[Timesheet] Submitted: ${emp.code} ${ms.getName()} Week ${b.weekIndex}`,
            htmlBody: `<p>Employee <b>${emp.code}</b> submitted their timesheet.</p>
                      <p><b>Month:</b> ${ms.getName()}<br/>
                      <b>Week:</b> ${b.weekIndex}</p>
                      ${isUnlockedUntil ? '<p><i>This was a re-submission after an unlock.</i></p>' : ''}`
          });
        }

        // Clear checkbox
        subCell.setValue(false);
      }
    });
  });
}

function appendRequest_(shReq, requestId, emp, monthSheet, weekIndex, weekStart, weekEnd, type, status, notes) {
  shReq.appendRow([
    requestId,
    new Date(),
    emp.code,
    emp.email,
    monthSheet,
    weekIndex,
    weekStart || '',
    weekEnd || '',
    type,
    status,
    '',
    notes || ''
  ]);
}

/* ===========================
   Admin approvals (Requests sheet) - SAFE
   =========================== */
function processRequests() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;

  try {
    const cfg = getConfig_();
    const unlockHours = Number(cfg.UNLOCK_HOURS || 24);

    const ctrl = getControllerSs_();
    const shReq = ctrl.getSheetByName(SHEET_REQUESTS);
    const shEmp = ctrl.getSheetByName(SHEET_EMPLOYEES);
    if (!shReq || !shEmp) return;

    const reqData = shReq.getDataRange().getValues();
    if (reqData.length < 2) return;

    const headerMap = buildHeaderMap_(shEmp);
    const colCode  = requireCol_(headerMap, 'Employee_Code');
    const colEmail = requireCol_(headerMap, 'Email');
    const colFile  = requireCol_(headerMap, 'Timesheet_File_ID');

    const empData = shEmp.getDataRange().getValues();
    const fileMap = new Map();
    for (let i = 1; i < empData.length; i++) {
      const code  = String(empData[i][colCode - 1] || '').trim();
      const email = String(empData[i][colEmail - 1] || '').trim();
      const fileId= String(empData[i][colFile - 1] || '').trim();
      if (code && fileId) fileMap.set(code, { fileId, email });
    }

    const now = new Date();

    for (let r = 1; r < reqData.length; r++) {
      const [requestId, ts, empCode, empEmail, monthSheet, weekIndex,
        weekStart, weekEnd, type, status, unlockUntil] = reqData[r];

      if (type !== 'Unlock') continue;

      const info = fileMap.get(String(empCode));
      if (!info) continue;

      if (status === 'Approved' && !unlockUntil) {
        const until = new Date(now.getTime() + unlockHours * 3600 * 1000);

        const ok = unlockWeek_(info.fileId, monthSheet, Number(weekIndex), until);
        if (!ok) {
          shReq.getRange(r + 1, 10).setValue('Error');
          shReq.getRange(r + 1, 12).setValue(`Unlock failed: month "${monthSheet}" or week "${weekIndex}" not found in employee file.`);
          continue; // don't email / don't mark unlocked
        }

        shReq.getRange(r + 1, 10).setValue('Unlocked');
        shReq.getRange(r + 1, 11).setValue(until);

        const to = String(empEmail || info.email || '').trim();
        if (to) {
          MailApp.sendEmail({
            to,
            subject: `[Timesheet] Week unlocked (${monthSheet} Week ${weekIndex})`,
            htmlBody: `<p>Your week has been unlocked until <b>${until}</b>.</p>
                      <p>Please fill it in within ${unlockHours} hours.</p>`
          });
        }
      }

      if (status === 'Unlocked' && unlockUntil instanceof Date && unlockUntil.getTime() < now.getTime()) {
        relockWeek_(info.fileId, monthSheet, Number(weekIndex));
        shReq.getRange(r + 1, 10).setValue('Expired');
      }
    }
  } catch (err) {
    console.error('processRequests() fatal error:', err);
  } finally {
    lock.releaseLock();
  }
}

function unlockWeek_(fileId, monthSheetName, weekIndex, until) {
  const empSs = SpreadsheetApp.openById(fileId);
  ensureEmployeeMeta_(empSs);
  const meta = readMeta_(empSs);

  monthSheetName = normalizeMonthSheetName_(monthSheetName, meta);

  const sh = empSs.getSheetByName(monthSheetName);
  if (!sh) return false;

  const colReq = colToIndex_(meta.ACTION_COL_REQUEST || 'P');
  const colSub = colToIndex_(meta.ACTION_COL_SUBMIT || 'Q');
  const colSta = colToIndex_(meta.ACTION_COL_STATUS || 'R');

  const blocks = getWeekBlocks_(sh, meta);
  const b = blocks.find(x => x.weekIndex === weekIndex);
  if (!b) return false;

  const actionRow = getActionRowForBlock_(sh, b, colReq, colSub);
  const statusInfo = readStatusForBlock_(sh, b, colSta, actionRow);

  setBlockStatus_(sh, b, colSta, statusInfo.row,
    `Unlocked until ${Utilities.formatDate(until, tz_(), 'dd-MMM HH:mm')}`
  );

  enforceLocksOnMonthSheet_(sh, new Date(), meta);
  SpreadsheetApp.flush();
  return true;
}

function relockWeek_(fileId, monthSheetName, weekIndex) {
  const empSs = SpreadsheetApp.openById(fileId);
  ensureEmployeeMeta_(empSs);
  const meta = readMeta_(empSs);

  monthSheetName = normalizeMonthSheetName_(monthSheetName, meta);

  const sh = empSs.getSheetByName(monthSheetName);
  if (!sh) return;

  const colReq = colToIndex_(meta.ACTION_COL_REQUEST || 'P');
  const colSub = colToIndex_(meta.ACTION_COL_SUBMIT || 'Q');
  const colSta = colToIndex_(meta.ACTION_COL_STATUS || 'R');

  const blocks = getWeekBlocks_(sh, meta);
  const b = blocks.find(x => x.weekIndex === weekIndex);
  if (!b) return;

  const actionRow = getActionRowForBlock_(sh, b, colReq, colSub);
  const statusInfo = readStatusForBlock_(sh, b, colSta, actionRow);

  setBlockStatus_(sh, b, colSta, statusInfo.row, 'Locked');

  enforceLocksOnMonthSheet_(sh, new Date(), meta);
  SpreadsheetApp.flush();
}

function normalizeMonthSheetName_(v, meta) {
  if (v instanceof Date) {
    const fmt = meta.MONTH_SHEET_NAME_FORMAT || 'MMM-yyyy';
    return Utilities.formatDate(monthStart_(v), tz_(), fmt);
  }
  return String(v || '').trim();
}

function makeSubmitRequestId_(empCode, monthSheetName, weekIndex, when) {
  const ts = Utilities.formatDate(when, tz_(), 'yyyyMMdd-HHmmss');
  return `${empCode} | ${monthSheetName} | W${weekIndex} | SUBMIT | ${ts}`;
}

function makeUnlockBaseKey_(empCode, monthSheetName, weekIndex) {
  return `${empCode} | ${monthSheetName} | W${weekIndex} | UNLOCK`;
}

function makeUnlockRequestId_(empCode, monthSheetName, weekIndex, when) {
  const base = makeUnlockBaseKey_(empCode, monthSheetName, weekIndex);
  const ts = Utilities.formatDate(when, tz_(), 'yyyyMMdd-HHmmss');
  return `${base} | ${ts}`;
}

/**
 * Returns true if there is already an "open" unlock request for that baseKey
 * (Pending / Approved / Unlocked). This prevents duplicate pending rows.
 */
function hasOpenUnlockRequest_(shReq, baseKey) {
  const data = shReq.getDataRange().getValues();
  // Columns based on the Requests sheet:
  // 0=Request ID, 8=Type, 9=Status, 10=Unlock Until
  for (let i = data.length - 1; i >= 1; i--) {
    const rid = String(data[i][0] || '');
    if (!rid.startsWith(baseKey)) continue;

    const type = String(data[i][8] || '');
    if (type !== 'Unlock') continue;

    const status = String(data[i][9] || '');
    if (['Pending', 'Approved', 'Unlocked'].includes(status)) return true;
  }
  return false;
}

/* ===========================
   Auto-submit past unsubmitted weeks
   =========================== */

/**
 * Classifies a week's time row into one of three states:
 *   'has-entries' — at least one genuine time value was entered.
 *   'no-shifts'   — no real times, but at least one cell is the No-Shift
 *                   sentinel (default "Didn't work"); i.e. the employee
 *                   actively declared they had no shifts.
 *   'blank'       — nothing filled in at all (placeholders/empties only);
 *                   i.e. the employee most likely forgot.
 *
 * The sentinel is configurable via _Meta key NO_SHIFT_VALUE so it stays in
 * sync with whatever option you list in the Time-Lists dropdown.
 */
function classifyWeekEntries_(sheet, timeRow, meta) {
  const timeParts = String(meta.TIME_EDIT_RANGE || 'B:O').split(':');
  const timeStart = colToIndex_(timeParts[0]);
  const timeEnd   = colToIndex_(timeParts[1]);
  const width     = (timeEnd - timeStart) + 1;
  const noShift   = String(meta.NO_SHIFT_VALUE || "Didn't work");

  const rowVals = sheet.getRange(timeRow, timeStart, 1, width).getValues()[0];

  const isEmpty = v => v === '' || v === null || v === undefined ||
                       v === 'Pick start time' || v === 'Pick end time';

  let anyReal = false;
  let anyNoShift = false;
  rowVals.forEach(v => {
    if (String(v) === noShift) anyNoShift = true;
    else if (!isEmpty(v)) anyReal = true;
  });

  if (anyReal) return 'has-entries';
  if (anyNoShift) return 'no-shifts';
  return 'blank';
}

/**
 * Back-compat helper: true only when a week is entirely blank.
 */
function isWeekTimeBlank_(sheet, timeRow, meta) {
  return classifyWeekEntries_(sheet, timeRow, meta) === 'blank';
}

/**
 * Automatically submits any past week that hasn't been submitted.
 * - Runs on every tick for Active and Left employees.
 * - Skips weeks that are: already submitted, actively unlocked, or have a pending unlock request.
 * - Sends notification emails to both admins and the employee.
 * - Flags blank weeks in the email.
 */
function autoSubmitPastWeeks_(empSs, emp, now) {
  const cfg = getConfig_();
  const admins = parseEmails_(cfg.ADMIN_EMAILS);

  const ctrl = getControllerSs_();
  const shReq = ctrl.getSheetByName(SHEET_REQUESTS);
  if (!shReq) return;

  const meta = readMeta_(empSs);
  const monthSheets = empSs.getSheets().filter(s => isMonthSheetName_(s.getName(), meta));

  const colReq = colToIndex_(meta.ACTION_COL_REQUEST || 'P');
  const colSub = colToIndex_(meta.ACTION_COL_SUBMIT || 'Q');
  const colSta = colToIndex_(meta.ACTION_COL_STATUS || 'R');

  monthSheets.forEach(ms => {
    const blocks = getWeekBlocks_(ms, meta);

    blocks.forEach(b => {
      const { start, end } = computeWeekStartEnd_(ms, b.weekRow, meta);
      if (!start || !end) return;

      // Only process past weeks
      if (now <= end) return;

      const actionRow = getActionRowForBlock_(ms, b, colReq, colSub);
      const statusInfo = readStatusForBlock_(ms, b, colSta, actionRow);
      const status = String(statusInfo.text || '');
      const statusRow = statusInfo.row;

      // Skip if already submitted
      if (status.startsWith('Submitted')) return;
      if (status.startsWith('Auto-submitted')) return;

      // Skip if actively unlocked (employee was given time to edit)
      if (status.startsWith('Unlocked until')) {
        const untilDate = parseUnlockUntilDate_(status, now);
        if (untilDate && now.getTime() <= untilDate.getTime()) return; // still within unlock window
        // If expired, fall through to auto-submit
      }

      // Skip if there's a pending/approved unlock request (admin is handling it)
      const baseKey = makeUnlockBaseKey_(emp.code, ms.getName(), b.weekIndex);
      if (hasOpenUnlockRequest_(shReq, baseKey)) return;

      // ---- Auto-submit this week ----
      const entryState = classifyWeekEntries_(ms, b.timeRow, meta);
      const blank    = entryState === 'blank';      // employee likely forgot
      const noShifts = entryState === 'no-shifts';  // employee declared no shifts
      const stateFlag = blank ? ' (NO TIME ENTRIES)' : (noShifts ? ' (NO SHIFTS)' : '');
      const stampText = Utilities.formatDate(now, tz_(), 'dd-MMM HH:mm');

      // Update status on the sheet
      setBlockStatus_(ms, b, colSta, statusRow, `Auto-submitted (${stampText})${stateFlag}`);

      // Clear any leftover checkbox states
      try { ms.getRange(actionRow, colReq).setValue(false); } catch (e) {}
      try { ms.getRange(actionRow, colSub).setValue(false); } catch (e) {}

      // Re-enforce locks so the week is now protected
      enforceLocksOnMonthSheet_(ms, now, meta);

      // Create a request row
      const requestId = makeSubmitRequestId_(emp.code, ms.getName(), b.weekIndex, now);
      const noteText =
        blank    ? 'Auto-submitted by automation (week had NO time entries)' :
        noShifts ? 'Auto-submitted by automation (employee marked all days as no shift)' :
                   'Auto-submitted by automation';
      appendRequest_(shReq, requestId, emp, ms.getName(), b.weekIndex, start, end, 'Submit', 'Submitted', noteText);

      // Build email content
      const monthLabel = ms.getName();
      const weekLabel  = `Week ${b.weekIndex}`;
      // Blank = actionable problem (red). No-shifts = informational (neutral).
      const stateNotice = blank
        ? '<p style="color:#d93025;"><b>Warning:</b> This week had <b>no time entries</b>. ' +
          'If this is incorrect, please request an unlock from your manager.</p>'
        : (noShifts
            ? '<p style="color:#555;">All days this week were marked as <b>no shift</b>; ' +
              'no hours were recorded.</p>'
            : '');

      // Email admins
      if (admins.length) {
        MailApp.sendEmail({
          to: admins.join(','),
          subject: `[Timesheet] Auto-submitted: ${emp.code} ${monthLabel} ${weekLabel}${stateFlag}`,
          htmlBody:
            `<p>Employee <b>${emp.code}</b> (${emp.first || ''} ${emp.last || ''}) ` +
            `did not submit their timesheet before the week ended.</p>` +
            `<p>The timesheet has been <b>automatically submitted</b> by the system.</p>` +
            `<p><b>Month:</b> ${monthLabel}<br/>` +
            `<b>Week:</b> ${b.weekIndex}<br/>` +
            `<b>Period:</b> ${Utilities.formatDate(start, tz_(), 'dd/MM/yyyy')} – ${Utilities.formatDate(end, tz_(), 'dd/MM/yyyy')}</p>` +
            stateNotice
        });
      }

      // Email the employee
      if (emp.email) {
        const fileLink = `https://docs.google.com/spreadsheets/d/${emp.fileId}/edit`;
        MailApp.sendEmail({
          to: emp.email,
          subject: `[Timesheet] Your timesheet was auto-submitted: ${monthLabel} ${weekLabel}`,
          htmlBody:
            `<p>Hi ${emp.first || ''},</p>` +
            `<p>Your timesheet for <b>${monthLabel} ${weekLabel}</b> ` +
            `(${Utilities.formatDate(start, tz_(), 'dd/MM/yyyy')} – ${Utilities.formatDate(end, tz_(), 'dd/MM/yyyy')}) ` +
            `was not submitted before the week ended.</p>` +
            `<p>It has been <b>automatically submitted</b> by the system.</p>` +
            stateNotice +
            `<p>If you need to make changes, please request an unlock from your manager.</p>` +
            `<p><a href="${fileLink}">Open your timesheet</a></p>` +
            `<p>Thanks!</p>`
        });
      }
    });
  });
}

/* ===========================
   Reminders - SAFE
   =========================== */
function sendReminders() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;

  try {
    const now = new Date();
    const employees = loadEmployees_();

    employees.forEach(emp => {
      try {
        if (emp.status !== STATUS_ACTIVE) return;
        if (!emp.email) return;

        const shareStart = computeShareStart_(emp, now);
        if (now < shareStart) return;

        if (!emp.fileId) return;

        const empSs = SpreadsheetApp.openById(emp.fileId);
        enforceEmployeeHiddenSheets_(empSs);

        ensureEmployeeMeta_(empSs);
        const meta = readMeta_(empSs);

        const fmt = meta.MONTH_SHEET_NAME_FORMAT || 'MMM-yyyy';
        const monthName = Utilities.formatDate(monthStart_(now), tz_(), fmt);
        const sh = empSs.getSheetByName(monthName);
        if (!sh) return;

        const blocks = getWeekBlocks_(sh, meta);
        const colSta = colToIndex_(meta.ACTION_COL_STATUS || 'R');

        const colReq = colToIndex_(meta.ACTION_COL_REQUEST || 'P');
        const colSub = colToIndex_(meta.ACTION_COL_SUBMIT || 'Q');

        for (const b of blocks) {
          const { start, end } = computeWeekStartEnd_(sh, b.weekRow, meta);
          if (!start || !end) continue;
          if (now < start || now > end) continue;

          const actionRow = getActionRowForBlock_(sh, b, colReq, colSub);
          const statusInfo = readStatusForBlock_(sh, b, colSta, actionRow);
          const status = String(statusInfo.text || '');

          if (status.startsWith('Submitted') || status.startsWith('Auto-submitted')) return;

          const timeParts = String(meta.TIME_EDIT_RANGE || 'B:O').split(':');
          const timeStart = colToIndex_(timeParts[0]);
          const timeEnd   = colToIndex_(timeParts[1]);
          const width      = (timeEnd - timeStart) + 1;

          const rowVals = sh.getRange(b.timeRow, timeStart, 1, width).getValues()[0];
          const hasPlaceholders = rowVals.some(v =>
            v === 'Pick start time' || v === 'Pick end time' || v === '' || v === null
          );

          if (hasPlaceholders) {
            MailApp.sendEmail({
              to: emp.email,
              subject: `[Timesheet Reminder] Please complete your timesheet`,
              htmlBody: `<p>Hi ${emp.first || ''},</p>
                         <p>This is a reminder to complete your timesheet for this week.</p>
                         <p>Thanks!</p>`
            });
          }
          return;
        }
      } catch (empErr) {
        console.error(`sendReminders(): employee ${emp.code || '(no code)'} failed:`, empErr);
      }
    });
  } catch (err) {
    console.error('sendReminders() fatal error:', err);
  } finally {
    lock.releaseLock();
  }
}