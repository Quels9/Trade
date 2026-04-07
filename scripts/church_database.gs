/**
 * Church Database Automation for Google Sheets
 *
 * Setup:
 * 1) Open Extensions -> Apps Script in your spreadsheet.
 * 2) Paste this file into Code.gs.
 * 3) Update SHEET_NAMES / prefixes if you renamed tabs.
 * 4) Run setupChurchDb() once.
 */

const SHEET_NAMES = {
  MEMBERS: 'Members',
  FAMILIES: 'Families',
  ATTENDANCE: 'Attendance',
  GIVING: 'Giving',
  NOTES: 'Notes',
};

const ID_PREFIX = {
  Members: 'M',
  Families: 'F',
  Attendance: 'A',
  Giving: 'G',
  Notes: 'N',
};

/**
 * Install trigger and initialize document properties.
 */
function setupChurchDb() {
  const ss = SpreadsheetApp.getActive();
  const triggerExists = ScriptApp.getProjectTriggers().some(
    t => t.getHandlerFunction() === 'onEdit'
  );

  if (!triggerExists) {
    ScriptApp.newTrigger('onEdit').forSpreadsheet(ss).onEdit().create();
  }

  SpreadsheetApp.getUi().alert('Setup complete: onEdit trigger is ready.');
}

/**
 * Main automation hook.
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();

  if (row < 2) return; // ignore headers

  if (!Object.keys(ID_PREFIX).includes(sheetName)) return;

  ensureId_(sheet, row, sheetName);

  if (sheetName === SHEET_NAMES.MEMBERS) {
    handleMembersEdit_(sheet, row, e.range.getColumn());
  }

  if (sheetName === SHEET_NAMES.ATTENDANCE) {
    handleAttendanceEdit_(sheet, row);
  }
}

function ensureId_(sheet, row, sheetName) {
  const idCell = sheet.getRange(row, 1); // column A = ID
  const existing = String(idCell.getValue() || '').trim();
  if (existing) return;

  const next = nextSequence_(sheetName);
  const prefix = ID_PREFIX[sheetName];
  const id = `${prefix}${String(next).padStart(5, '0')}`;
  idCell.setValue(id);

  setCreatedAndUpdated_(sheet, row);
}

function setCreatedAndUpdated_(sheet, row) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const createdIdx = headers.findIndex(h => String(h).toLowerCase() === 'created_at');
  const updatedIdx = headers.findIndex(h => String(h).toLowerCase() === 'updated_at');

  const now = new Date();

  if (createdIdx >= 0) {
    const createdCell = sheet.getRange(row, createdIdx + 1);
    if (!createdCell.getValue()) createdCell.setValue(now);
  }

  if (updatedIdx >= 0) {
    sheet.getRange(row, updatedIdx + 1).setValue(now);
  }
}

function handleMembersEdit_(sheet, row, col) {
  setCreatedAndUpdated_(sheet, row);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const emailCol = headers.indexOf('email') + 1;
  const phoneCol = headers.indexOf('phone') + 1;

  if (col === emailCol) checkDuplicateInColumn_(sheet, emailCol, row, 'email');
  if (col === phoneCol) checkDuplicateInColumn_(sheet, phoneCol, row, 'phone');
}

function handleAttendanceEdit_(attendanceSheet, row) {
  const headers = attendanceSheet.getRange(1, 1, 1, attendanceSheet.getLastColumn()).getValues()[0];
  const memberIdCol = headers.indexOf('member_id') + 1;
  const dateCol = headers.indexOf('date') + 1;
  const presentCol = headers.indexOf('present') + 1;

  if (!memberIdCol || !dateCol || !presentCol) return;

  const memberId = String(attendanceSheet.getRange(row, memberIdCol).getValue() || '').trim();
  const dateVal = attendanceSheet.getRange(row, dateCol).getValue();
  const presentVal = attendanceSheet.getRange(row, presentCol).getValue();

  if (!memberId || !dateVal || presentVal !== true) return;

  const membersSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.MEMBERS);
  if (!membersSheet) return;

  const memberHeaders = membersSheet.getRange(1, 1, 1, membersSheet.getLastColumn()).getValues()[0];
  const memberIdMembersCol = memberHeaders.indexOf('member_id') + 1;
  const lastAttendanceCol = memberHeaders.indexOf('last_attendance_date') + 1;
  const updatedAtCol = memberHeaders.indexOf('updated_at') + 1;

  if (!memberIdMembersCol || !lastAttendanceCol) return;

  const values = membersSheet.getRange(2, memberIdMembersCol, Math.max(0, membersSheet.getLastRow() - 1), 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === memberId) {
      const targetRow = i + 2;
      membersSheet.getRange(targetRow, lastAttendanceCol).setValue(dateVal);
      if (updatedAtCol) membersSheet.getRange(targetRow, updatedAtCol).setValue(new Date());
      break;
    }
  }

  setCreatedAndUpdated_(attendanceSheet, row);
}

function checkDuplicateInColumn_(sheet, col, row, label) {
  const input = String(sheet.getRange(row, col).getValue() || '').trim().toLowerCase();
  if (!input) return;

  const values = sheet.getRange(2, col, Math.max(0, sheet.getLastRow() - 1), 1).getValues()
    .map(r => String(r[0] || '').trim().toLowerCase());

  const duplicates = values.filter(v => v && v === input).length;

  const noteCell = sheet.getRange(row, col);
  if (duplicates > 1) {
    noteCell.setNote(`Duplicate ${label} detected. Please verify this record.`);
    SpreadsheetApp.getActive().toast(`Duplicate ${label} detected on row ${row}.`, 'Church DB', 5);
  } else {
    noteCell.setNote('');
  }
}

function nextSequence_(sheetName) {
  const props = PropertiesService.getDocumentProperties();
  const key = `SEQ_${sheetName.toUpperCase()}`;
  const current = Number(props.getProperty(key) || 0);
  const next = current + 1;
  props.setProperty(key, String(next));
  return next;
}
