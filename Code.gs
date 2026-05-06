/**
 * Case Dashboard Customizer & Tracker
 * Google Apps Script web app backed by multiple Google Sheets.
 *
 * Setup:
 *   1. Create a Google Sheet (or use the bound spreadsheet).
 *   2. Set the script property `SPREADSHEET_ID` to that sheet's id, OR
 *      bind this script to the sheet (Extensions -> Apps Script).
 *   3. Deploy -> New deployment -> Web app.
 *
 * Sheets (auto-created on first run):
 *   - Cases:        ID | Initials | MainLink | <dynamic custom columns...>
 *   - TeamMembers:  Initials | Name | Role
 *   - Columns:      ColumnId | ColumnName | ColumnType | Options | Order
 */

const CASES_SHEET = 'Cases';
const TEAM_SHEET = 'TeamMembers';
const COLUMNS_SHEET = 'Columns';

const CASES_FIXED_HEADERS = ['ID', 'Initials', 'MainLink'];
const TEAM_HEADERS = ['Initials', 'Name', 'Role'];
const COLUMNS_HEADERS = ['ColumnId', 'ColumnName', 'ColumnType', 'Options', 'Order'];

const COLUMN_TYPES = ['text', 'number', 'date', 'link', 'dropdown', 'checkbox'];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Case Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error('No spreadsheet bound. Set the SPREADSHEET_ID script property.');
}

function getOrCreateSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function readSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { headers: [], rows: [] };
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(String);
  const rows = values.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });
  return { headers, rows };
}

function ensureSchema_() {
  const ss = getSpreadsheet_();
  getOrCreateSheet_(ss, CASES_SHEET, CASES_FIXED_HEADERS);
  getOrCreateSheet_(ss, TEAM_SHEET, TEAM_HEADERS);
  getOrCreateSheet_(ss, COLUMNS_SHEET, COLUMNS_HEADERS);
  syncCasesHeaders_();
}

function syncCasesHeaders_() {
  const ss = getSpreadsheet_();
  const cases = ss.getSheetByName(CASES_SHEET);
  const cols = getCustomColumns_();
  const desired = CASES_FIXED_HEADERS.concat(cols.map(c => c.ColumnName));
  const lastCol = Math.max(cases.getLastColumn(), 1);
  const current = cases.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const trimmed = current.filter(h => h !== '');
  const same = trimmed.length === desired.length && trimmed.every((h, i) => h === desired[i]);
  if (same) return;

  // Map existing rows to objects keyed by current header, then rewrite with desired headers.
  const lastRow = cases.getLastRow();
  let dataRows = [];
  if (lastRow > 1) {
    const values = cases.getRange(2, 1, lastRow - 1, lastCol).getValues();
    dataRows = values
      .filter(r => r.some(c => c !== '' && c !== null))
      .map(r => {
        const o = {};
        trimmed.forEach((h, i) => { o[h] = r[i]; });
        return o;
      });
  }
  cases.clear();
  cases.getRange(1, 1, 1, desired.length).setValues([desired]);
  cases.setFrozenRows(1);
  cases.getRange(1, 1, 1, desired.length).setFontWeight('bold');
  if (dataRows.length) {
    const out = dataRows.map(o => desired.map(h => (h in o ? o[h] : '')));
    cases.getRange(2, 1, out.length, desired.length).setValues(out);
  }
}

/* ---------- Bootstrap ---------- */

function getDashboardData() {
  ensureSchema_();
  return {
    cases: getCases(),
    team: getTeamMembers(),
    columns: getCustomColumns(),
    columnTypes: COLUMN_TYPES
  };
}

/* ---------- Cases ---------- */

function getCases() {
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, CASES_SHEET, CASES_FIXED_HEADERS);
  return readSheet_(sheet).rows.map(r => {
    if (r.MainLink instanceof Date) r.MainLink = r.MainLink.toISOString();
    Object.keys(r).forEach(k => { if (r[k] instanceof Date) r[k] = r[k].toISOString(); });
    return r;
  });
}

function addCase(payload) {
  ensureSchema_();
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CASES_SHEET);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const id = payload.ID || Utilities.getUuid();
  const row = headers.map(h => {
    if (h === 'ID') return id;
    return payload[h] != null ? payload[h] : '';
  });
  sheet.appendRow(row);
  return { ok: true, id };
}

function updateCase(payload) {
  if (!payload || !payload.ID) throw new Error('updateCase requires ID');
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CASES_SHEET);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0];
  const idCol = headers.indexOf('ID');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(payload.ID)) {
      const row = headers.map(h => (h in payload ? payload[h] : values[i][headers.indexOf(h)]));
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      return { ok: true };
    }
  }
  throw new Error('Case not found: ' + payload.ID);
}

function deleteCase(id) {
  if (!id) throw new Error('deleteCase requires id');
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CASES_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true };
  const values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  const idCol = values[0].indexOf('ID');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false };
}

/* ---------- Team Members ---------- */

function getTeamMembers() {
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, TEAM_SHEET, TEAM_HEADERS);
  return readSheet_(sheet).rows;
}

function saveTeamMember(payload) {
  ensureSchema_();
  if (!payload || !payload.Initials) throw new Error('Initials are required');
  const initials = String(payload.Initials).trim().toUpperCase();
  const name = String(payload.Name || '').trim();
  const role = String(payload.Role || '').trim();
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(TEAM_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const range = sheet.getRange(2, 1, lastRow - 1, TEAM_HEADERS.length).getValues();
    for (let i = 0; i < range.length; i++) {
      if (String(range[i][0]).trim().toUpperCase() === initials) {
        sheet.getRange(i + 2, 1, 1, TEAM_HEADERS.length).setValues([[initials, name, role]]);
        return { ok: true, updated: true };
      }
    }
  }
  sheet.appendRow([initials, name, role]);
  return { ok: true, created: true };
}

/**
 * Update an existing team member identified by OriginalInitials.
 * If the initials change, the row is renamed and any cases that reference
 * the old initials are updated to the new ones to preserve linkage.
 */
function updateTeamMember(payload) {
  ensureSchema_();
  if (!payload || !payload.OriginalInitials) {
    throw new Error('updateTeamMember requires OriginalInitials');
  }
  if (!payload.Initials) throw new Error('Initials are required');

  const original = String(payload.OriginalInitials).trim().toUpperCase();
  const newInitials = String(payload.Initials).trim().toUpperCase();
  const name = String(payload.Name || '').trim();
  const role = String(payload.Role || '').trim();
  if (!newInitials) throw new Error('Initials cannot be empty');

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(TEAM_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No team members to update');

  const values = sheet.getRange(2, 1, lastRow - 1, TEAM_HEADERS.length).getValues();

  if (original !== newInitials) {
    const conflict = values.some(r => String(r[0]).trim().toUpperCase() === newInitials);
    if (conflict) throw new Error('A team member with initials "' + newInitials + '" already exists');
  }

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === original) {
      sheet.getRange(i + 2, 1, 1, TEAM_HEADERS.length).setValues([[newInitials, name, role]]);
      if (original !== newInitials) updateCasesInitials_(original, newInitials);
      return { ok: true };
    }
  }
  throw new Error('Team member not found: ' + original);
}

function updateCasesInitials_(oldInitials, newInitials) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CASES_SHEET);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.indexOf('Initials');
  if (idx === -1) return;
  const range = sheet.getRange(2, idx + 1, lastRow - 1, 1);
  const cells = range.getValues();
  let changed = false;
  for (let i = 0; i < cells.length; i++) {
    if (String(cells[i][0]).trim().toUpperCase() === oldInitials) {
      cells[i][0] = newInitials;
      changed = true;
    }
  }
  if (changed) range.setValues(cells);
}

function deleteTeamMember(initials) {
  if (!initials) throw new Error('initials required');
  const target = String(initials).trim().toUpperCase();
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(TEAM_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true };
  const values = sheet.getRange(2, 1, lastRow - 1, TEAM_HEADERS.length).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === target) {
      sheet.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: false };
}

/* ---------- Custom Columns ---------- */

function getCustomColumns_() {
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, COLUMNS_SHEET, COLUMNS_HEADERS);
  const rows = readSheet_(sheet).rows;
  rows.sort((a, b) => (Number(a.Order) || 0) - (Number(b.Order) || 0));
  return rows;
}

function getCustomColumns() {
  return getCustomColumns_();
}

function saveColumn(payload) {
  ensureSchema_();
  if (!payload || !payload.ColumnName) throw new Error('ColumnName is required');
  const type = String(payload.ColumnType || 'text').toLowerCase();
  if (COLUMN_TYPES.indexOf(type) === -1) throw new Error('Invalid column type: ' + type);
  const name = String(payload.ColumnName).trim();
  if (CASES_FIXED_HEADERS.indexOf(name) !== -1) {
    throw new Error('Column name conflicts with a built-in column');
  }
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(COLUMNS_SHEET);
  const lastRow = sheet.getLastRow();
  const id = payload.ColumnId || Utilities.getUuid();
  const order = payload.Order != null && payload.Order !== '' ? Number(payload.Order) : (lastRow);
  const options = String(payload.Options || '').trim();

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, COLUMNS_HEADERS.length).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]) === id) {
        const oldName = String(values[i][1]);
        sheet.getRange(i + 2, 1, 1, COLUMNS_HEADERS.length).setValues([[id, name, type, options, order]]);
        if (oldName !== name) renameCasesHeader_(oldName, name);
        syncCasesHeaders_();
        return { ok: true, id };
      }
    }
  }
  sheet.appendRow([id, name, type, options, order]);
  syncCasesHeaders_();
  return { ok: true, id };
}

function deleteColumn(columnId) {
  if (!columnId) throw new Error('columnId required');
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(COLUMNS_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true };
  const values = sheet.getRange(2, 1, lastRow - 1, COLUMNS_HEADERS.length).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(columnId)) {
      const colName = String(values[i][1]);
      sheet.deleteRow(i + 2);
      removeCasesColumn_(colName);
      return { ok: true };
    }
  }
  return { ok: false };
}

function reorderColumns(orderedIds) {
  if (!Array.isArray(orderedIds)) throw new Error('orderedIds must be an array');
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(COLUMNS_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true };
  const range = sheet.getRange(2, 1, lastRow - 1, COLUMNS_HEADERS.length);
  const values = range.getValues();
  values.forEach(row => {
    const idx = orderedIds.indexOf(String(row[0]));
    row[4] = idx === -1 ? 999 : idx;
  });
  range.setValues(values);
  syncCasesHeaders_();
  return { ok: true };
}

function renameCasesHeader_(oldName, newName) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CASES_SHEET);
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  const headers = headerRange.getValues()[0];
  const idx = headers.indexOf(oldName);
  if (idx !== -1) {
    headers[idx] = newName;
    headerRange.setValues([headers]);
  }
}

function removeCasesColumn_(name) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CASES_SHEET);
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.indexOf(name);
  if (idx !== -1) sheet.deleteColumn(idx + 1);
}
