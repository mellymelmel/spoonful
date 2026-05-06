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
 *   - Cases:        ID | MemberId | MainLink | <dynamic custom columns...>
 *   - TeamMembers:  MemberId | Initials | Name | Role
 *   - Columns:      ColumnId | ColumnName | ColumnType | Options | Order
 *
 * Cases link to team members by MemberId (required). Initials are kept as a
 * display label only; renaming initials does not break case linkage.
 */

const CASES_SHEET = 'Cases';
const TEAM_SHEET = 'TeamMembers';
const COLUMNS_SHEET = 'Columns';

const CASES_FIXED_HEADERS = ['ID', 'MemberId', 'MainLink'];
const TEAM_HEADERS = ['MemberId', 'Initials', 'Name', 'Role'];
const COLUMNS_HEADERS = ['ColumnId', 'ColumnName', 'DisplayName', 'ColumnType', 'Options', 'Order'];

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
  migrateInitialsToMemberId_();
  migrateColumnsAddDisplayName_();
  ensureCasesHeaders_(ss.getSheetByName(CASES_SHEET));
}

/**
 * Add a DisplayName column to the Columns sheet for projects created before
 * the display-name feature. Defaults DisplayName to ColumnName for each row.
 */
function migrateColumnsAddDisplayName_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(COLUMNS_SHEET);
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  if (headers.indexOf('DisplayName') !== -1) return;
  sheet.insertColumnAfter(2);
  sheet.getRange(1, 3).setValue('DisplayName').setFontWeight('bold');
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const names = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    sheet.getRange(2, 3, lastRow - 1, 1).setValues(names);
  }
}

/**
 * Upgrade sheets created under the previous schema (Initials as primary key)
 * by introducing MemberId as the foreign key in Cases and primary key in
 * TeamMembers. Idempotent: returns early if MemberId is already present.
 */
function migrateInitialsToMemberId_() {
  const ss = getSpreadsheet_();
  const team = ss.getSheetByName(TEAM_SHEET);
  if (!team) return;
  const lastCol = team.getLastColumn();
  if (lastCol < 1) return;
  const headers = team.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  if (headers.indexOf('MemberId') !== -1) return; // already migrated
  if (headers[0] !== 'Initials') return; // unknown shape; leave alone

  const lastRow = team.getLastRow();
  const initialsToId = {};
  team.insertColumnBefore(1);
  team.getRange(1, 1).setValue('MemberId');
  team.getRange(1, 1).setFontWeight('bold');
  if (lastRow >= 2) {
    const initialsCol = team.getRange(2, 2, lastRow - 1, 1).getValues();
    const ids = [];
    for (let i = 0; i < initialsCol.length; i++) {
      const init = String(initialsCol[i][0]).trim().toUpperCase();
      const id = Utilities.getUuid();
      ids.push([id]);
      if (init) initialsToId[init] = id;
    }
    team.getRange(2, 1, ids.length, 1).setValues(ids);
  }

  const cases = ss.getSheetByName(CASES_SHEET);
  if (!cases) return;
  const cLastCol = cases.getLastColumn();
  if (cLastCol < 1) return;
  const cHeaders = cases.getRange(1, 1, 1, cLastCol).getValues()[0].map(String);
  const idx = cHeaders.indexOf('Initials');
  if (idx === -1) return;
  cases.getRange(1, idx + 1).setValue('MemberId');
  const cLastRow = cases.getLastRow();
  if (cLastRow >= 2) {
    const range = cases.getRange(2, idx + 1, cLastRow - 1, 1);
    const vals = range.getValues();
    for (let i = 0; i < vals.length; i++) {
      const init = String(vals[i][0]).trim().toUpperCase();
      if (init && initialsToId[init]) vals[i][0] = initialsToId[init];
    }
    range.setValues(vals);
  }
}

/**
 * Make sure every designated header (fixed + configured custom columns)
 * exists in the Cases sheet. Additive only: never deletes or reorders
 * existing columns. Extra columns the user has in the sheet are preserved.
 */
function ensureCasesHeaders_(sheet) {
  if (!sheet) return;
  const customCols = getCustomColumns_().map(c => c.ColumnName);
  const desired = CASES_FIXED_HEADERS.concat(customCols);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  let current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  while (current.length && current[current.length - 1] === '') current.pop();
  if (current.length === 0) {
    sheet.getRange(1, 1, 1, desired.length).setValues([desired]);
    sheet.getRange(1, 1, 1, desired.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return;
  }
  const missing = desired.filter(d => current.indexOf(d) === -1);
  if (missing.length === 0) return;
  const newHeaders = current.concat(missing);
  sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  sheet.getRange(1, 1, 1, newHeaders.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/* ---------- Bootstrap ---------- */

function getDashboardData() {
  ensureSchema_();
  const auto = autoReconcileDuplicates();
  return {
    cases: getCases(),
    team: getTeamMembers(),
    columns: getCustomColumns(),
    columnTypes: COLUMN_TYPES,
    autoReconciled: auto
  };
}

/* ---------- Cases ---------- */

/**
 * Return cases projected to the designated columns only (fixed + configured
 * custom columns). Extra columns present in the Cases sheet are not exposed.
 */
function getCases() {
  ensureSchema_();
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, CASES_SHEET, CASES_FIXED_HEADERS);
  const designated = CASES_FIXED_HEADERS.concat(getCustomColumns_().map(c => c.ColumnName));
  return readSheet_(sheet).rows.map(r => {
    const out = {};
    designated.forEach(k => {
      let v = r[k];
      if (v instanceof Date) v = v.toISOString();
      out[k] = v == null ? '' : v;
    });
    return out;
  });
}

function addCase(payload) {
  ensureSchema_();
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CASES_SHEET);
  ensureCasesHeaders_(sheet);
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

/**
 * Create or update a team member, keyed by MemberId.
 * MemberId is required; Initials and Name are required for display.
 */
function saveTeamMember(payload) {
  ensureSchema_();
  if (!payload) throw new Error('payload required');
  const memberId = String(payload.MemberId || '').trim();
  if (!memberId) throw new Error('MemberId is required');
  const initials = String(payload.Initials || '').trim().toUpperCase();
  const name = String(payload.Name || '').trim();
  const role = String(payload.Role || '').trim();
  if (!initials) throw new Error('Initials are required');
  if (!name) throw new Error('Name is required');

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(TEAM_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, TEAM_HEADERS.length).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === memberId) {
        sheet.getRange(i + 2, 1, 1, TEAM_HEADERS.length)
          .setValues([[memberId, initials, name, role]]);
        return { ok: true, updated: true };
      }
    }
  }
  sheet.appendRow([memberId, initials, name, role]);
  return { ok: true, created: true };
}

/**
 * Update a team member identified by OriginalMemberId. If MemberId changes,
 * the row is renamed and any cases linked by the old MemberId are updated
 * so linkage is preserved.
 */
function updateTeamMember(payload) {
  ensureSchema_();
  if (!payload || !payload.OriginalMemberId) {
    throw new Error('updateTeamMember requires OriginalMemberId');
  }
  const original = String(payload.OriginalMemberId).trim();
  const newId = String(payload.MemberId || '').trim();
  const initials = String(payload.Initials || '').trim().toUpperCase();
  const name = String(payload.Name || '').trim();
  const role = String(payload.Role || '').trim();
  if (!newId) throw new Error('MemberId is required');
  if (!initials) throw new Error('Initials are required');
  if (!name) throw new Error('Name is required');

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(TEAM_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No team members to update');
  const values = sheet.getRange(2, 1, lastRow - 1, TEAM_HEADERS.length).getValues();

  if (original !== newId) {
    const conflict = values.some(r => String(r[0]).trim() === newId);
    if (conflict) throw new Error('A team member with MemberId "' + newId + '" already exists');
  }

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === original) {
      sheet.getRange(i + 2, 1, 1, TEAM_HEADERS.length)
        .setValues([[newId, initials, name, role]]);
      if (original !== newId) updateCasesMemberId_(original, newId);
      return { ok: true };
    }
  }
  throw new Error('Team member not found: ' + original);
}

function updateCasesMemberId_(oldId, newId) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CASES_SHEET);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.indexOf('MemberId');
  if (idx === -1) return;
  const range = sheet.getRange(2, idx + 1, lastRow - 1, 1);
  const cells = range.getValues();
  let changed = false;
  for (let i = 0; i < cells.length; i++) {
    if (String(cells[i][0]).trim() === oldId) {
      cells[i][0] = newId;
      changed = true;
    }
  }
  if (changed) range.setValues(cells);
}

function deleteTeamMember(memberId) {
  if (!memberId) throw new Error('memberId required');
  const target = String(memberId).trim();
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(TEAM_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true };
  const values = sheet.getRange(2, 1, lastRow - 1, TEAM_HEADERS.length).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === target) {
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

/**
 * ColumnName is the actual header in the Cases sheet (the "key"). Edits to it
 * never rename the underlying sheet header — they just point this dashboard
 * column at a different sheet column. DisplayName is the cosmetic label shown
 * in the dashboard; defaults to ColumnName when blank.
 */
function saveColumn(payload) {
  ensureSchema_();
  if (!payload || !payload.ColumnName) throw new Error('Sheet column name is required');
  const type = String(payload.ColumnType || 'text').toLowerCase();
  if (COLUMN_TYPES.indexOf(type) === -1) throw new Error('Invalid column type: ' + type);
  const name = String(payload.ColumnName).trim();
  const displayName = String(payload.DisplayName == null ? '' : payload.DisplayName).trim() || name;
  if (CASES_FIXED_HEADERS.indexOf(name) !== -1) {
    throw new Error('"' + name + '" is a built-in column name');
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
        sheet.getRange(i + 2, 1, 1, COLUMNS_HEADERS.length)
          .setValues([[id, name, displayName, type, options, order]]);
        ensureCasesHeaders_(ss.getSheetByName(CASES_SHEET));
        return { ok: true, id };
      }
    }
  }
  sheet.appendRow([id, name, displayName, type, options, order]);
  ensureCasesHeaders_(ss.getSheetByName(CASES_SHEET));
  return { ok: true, id };
}

/**
 * Remove a column from the dashboard config. Data in the underlying Cases
 * sheet is preserved — the column simply stops being displayed.
 */
function deleteColumn(columnId) {
  if (!columnId) throw new Error('columnId required');
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(COLUMNS_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true };
  const values = sheet.getRange(2, 1, lastRow - 1, COLUMNS_HEADERS.length).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(columnId)) {
      sheet.deleteRow(i + 2);
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
  return { ok: true };
}

/* ---------- Duplicate detection ---------- */

/**
 * Auto-resolve duplicate groups where every designated column value matches
 * across all rows in the group. The first row is kept; the rest are deleted.
 * Groups with any differing column are left alone for manual reconciliation.
 * Returns { deleted, groups } counts.
 */
function autoReconcileDuplicates() {
  const groups = findDuplicateCases();
  if (!groups || groups.length === 0) return { deleted: 0, groups: 0 };
  const designated = CASES_FIXED_HEADERS.concat(getCustomColumns_().map(c => c.ColumnName));
  const compareKeys = designated.filter(k => k !== 'ID');
  let deleted = 0;
  let groupsResolved = 0;
  groups.forEach(g => {
    const rows = g.rows;
    const first = rows[0];
    const allMatch = rows.every(r =>
      compareKeys.every(k => normalizeForCompare_(r[k]) === normalizeForCompare_(first[k]))
    );
    if (!allMatch) return;
    for (let i = 1; i < rows.length; i++) {
      deleteCase(rows[i].ID);
      deleted++;
    }
    groupsResolved++;
  });
  return { deleted: deleted, groups: groupsResolved };
}

function normalizeForCompare_(v) {
  if (v == null) return '';
  return String(v).trim();
}

/**
 * Group cases that share a Main Link (case-insensitive, ignoring trailing
 * slashes). Empty/missing Main Links are skipped. Each returned group has
 * 2+ rows. Useful for the Reconcile Duplicates flow.
 */
function findDuplicateCases() {
  const cases = getCases();
  const groups = {};
  cases.forEach(c => {
    const link = String(c.MainLink || '').trim().toLowerCase().replace(/\/+$/, '');
    if (!link) return;
    if (!groups[link]) groups[link] = [];
    groups[link].push(c);
  });
  const out = [];
  Object.keys(groups).forEach(k => {
    if (groups[k].length > 1) out.push({ key: k, rows: groups[k] });
  });
  return out;
}

/* ---------- Bulk column import ----------
 * Accepts either a Google Sheets URL (+ optional tab) or pasted CSV/TSV.
 * Source rows are expected to have headers: "Column Name", "Type", "Options".
 * Header matching is loose (case-insensitive substring). If no headers are
 * detected the first column is treated as the name and type defaults to text.
 */

function previewColumnsFromSheet(url, tabName) {
  const rows = readRowsFromSourceSheet_(url, tabName);
  return parseColumnRows_(rows);
}

function importColumnsFromSheet(url, tabName) {
  const parsed = previewColumnsFromSheet(url, tabName);
  return applyColumnImport_(parsed.items);
}

function previewColumnsFromText(text) {
  const rows = parseDelimitedText_(text);
  return parseColumnRows_(rows);
}

function importColumnsFromText(text) {
  const parsed = previewColumnsFromText(text);
  return applyColumnImport_(parsed.items);
}

function readRowsFromSourceSheet_(url, tabName) {
  if (!url) throw new Error('Sheet URL is required');
  let ss;
  try {
    ss = SpreadsheetApp.openByUrl(String(url));
  } catch (e) {
    const id = extractSheetId_(url);
    try { ss = SpreadsheetApp.openById(id); }
    catch (ee) { throw new Error('Could not open spreadsheet. Check the URL and that you have view access.'); }
  }
  let sheet;
  if (tabName && String(tabName).trim()) {
    sheet = ss.getSheetByName(String(tabName).trim());
    if (!sheet) throw new Error('Tab not found: ' + tabName);
  } else {
    sheet = ss.getSheets()[0];
  }
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return [];
  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

function extractSheetId_(url) {
  const m = String(url || '').match(/\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : String(url || '').trim();
}

function parseDelimitedText_(text) {
  const raw = String(text || '').replace(/\r/g, '');
  if (!raw.trim()) return [];
  // Detect delimiter: tab if any tabs exist, else comma.
  const delim = raw.indexOf('\t') !== -1 ? '\t' : ',';
  const out = [];
  let cur = '';
  let row = [];
  let inQuote = false;
  for (let i = 0; i < raw.length; i++) {
    const c = raw[i];
    if (inQuote) {
      if (c === '"') {
        if (raw[i + 1] === '"') { cur += '"'; i++; }
        else inQuote = false;
      } else cur += c;
    } else {
      if (c === '"') inQuote = true;
      else if (c === delim) { row.push(cur); cur = ''; }
      else if (c === '\n') { row.push(cur); out.push(row); row = []; cur = ''; }
      else cur += c;
    }
  }
  if (cur.length || row.length) { row.push(cur); out.push(row); }
  return out.filter(r => r.some(c => String(c).trim() !== ''));
}

function parseColumnRows_(rows) {
  if (!rows || rows.length === 0) return { items: [], hadHeaders: false };
  const firstRow = rows[0].map(c => String(c == null ? '' : c).toLowerCase().trim());
  const looksLikeHeader = firstRow.some(c =>
    c.includes('name') || c.includes('type') || c.includes('option') || c === 'column'
  );
  let dataRows, idxName, idxDisplay, idxType, idxOptions;
  if (looksLikeHeader) {
    idxDisplay = firstRow.findIndex(c => c.includes('display'));
    idxName = firstRow.findIndex(c => (c.includes('name') && !c.includes('display')) || c === 'column');
    if (idxName === -1) idxName = 0;
    idxType = firstRow.findIndex(c => c.includes('type'));
    idxOptions = firstRow.findIndex(c => c.includes('option'));
    dataRows = rows.slice(1);
  } else {
    idxName = 0;
    idxDisplay = -1;
    idxType = rows[0].length > 1 ? 1 : -1;
    idxOptions = rows[0].length > 2 ? 2 : -1;
    dataRows = rows;
  }

  const items = [];
  dataRows.forEach(r => {
    const name = String(r[idxName] == null ? '' : r[idxName]).trim();
    if (!name) return;
    const display = idxDisplay >= 0 ? String(r[idxDisplay] == null ? '' : r[idxDisplay]).trim() : '';
    const rawType = idxType >= 0 ? String(r[idxType] == null ? '' : r[idxType]).toLowerCase().trim() : '';
    const type = rawType || 'text';
    const opts = idxOptions >= 0 ? String(r[idxOptions] == null ? '' : r[idxOptions]).trim() : '';
    let valid = true;
    let error = '';
    if (CASES_FIXED_HEADERS.indexOf(name) !== -1) {
      valid = false;
      error = 'Reserved column name';
    } else if (COLUMN_TYPES.indexOf(type) === -1) {
      valid = false;
      error = 'Invalid type "' + type + '" (allowed: ' + COLUMN_TYPES.join(', ') + ')';
    }
    items.push({
      ColumnName: name,
      DisplayName: display || name,
      ColumnType: type,
      Options: opts,
      valid: valid,
      error: error
    });
  });
  return { items: items, hadHeaders: looksLikeHeader };
}

function applyColumnImport_(items) {
  ensureSchema_();
  const existing = getCustomColumns_();
  const existingNames = {};
  existing.forEach(c => { existingNames[String(c.ColumnName).toLowerCase()] = true; });
  const result = { added: 0, skipped: 0, errors: [] };
  let order = existing.length;
  (items || []).forEach(item => {
    if (!item.valid) {
      result.errors.push((item.ColumnName || '(blank)') + ': ' + item.error);
      return;
    }
    if (existingNames[String(item.ColumnName).toLowerCase()]) {
      result.skipped++;
      return;
    }
    saveColumn({
      ColumnName: item.ColumnName,
      DisplayName: item.DisplayName,
      ColumnType: item.ColumnType,
      Options: item.Options,
      Order: order++
    });
    existingNames[String(item.ColumnName).toLowerCase()] = true;
    result.added++;
  });
  return result;
}
