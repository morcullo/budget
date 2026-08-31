// Budget app -> Google Sheets backend
//
// Web app setup:
// 1. Create/open a Google Sheet.
// 2. Extensions -> Apps Script.
// 3. Replace the default code with this file.
// 4. Deploy -> New deployment -> Web app.
//    Execute as: Me
//    Who has access: Anyone
// 5. Copy the /exec URL into google-sheets-config.js.
// 6. Optionally set the same secret in SECRET and GOOGLE_SHEETS_TOKEN.
//
// iOS Shortcut POST format:
// {
//   "action": "addTransaction",
//   "amount": 12.50,
//   "category": "Food",
//   "description": "Coffee",
//   "date": "2026-08-31"
// }
// A positive amount is treated as spending and stored as -12.50,
// matching the Budget app's transaction convention. A negative amount
// is preserved as income/credit.

const SECRET = '';
const SHEET_NAME = 'BudgetData';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 3).setValues([['key', 'value', 'updatedAt']]);
  }
  return sheet;
}

function authorized_(token) {
  return !SECRET || token === SECRET;
}

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'load';
  const token = (e.parameter && e.parameter.token) || '';
  if (!authorized_(token)) return json_({error: 'Unauthorized'});
  if (action !== 'load') return json_({error: 'Unknown action'});

  try {
    return json_({state: readState_()});
  } catch (err) {
    return json_({error: String(err)});
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (!authorized_(payload.token || '')) return json_({error: 'Unauthorized'});

    if (payload.action === 'save') {
      if (!payload.state) return json_({error: 'Invalid request'});
      saveState_(payload.state);
      return json_({ok: true});
    }

    if (payload.action === 'addTransaction') {
      const result = addTransaction_(payload);
      return json_({ok: true, transaction: result.transaction, state: result.state});
    }

    return json_({error: 'Unknown action'});
  } catch (err) {
    return json_({error: String(err)});
  }
}

function readState_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === 'state') {
      try { return JSON.parse(values[i][1]); }
      catch (err) { throw new Error('Stored state is invalid'); }
    }
  }
  return null;
}

function saveState_(state) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const serialized = JSON.stringify(state);
    const values = sheet.getDataRange().getValues();
    let row = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === 'state') { row = i + 1; break; }
    }
    const data = ['state', serialized, new Date()];
    if (row === -1) sheet.appendRow(data);
    else sheet.getRange(row, 1, 1, 3).setValues([data]);
  } finally {
    lock.releaseLock();
  }
}

function addTransaction_(payload) {
  const amountInput = Number(payload.amount);
  const category = String(payload.category || '').trim();
  const description = String(payload.description || '').trim();
  const date = normalizeDate_(payload.date);

  if (!isFinite(amountInput) || amountInput === 0) throw new Error('amount must be a non-zero number');
  if (!category) throw new Error('category is required');
  if (!description) throw new Error('description is required');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('date must use YYYY-MM-DD');
  if (!isValidDate_(date)) throw new Error('date is invalid');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let state = readState_();
    if (!state) throw new Error('No Budget state exists yet. Open the Budget app once and save it first.');

    const transaction = {
      id: 'shortcut_' + new Date().getTime() + '_' + Math.random().toString(36).slice(2, 7),
      date: date,
      desc: description,
      amount: amountInput > 0 ? -Math.abs(amountInput) : amountInput,
      cat: category
    };

    state.transactions = Array.isArray(state.transactions) ? state.transactions : [];
    state.archive = Array.isArray(state.archive) ? state.archive : [];

    // Keep transactions from the current month in transactions and older ones
    // in archive, matching the Budget app's existing data model.
    const currentYm = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
    if (date.substring(0, 7) === currentYm) {
      state.transactions.unshift(transaction);
    } else {
      transaction.archivedAt = new Date().toISOString();
      state.archive.unshift(transaction);
    }

    saveStateUnlocked_(state);
    return {transaction: transaction, state: state};
  } finally {
    lock.releaseLock();
  }
}

// Called while the script lock is already held.
function saveStateUnlocked_(state) {
  const sheet = getSheet_();
  const serialized = JSON.stringify(state);
  const values = sheet.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === 'state') { row = i + 1; break; }
  }
  const data = ['state', serialized, new Date()];
  if (row === -1) sheet.appendRow(data);
  else sheet.getRange(row, 1, 1, 3).setValues([data]);
}

function normalizeDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const raw = String(value || '').trim();
  if (!raw) return '';
  // Accept YYYY-MM-DD directly, or ISO date/time strings from Shortcuts.
  const m = raw.match(/^(\d{4}-\d{2}-\d{2})(?:$|T)/);
  if (m) return m[1];
  const parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return raw;
}

function isValidDate_(value) {
  const parts = value.split('-').map(Number);
  if (parts.length !== 3) return false;
  const d = new Date(parts[0], parts[1] - 1, parts[2]);
  return d.getFullYear() === parts[0] && d.getMonth() === parts[1] - 1 && d.getDate() === parts[2];
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
