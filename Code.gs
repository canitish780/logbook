/**
 * Field Logbook — Google Sheets backend.
 *
 * SETUP:
 * 1. Create a new Google Sheet (any name).
 * 2. Extensions → Apps Script. Delete any starter code and paste this whole file in.
 * 3. Change SECRET below to a long random string — this is your app's write key,
 *    separate from the human passphrase you'll set up in the app itself.
 * 4. Deploy → New deployment → type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 *    Click Deploy, authorize it, and copy the "Web app URL" it gives you
 *    (ends in /exec).
 * 5. Paste that URL into the SHEETS_URL constant near the top of index.html
 *    and add.html (search for "PASTE_YOUR_APPS_SCRIPT_URL_HERE"), then
 *    re-upload those two files to your site.
 * 6. Open the site, go to Settings → Advanced setup, enter the same SECRET
 *    you set below as the "App key", pick a passphrase, and Save & test.
 *
 * The sheet will grow "Tasks", "Projects", and "Config" tabs automatically —
 * you don't need to create them yourself.
 */

const SECRET = 'CHANGE_ME_TO_A_LONG_RANDOM_STRING';

const TASK_COLUMNS = ['id','title','notes','date','place','priority','completed','createdAt','completedAt'];
const PROJECT_COLUMNS = ['id','title','notes','date','place','completed','createdAt','completedAt','links'];

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getSyncBlob') {
    return respond({ ok: true, blob: getConfigCell() });
  }
  if (action === 'getData') {
    if (e.parameter.key !== SECRET) return respond({ ok: false, error: 'Invalid key' });
    return respond({ ok: true, tasks: readSheet('Tasks'), projects: readSheet('Projects') });
  }
  return respond({ ok: false, error: 'Unknown action' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return respond({ ok: false, error: 'Bad request body' }); }

  if (body.key !== SECRET) return respond({ ok: false, error: 'Invalid key' });

  if (body.action === 'saveSyncBlob') {
    setConfigCell(body.blob || '');
    return respond({ ok: true });
  }

  if (body.action === 'saveAll') {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      writeSheet('Tasks', body.tasks || [], TASK_COLUMNS);
      writeSheet('Projects', body.projects || [], PROJECT_COLUMNS);
      return respond({ ok: true });
    } finally {
      lock.releaseLock();
    }
  }

  return respond({ ok: false, error: 'Unknown action' });
}

function readSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const cols = name === 'Tasks' ? TASK_COLUMNS : PROJECT_COLUMNS;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, cols.length).getValues();
  return values.filter(r => r[0]).map(r => {
    const obj = {};
    cols.forEach((c, i) => {
      let v = r[i];
      if (c === 'completed') v = (v === true || v === 'TRUE' || v === 'true');
      else if (c === 'links') { try { v = v ? JSON.parse(v) : []; } catch (e) { v = []; } }
      else if (v instanceof Date) {
        v = (c === 'date')
          ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : v.toISOString();
      } else if (v === '') {
        v = null;
      }
      obj[c] = v;
    });
    return obj;
  });
}

function writeSheet(name, rows, cols) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
  if (rows.length === 0) return;
  const values = rows.map(r => cols.map(c => {
    let v = r[c];
    if (c === 'links') v = JSON.stringify(v || []);
    if (v === null || v === undefined) v = '';
    return v;
  }));
  sheet.getRange(2, 1, values.length, cols.length).setValues(values);
}

function getConfigCell() {
  return getConfigSheet().getRange('A1').getValue() || '';
}
function setConfigCell(value) {
  getConfigSheet().getRange('A1').setValue(value);
}
function getConfigSheet() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Config');
  return sheet;
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
