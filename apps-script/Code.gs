const SPREADSHEET_ID = '1IHGSW_GHv67sa3WNqmGqAy0u2p19fIm0mnxG34W6rFU';
const SHEET_NAME = 'Shared State';
const CHUNK_SIZE = 40000;
const MAX_CHUNKS = 200;

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  const headers = ['state_json', 'revision', 'updated_at', 'updated_by', 'workspace_key'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sh;
}

function readState_(sh) {
  const values = sh.getRange(2, 1, MAX_CHUNKS, 1).getDisplayValues();
  let json = '';
  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) break;
    json += values[i][0];
  }
  if (!json) json = '{}';

  let state = {};
  try {
    state = JSON.parse(json);
  } catch (err) {
    throw new Error('Stored shared state is invalid JSON.');
  }

  const meta = sh.getRange(2, 2, 1, 4).getValues()[0];
  return {
    state,
    revision: Number(meta[0] || 0),
    updatedAt: meta[1] ? String(meta[1]) : '',
    updatedBy: meta[2] ? String(meta[2]) : '',
    workspaceKey: meta[3] ? String(meta[3]) : ''
  };
}

function writeState_(sh, state, revision, updatedBy, workspaceKey) {
  const json = JSON.stringify(state);
  const chunks = [];
  for (let i = 0; i < json.length; i += CHUNK_SIZE) {
    chunks.push([json.slice(i, i + CHUNK_SIZE)]);
  }
  if (!chunks.length) chunks.push(['{}']);
  if (chunks.length > MAX_CHUNKS) throw new Error('Shared state is too large.');

  sh.getRange(2, 1, MAX_CHUNKS, 1).clearContent();
  sh.getRange(2, 1, chunks.length, 1).setValues(chunks);
  sh.getRange(2, 2, 1, 4).setValues([[
    revision,
    new Date().toISOString(),
    updatedBy || 'Unknown',
    workspaceKey
  ]]);
}

function validateOrBindKey_(sh, suppliedKey) {
  if (!suppliedKey) throw new Error('Missing workspace key.');
  const cell = sh.getRange(2, 5);
  const existing = String(cell.getDisplayValue() || '');

  if (!existing) {
    cell.setValue(suppliedKey);
    return suppliedKey;
  }
  if (existing !== suppliedKey) throw new Error('Invalid workspace key.');
  return existing;
}

function response_(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const callback = e && e.parameter ? e.parameter.callback : '';
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = getSheet_();
      validateOrBindKey_(sh, e.parameter.key || '');
      const data = readState_(sh);
      const uninitialized = data.revision === 0 && (!data.state || Object.keys(data.state).length === 0);

      return response_({
        ok: true,
        uninitialized,
        state: data.state,
        revision: data.revision,
        updatedAt: data.updatedAt,
        updatedBy: data.updatedBy
      }, callback);
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return response_({ ok: false, error: String(err.message || err) }, callback);
  }
}

function doPost(e) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = getSheet_();
      const key = validateOrBindKey_(sh, e.parameter.key || '');
      const raw = e.parameter.state || '';
      if (!raw) throw new Error('Missing state payload.');

      let state;
      try {
        state = JSON.parse(raw);
      } catch (err) {
        throw new Error('State payload is invalid JSON.');
      }
      if (!state || typeof state !== 'object' || Array.isArray(state)) {
        throw new Error('State payload must be an object.');
      }

      const current = readState_(sh);
      const revision = Number(current.revision || 0) + 1;
      const updatedBy = String(e.parameter.updatedBy || 'Unknown').slice(0, 80);
      writeState_(sh, state, revision, updatedBy, key);

      return response_({ ok: true, revision }, '');
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return response_({ ok: false, error: String(err.message || err) }, '');
  }
}
