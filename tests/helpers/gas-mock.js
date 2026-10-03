'use strict';
// ===============================================================
// gas-mock.js
// A dependency-free in-memory implementation of the Google Apps Script
// services used by Code.gs. Enough surface to unit test business logic
// (permissions, salary advances, attendance) without network or Sheets.
// ===============================================================

const crypto = require('node:crypto');

class GasRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }

  _clone(v) {
    return v instanceof Date ? new Date(v.getTime()) : v;
  }

  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const row = [];
      for (let c = 0; c < this.numCols; c++) {
        row.push(this._clone(this.sheet._cell(this.row + r, this.col + c)));
      }
      out.push(row);
    }
    return out;
  }

  getDisplayValues() {
    return this.getValues().map(r => r.map(v => (v === null || v === undefined ? '' : String(v))));
  }

  getValue() {
    return this.sheet._cell(this.row, this.col);
  }

  setValue(v) {
    this.sheet._set(this.row, this.col, v);
    return this;
  }

  setValues(values) {
    for (let r = 0; r < values.length; r++) {
      const row = values[r] || [];
      for (let c = 0; c < row.length; c++) {
        this.sheet._set(this.row + r, this.col + c, row[c]);
      }
    }
    return this;
  }

  getLastRow() {
    return this.sheet.getLastRow();
  }

  clear() {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) this.sheet._set(this.row + r, this.col + c, '');
    }
    return this;
  }

  clearContent() {
    return this.clear();
  }

  setNumberFormat(fmt) {
    this.sheet._numberFormat = fmt;
    return this;
  }

  getNumberFormat() {
    return this.sheet._numberFormat || 'General';
  }

  setFont() { return this; }
  setBackground() { return this; }
  setHorizontalAlignment() { return this; }
  setVerticalAlignment() { return this; }
  setWrap() { return this; }
  setBorder() { return this; }
  setFontColor() { return this; }
  setFontWeight() { return this; }
  setFontSize() { return this; }
  merge() { return this; }
  setFormula() { return this; }
  clearFormat() { return this; }
  sort() { return this; }
  setNotes() { return this; }
}

class GasSheet {
  constructor(spreadsheet, name, data) {
    this.spreadsheet = spreadsheet;
    this._name = name;
    this._data = Array.isArray(data) ? data.map(r => (Array.isArray(r) ? r.slice() : [r])) : [];
    this._frozenRows = 0;
    this._numberFormat = 'General';
    this._columnWidths = {};
  }

  getName() { return this._name; }
  getSheetId() { return this.spreadsheet._nextSheetId++; }
  setFrozenRows(n) { this._frozenRows = n; return this; }
  getFrozenRows() { return this._frozenRows; }
  setColumnWidth() { return this; }
  autoResizeColumns() { return this; }
  setTabColor() { return this; }
  hideColumn() { return this; }
  showColumn() { return this; }
  moveColumn() { return this; }
  setName(n) { this._name = n; return this; }

  _ensure(r, c) {
    while (this._data.length < r) this._data.push([]);
    const row = this._data[r - 1];
    while (row.length < c) row.push('');
  }

  _cell(r, c) {
    if (r < 1 || c < 1) return '';
    const row = this._data[r - 1];
    if (!row) return '';
    const v = row[c - 1];
    return v === undefined ? '' : v;
  }

  _set(r, c, v) {
    this._ensure(r, c);
    this._data[r - 1][c - 1] = v;
  }

  getLastRow() {
    for (let i = this._data.length - 1; i >= 0; i--) {
      const row = this._data[i] || [];
      for (let j = row.length - 1; j >= 0; j--) {
        const v = row[j];
        if (v !== '' && v !== null && v !== undefined) return i + 1;
      }
    }
    return 0;
  }

  getLastColumn() {
    let max = 0;
    for (const row of this._data) {
      for (let j = row.length - 1; j >= 0; j--) {
        const v = row[j];
        if (v !== '' && v !== null && v !== undefined) { max = Math.max(max, j + 1); break; }
      }
    }
    return max;
  }

  getMaxRows() { return Math.max(this._data.length, 1); }
  getMaxColumns() { return Math.max(this.getLastColumn(), 1); }
  getDataRange() { return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }

  getRange(row, col, numRows, numCols) {
    if (numRows === undefined) return new GasRange(this, row, col, 1, 1);
    if (numCols === undefined) return new GasRange(this, row, col, numRows, 1);
    return new GasRange(this, row, col, numRows, numCols);
  }

  getRangeByA1() { return this.getDataRange(); }

  appendRow(values) {
    this._data.push((values || []).slice());
    return this;
  }

  insertRowAfter() { this._data.push([]); return this; }
  insertRow() { this._data.push([]); return this; }
  deleteRow(index) { this._data.splice(index - 1, 1); return this; }
  deleteRows(start, count) { this._data.splice(start - 1, count); return this; }
  clear() { this._data = []; return this; }
  sort() { return this; }
  setColumnWidths() { return this; }
  getRangeList() { return { clear: () => {} }; }
  duplicate() { return new GasSheet(this.spreadsheet, this._name + ' copy', this._data); }
}

class GasSpreadsheet {
  constructor(data) {
    this._sheets = new Map();
    this._nextSheetId = 1;
    for (const [name, rows] of Object.entries(data || {})) {
      this._sheets.set(name, new GasSheet(this, name, rows));
    }
  }

  getSheetByName(name) { return this._sheets.get(String(name)) || null; }

  getSheets() { return Array.from(this._sheets.values()); }

  insertSheet(name) {
    if (this._sheets.has(name)) throw new Error('Sheet already exists: ' + name);
    const s = new GasSheet(this, name, []);
    this._sheets.set(name, s);
    return s;
  }

  deleteSheet(name) { return this._sheets.delete(String(name)); }
  getName() { return 'MockSpreadsheet'; }
  getId() { return 'mock-spreadsheet-id'; }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/mock'; }
  setSpreadsheetTimeZone() {}
  getSpreadsheetTimeZone() { return 'Asia/Amman'; }
}

class GasCache {
  constructor() { this._map = new Map(); }
  get(key) { return this._map.has(key) ? this._map.get(key) : null; }
  put(key, value, seconds) {
    this._map.set(key, value);
    if (seconds) this._map.set(key + ':__exp', Date.now() + seconds * 1000);
    return this;
  }
  remove(key) { this._map.delete(key); return this; }
  removeAll(keys) { (keys || []).forEach(k => this._map.delete(k)); return this; }
  _reset() { this._map.clear(); }
}

class GasProperties {
  constructor() { this._map = new Map(); }
  getProperty(k) { return this._map.has(k) ? this._map.get(k) : null; }
  setProperty(k, v) { this._map.set(k, String(v)); return this; }
  deleteProperty(k) { this._map.delete(k); return this; }
  getProperties() { return Object.fromEntries(this._map); }
}

class GasFile {
  constructor(id, name) { this.id = id; this._name = name; this._sharing = null; }
  getName() { return this._name; }
  setName(n) { this._name = n; return this; }
  getId() { return this.id; }
  getUrl() { return 'https://drive.google.com/file/d/' + this.id + '/view'; }
  setSharing(perm, role) { this._sharing = { perm, role }; return this; }
  getSharing() { return this._sharing; }
  makeCopy() { return new GasFile(this.id + '-copy', this._name + ' copy'); }
  getAs() { return { setMimeType: () => {} }; }
}

// The shared environment instance. Tests call createGasEnv() to get a fresh one.
function createGasEnv(initialData) {
  const spreadsheet = new GasSpreadsheet(initialData || {});
  const scriptCache = new GasCache();
  const docCache = new GasCache();
  const props = new GasProperties();
  const files = new Map();
  let triggerId = 0;
  let triggerHandler = null;

  const env = {
    // ---- state handles used by tests ----
    _spreadsheet: spreadsheet,
    _scriptCache: scriptCache,
    _driveFiles: files,
    _scriptTriggers: [],
    _setTriggerHandler(fn) { triggerHandler = fn; },
    _getTriggerHandler() { return triggerHandler; },

    reset: () => { scriptCache._reset(); docCache._reset(); },

    // ---- SpreadsheetApp ----
    SpreadsheetApp: {
      getActiveSpreadsheet: () => spreadsheet,
      getActive: () => spreadsheet,
      getUi: () => ({
        alert: () => {},
        showSidebar: () => {},
        createSidebar: () => ({ setTitle: () => ({ setContentHtml: () => ({}) }) }),
      }),
      flush: () => {},
      newSpreadsheet: () => new GasSpreadsheet({}),
    },

    // ---- CacheService ----
    CacheService: {
      getScriptCache: () => scriptCache,
      getDocumentCache: () => docCache,
    },

    // ---- PropertiesService ----
    PropertiesService: {
      getScriptProperties: () => props,
      getDocumentProperties: () => props,
      getUserProperties: () => props,
    },

    // ---- Utilities ----
    Utilities: {
      base64Encode: (input) => Buffer.from(String(input), 'utf8').toString('base64'),
      base64Decode: (input) => Buffer.from(String(input), 'base64').toString('utf8'),
      base64EncodeWebSafe: (input) => Buffer.from(String(input), 'utf8').toString('base64url'),
      base64DecodeWebSafe: (input) => Buffer.from(String(input), 'base64url').toString('utf8'),
      getUuid: () => 'uuid-' + crypto.randomUUID(),
      // Apps Script returns a SIGNED byte array, not base64.
      computeHmacSha256Signature: (value, key) => {
        const digest = crypto.createHmac('sha256', String(key)).update(String(value)).digest();
        return Array.from(digest, (b) => (b > 127 ? b - 256 : b));
      },
      computeDigest: (algo, value) => {
        const digest = crypto.createHash(String(algo)).update(String(value)).digest();
        return Array.from(digest, (b) => (b > 127 ? b - 256 : b));
      },
      formatDate: (date, tz, fmt) => {
        const d = date instanceof Date ? date : new Date(date);
        return fmt.replace('yyyy', d.getFullYear())
          .replace('MM', String(d.getMonth() + 1).padStart(2, '0'))
          .replace('dd', String(d.getDate()).padStart(2, '0'));
      },
      newBlob: (content, type, name) => ({
        getContentType: () => type || 'text/plain',
        getName: () => name || 'blob',
        setName: (n) => ({ getContentType: () => type, getName: () => n }),
        getBytes: () => (Array.isArray(content) ? content : Array.from(Buffer.from(String(content), 'utf8'))),
        getAs: (mime) => ({ getContentType: () => mime }),
      }),
      sleep: (ms) => { throw new Error('Utilities.sleep not supported in tests: ' + ms); },
      getScriptTimeZone: () => 'Asia/Amman',
    },

    // ---- DriveApp ----
    DriveApp: {
      Access: { PRIVATE: 'PRIVATE', ANYONE: 'ANYONE_WITH_LINK', DOMAIN: 'DOMAIN' },
      Permission: { VIEW: 'VIEW', EDIT: 'EDIT', NONE: 'NONE' },
      getFolderByName: (name) => {
        if (!files.has('folder:' + name)) {
          files.set('folder:' + name, env.DriveApp.createFolder(name));
        }
        return files.get('folder:' + name);
      },
      getFolderById: (id) => files.get(id),
      getFileById: (id) => files.get(id) || null,
      createFolder: (name) => {
        const f = {
          id: 'folder-' + name, _name: name, isFolder: true,
          getId: function () { return this.id; },
          getName: function () { return this._name; },
          setName(n) { this._name = n; return this; },
          createFile: (blob) => {
            const file = new GasFile('file-' + (files.size + 1), (blob && blob.getName && blob.getName()) || 'file');
            file._blob = blob;
            files.set(file.id, file);
            return file;
          },
        };
        files.set(f.id, f);
        return f;
      },
      createFile: (blob) => {
        const f = new GasFile('file-' + (files.size + 1), 'file');
        f._blob = blob;
        files.set(f.id, f);
        return f;
      },
      getFilesByName: (name) => Array.from(files.values()).filter(f => f._name === name),
    },

    // ---- Session ----
    Session: {
      getActiveUser: () => ({ getEmail: () => 'tester@example.com', getUsername: () => 'tester' }),
      getEffectiveUser: () => ({ getEmail: () => 'tester@example.com' }),
      getScriptTimeZone: () => 'Asia/Amman',
    },

    // ---- ScriptApp ----
    ScriptApp: {
      getProjectTriggers: () => env._scriptTriggers,
      newTrigger: (fn) => {
        const t = { uid: 'trigger-' + (++triggerId), handler: fn };
        env._scriptTriggers.push(t);
        return t;
      },
    },

    // ---- ContentService ----
    ContentService: {
      MimeType: { JSON: 'application/json', TEXT: 'text/plain', HTML: 'text/html', CSV: 'text/csv' },
      createTextOutput: (text) => {
        const out = {
          _text: text,
          _type: 'text/plain',
          setMimeType(t) { out._type = t; return out; },
          setContentType(t) { out._type = t; return out; },
          append() { return out; },
          getContent: () => text,
        };
        return out;
      },
    },

    // ---- Date / globals helpers ----
    __mock: true,
  };

  return env;
}

module.exports = { createGasEnv, GasSheet, GasSpreadsheet, GasRange, GasCache };
