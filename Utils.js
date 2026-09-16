/**
 * Utils.js
 * Reusable backend helpers: spreadsheet access, sheet reads/writes, ID
 * generation, date handling, and validation.
 *
 * SCOPE: infrastructure only. No student / staff / finance / inventory
 * business rules belong in this file.
 *
 * CONVENTIONS
 *   - Trailing underscore marks an internal helper (not part of the API).
 *   - All spreadsheet access goes through getSpreadsheet_(), so
 *     SpreadsheetApp.openById() appears in exactly one place.
 *   - Every mutating operation runs inside withScriptLock_() so concurrent
 *     web-app requests cannot interleave and corrupt data.
 */

/* ==========================================================================
 * Errors
 * ======================================================================== */

/**
 * Build an Error carrying an application error code and optional details, so
 * the router can convert it into a standard failure envelope.
 *
 * @param {string} message Human-readable message.
 * @param {string} code One of ERROR_CODES (defaults to SERVER_ERROR).
 * @param {*} details Optional machine-readable context.
 * @return {Error} Error with .code and .details attached.
 */
function appError_(message, code = ERROR_CODES.SERVER_ERROR, details = null) {
  const error = new Error(message);
  error.code = code;
  error.details = details;
  return error;
}

/**
 * Throw a structured application error.
 * @param {string} message Human-readable message.
 * @param {string} code One of ERROR_CODES.
 * @param {*} details Optional machine-readable context.
 */
function throwError_(message, code, details) {
  throw appError_(message, code, details);
}

/* ==========================================================================
 * Spreadsheet / sheet access
 * ======================================================================== */

/**
 * Open the configured spreadsheet by explicit ID.
 *
 * openById() is used deliberately: this script is container-bound, but a web
 * app request (doGet/doPost) executes outside the container context, where
 * getActiveSpreadsheet() returns null.
 *
 * @return {Spreadsheet} The spreadsheet.
 * @throws {Error} SERVER_ERROR when the spreadsheet cannot be opened.
 */
function getSpreadsheet_() {
  try {
    return SpreadsheetApp.openById(CONFIG.SHEET_ID);
  } catch (err) {
    throwError_(
      'Could not open spreadsheet "' + CONFIG.SHEET_ID + '". Check CONFIG.SHEET_ID and that the executing user has access to it.',
      ERROR_CODES.SERVER_ERROR,
      { spreadsheetId: CONFIG.SHEET_ID, cause: err && err.message ? err.message : String(err) }
    );
  }
}

/**
 * @return {string[]} Every real tab name, in spreadsheet order.
 */
function listSheetNames_() {
  return getSpreadsheet_().getSheets().map(function (sheet) {
    return sheet.getName();
  });
}

/**
 * @param {string} name Sheet (tab) name.
 * @return {Sheet|null} The sheet, or null when it does not exist.
 */
function getSheetOrNull_(name) {
  if (isBlank_(name)) return null;
  return getSpreadsheet_().getSheetByName(toTrimmedString_(name));
}

/**
 * @param {string} name Sheet (tab) name.
 * @return {boolean} True when a tab with this exact name exists.
 */
function sheetExists_(name) {
  return getSheetOrNull_(name) !== null;
}

/**
 * Get a sheet by name or fail with a useful error listing what does exist.
 *
 * @param {string} name Sheet (tab) name.
 * @return {Sheet} The sheet.
 * @throws {Error} VALIDATION_ERROR when name is blank, NOT_FOUND when absent.
 */
function getSheet_(name) {
  if (isBlank_(name)) {
    throwError_('A sheet name is required.', ERROR_CODES.VALIDATION_ERROR, { received: name });
  }
  const sheet = getSheetOrNull_(name);
  if (!sheet) {
    throwError_(
      'Sheet "' + toTrimmedString_(name) + '" was not found in the spreadsheet.',
      ERROR_CODES.NOT_FOUND,
      { sheet: toTrimmedString_(name), availableSheets: listSheetNames_() }
    );
  }
  return sheet;
}

/* ==========================================================================
 * Reading
 * ======================================================================== */

/**
 * Read the header row of a sheet.
 *
 * Blank header cells are preserved as '' so that every header index still
 * lines up with its column position in the sheet.
 *
 * @param {Sheet} sheet The sheet.
 * @return {string[]} Trimmed header values (may contain empty strings).
 */
function getHeaders_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return [];
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function (header) {
    return toTrimmedString_(header);
  });
}

/**
 * Map a row array onto an object keyed by the sheet headers.
 * Columns with a blank header are skipped.
 *
 * @param {string[]} headers Header values aligned to column position.
 * @param {Array<*>} row The row values.
 * @return {Object} Header-keyed record.
 */
function rowToObject_(headers, row) {
  const record = {};
  for (let i = 0; i < headers.length; i++) {
    const key = headers[i];
    if (key === '') continue;
    record[key] = row[i] === undefined || row[i] === null ? '' : row[i];
  }
  return record;
}

/**
 * @param {Array<*>} row A row of cell values.
 * @return {boolean} True when every cell is blank.
 */
function isBlankRow_(row) {
  for (let i = 0; i < row.length; i++) {
    if (!isBlank_(row[i])) return false;
  }
  return true;
}

/**
 * Read every data row of a named sheet as an array of header-keyed objects.
 * Blank rows are skipped. Returns [] for an empty or header-only sheet.
 *
 * @param {string} name Sheet (tab) name.
 * @return {Object[]} Records.
 */
function readAll_(name) {
  const sheet = getSheet_(name);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = values[0].map(function (header) {
    return toTrimmedString_(header);
  });

  const records = [];
  for (let i = 1; i < values.length; i++) {
    if (isBlankRow_(values[i])) continue;
    records.push(rowToObject_(headers, values[i]));
  }
  return records;
}

/* ==========================================================================
 * Writing
 * ======================================================================== */

/**
 * Run `fn` while holding the document script lock.
 *
 * Google Sheets has no transactions. Without this, two staff members saving at
 * the same moment can interleave read-modify-write sequences and lose data.
 * This must wrap any read-then-write sequence.
 *
 * @param {Function} fn The work to perform under lock.
 * @return {*} Whatever fn returns.
 * @throws {Error} CONFLICT when the lock cannot be acquired in time.
 */
function withScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  const acquired = lock.tryLock(CONFIG.LOCK_TIMEOUT_MS);
  if (!acquired) {
    throwError_(
      'The system is busy handling another request. Please try again.',
      ERROR_CODES.CONFLICT,
      { lockTimeoutMs: CONFIG.LOCK_TIMEOUT_MS }
    );
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Convert a header-keyed record into values aligned to sheet column order.
 *
 * @param {string} sheetName Sheet name, for error messages.
 * @param {string[]} headers The sheet headers.
 * @param {Object} record Header-keyed values.
 * @return {Array<*>} Values in column order.
 * @throws {Error} VALIDATION_ERROR when the record contains unknown columns.
 */
function recordToValues_(sheetName, headers, record) {
  const values = [];
  const unknown = [];

  Object.keys(record).forEach(function (key) {
    const index = headers.indexOf(key);
    if (index === -1) {
      unknown.push(key);
      return;
    }
    values[index] = record[key];
  });

  if (unknown.length) {
    throwError_(
      'Unknown column(s) for sheet "' + sheetName + '": ' + unknown.join(', ') + '.',
      ERROR_CODES.VALIDATION_ERROR,
      { sheet: sheetName, unknownColumns: unknown, validColumns: headers }
    );
  }

  for (let i = 0; i < headers.length; i++) {
    if (values[i] === undefined) values[i] = '';
  }
  return values;
}

/**
 * Append one row to a named sheet.
 *
 * `row` may be an object keyed by header (preferred) or an array already in
 * sheet column order. Objects are matched to headers by name and re-ordered,
 * so changing the order of keys in code can never write to the wrong column.
 *
 * @param {string} name Sheet (tab) name.
 * @param {Object|Array<*>} row The row to append.
 * @return {Object} The appended record, keyed by header.
 */
function appendRow_(name, row) {
  const sheet = getSheet_(name);
  const headers = getHeaders_(sheet);

  if (headers.length === 0) {
    throwError_(
      'Sheet "' + name + '" has no header row, so records cannot be mapped to columns.',
      ERROR_CODES.SERVER_ERROR,
      { sheet: name }
    );
  }

  return withScriptLock_(function () {
    let values;

    if (Array.isArray(row)) {
      values = row.slice(0, headers.length);
    } else if (row !== null && typeof row === 'object') {
      values = recordToValues_(name, headers, row);
    } else {
      throwError_(
        'appendRow_ expects an object keyed by header or an array, but received ' + typeof row + '.',
        ERROR_CODES.VALIDATION_ERROR,
        { sheet: name, receivedType: typeof row }
      );
    }

    const padded = [];
    for (let i = 0; i < headers.length; i++) {
      padded.push(values[i] === undefined ? '' : values[i]);
    }

    const targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1, 1, headers.length).setValues([padded]);

    return rowToObject_(headers, padded);
  });
}

/**
 * Find the first row whose id column equals `id`.
 *
 * @param {string} name Sheet (tab) name.
 * @param {string} id The ID to search for.
 * @param {string=} idColumn Header to search in. Defaults to the first header.
 * @return {{sheetRow: number, record: Object}|null} sheetRow is 1-based so it
 *     can be passed directly to getRange(). Null when nothing matches.
 */
function findRowById_(name, id, idColumn) {
  const sheet = getSheet_(name);
  const headers = getHeaders_(sheet);

  if (headers.length === 0) {
    throwError_(
      'Sheet "' + name + '" has no header row.',
      ERROR_CODES.SERVER_ERROR,
      { sheet: name }
    );
  }

  const column = isBlank_(idColumn) ? headers[0] : toTrimmedString_(idColumn);
  const columnIndex = headers.indexOf(column);
  if (columnIndex === -1) {
    throwError_(
      'Column "' + column + '" was not found in sheet "' + name + '".',
      ERROR_CODES.NOT_FOUND,
      { sheet: name, column: column, validColumns: headers }
    );
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const search = toTrimmedString_(id);
  const ids = sheet.getRange(2, columnIndex + 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (toTrimmedString_(ids[i][0]) === search) {
      const sheetRow = i + 2;
      const rowValues = sheet.getRange(sheetRow, 1, 1, headers.length).getValues()[0];
      return { sheetRow: sheetRow, record: rowToObject_(headers, rowValues) };
    }
  }
  return null;
}

/* ==========================================================================
 * ID generation
 * ======================================================================== */

/**
 * Generate a unique, readable ID such as "STU-20260916-4K7QX2A9B3C1".
 *
 * WHY THIS SHAPE: uniqueness must not depend on reading the sheet. A
 * "last row number + 1" scheme has a read-modify-write race, so two staff
 * saving at the same moment would generate the same ID. This scheme needs no
 * sheet access, so it is safe under concurrency.
 *
 * Collision space is 16^12 (~2.8e14) per prefix per day. Across a million IDs
 * generated over the whole life of the system the chance of any collision is
 * roughly 0.2%, so it is collision-RESISTANT rather than mathematically
 * guaranteed. A caller that needs a hard guarantee can confirm uniqueness
 * cheaply with findRowById_() before appending.
 *
 * @param {string} prefix Short prefix, e.g. 'STU'. Non-alphanumerics are
 *     stripped and the result is upper-cased.
 * @return {string} e.g. 'STU-20260916-4K7QX2A9B3C1'.
 */
function generateId_(prefix) {
  const clean = toTrimmedString_(prefix).toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (clean === '') {
    throwError_(
      'generateId_ requires a non-empty prefix (e.g. "STU").',
      ERROR_CODES.VALIDATION_ERROR,
      { received: prefix }
    );
  }

  const stamp = formatDate_(new Date()).replace(/-/g, '');
  const random = Utilities.getUuid().replace(/-/g, '').slice(0, 12).toUpperCase();
  return clean + '-' + stamp + '-' + random;
}

/* ==========================================================================
 * Dates and time
 * ======================================================================== */

/**
 * @return {string} The script's timezone (from appsscript.json).
 */
function getTimeZone_() {
  return Session.getScriptTimeZone();
}

/**
 * @return {Date} The current date/time.
 */
function now_() {
  return new Date();
}

/**
 * @return {string} UTC ISO-8601 timestamp, for machine-readable logging.
 *     Note this is UTC; use formatDateTime_() for display in school-local time.
 */
function nowIso_() {
  return new Date().toISOString();
}

/**
 * Coerce a value into a Date.
 *
 * @param {*} value Date, number (epoch millis), or a date string.
 * @return {Date|null} The Date, or null when the value is not a valid date.
 */
function toDate_(value) {
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }
  if (isBlank_(value)) return null;

  if (typeof value === 'number') {
    const fromMillis = new Date(value);
    return isNaN(fromMillis.getTime()) ? null : fromMillis;
  }

  const parsed = new Date(String(value).trim());
  return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Format a value as 'yyyy-MM-dd' in the script timezone.
 * @param {*} value Date / number / date string.
 * @return {string} Formatted date, or '' when the value is not a valid date.
 */
function formatDate_(value) {
  const date = toDate_(value);
  if (!date) return '';
  return Utilities.formatDate(date, getTimeZone_(), 'yyyy-MM-dd');
}

/**
 * Format a value as 'yyyy-MM-dd HH:mm:ss' in the script timezone.
 * @param {*} value Date / number / date string.
 * @return {string} Formatted timestamp, or '' when invalid.
 */
function formatDateTime_(value) {
  const date = toDate_(value);
  if (!date) return '';
  return Utilities.formatDate(date, getTimeZone_(), 'yyyy-MM-dd HH:mm:ss');
}

/* ==========================================================================
 * Validation helpers (pure -- no sheet access)
 * ======================================================================== */

/**
 * @param {*} value Any value.
 * @return {boolean} True for null, undefined, or a whitespace-only string.
 */
function isBlank_(value) {
  if (value === null || value === undefined) return true;
  if (typeof value === 'string') return value.trim() === '';
  return false;
}

/**
 * @param {*} value Any value.
 * @return {string} A trimmed string, or '' for null/undefined.
 */
function toTrimmedString_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

/**
 * @param {*} value Any value.
 * @return {boolean} True when the value is a non-empty string.
 */
function isNonEmptyString_(value) {
  return typeof value === 'string' && value.trim() !== '';
}

/**
 * Basic email shape check. Deliberately not a full RFC 5322 implementation --
 * it rejects obvious mistakes without rejecting valid addresses.
 *
 * @param {*} value Any value.
 * @return {boolean} True when the value looks like an email address.
 */
function isValidEmail_(value) {
  return /^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/.test(toTrimmedString_(value));
}

/**
 * Coerce a value to a finite number.
 * @param {*} value Number or numeric string.
 * @return {number|null} The number, or null when it is not numeric.
 */
function toNumber_(value) {
  if (isBlank_(value)) return null;
  const num = typeof value === 'number' ? value : Number(String(value).trim());
  if (isNaN(num) || !isFinite(num)) return null;
  return num;
}

/**
 * @param {*} value Any value.
 * @return {boolean} True when the value is numeric and >= 0.
 */
function isNonNegativeNumber_(value) {
  const num = toNumber_(value);
  return num !== null && num >= 0;
}

/**
 * @param {*} value Any value.
 * @return {boolean} True when the value is numeric and > 0.
 */
function isPositiveNumber_(value) {
  const num = toNumber_(value);
  return num !== null && num > 0;
}

/**
 * @param {*} value Any value.
 * @param {Array<*>} allowed Allowed values.
 * @return {boolean} True when the value is exactly one of `allowed`.
 */
function isOneOf_(value, allowed) {
  return Array.isArray(allowed) && allowed.indexOf(value) !== -1;
}

/* ==========================================================================
 * Assertions (throw structured errors for use by API actions)
 * ======================================================================== */

/**
 * Assert that a record is an object and that every named field is present.
 *
 * All missing fields are reported at once so the caller (and the React form)
 * can show every problem in a single round trip.
 *
 * @param {Object} record The record to check.
 * @param {string[]} fields Required field names.
 * @return {Object} The same record, for chaining.
 * @throws {Error} VALIDATION_ERROR listing the missing fields.
 */
function assertRequired_(record, fields) {
  if (record === null || typeof record !== 'object' || Array.isArray(record)) {
    throwError_(
      'Expected an object but received ' + (Array.isArray(record) ? 'array' : typeof record) + '.',
      ERROR_CODES.VALIDATION_ERROR,
      { receivedType: Array.isArray(record) ? 'array' : typeof record }
    );
  }

  const missing = [];
  for (let i = 0; i < fields.length; i++) {
    if (isBlank_(record[fields[i]])) missing.push(fields[i]);
  }

  if (missing.length) {
    throwError_(
      'Missing required field(s): ' + missing.join(', ') + '.',
      ERROR_CODES.VALIDATION_ERROR,
      { missingFields: missing }
    );
  }
  return record;
}

/**
 * Assert that a value is one of an allowed set (enum check).
 *
 * Intended for the enums in CONFIG, e.g.
 *   assertOneOf_(record.Status, CONFIG.VALUES.STUDENT_STATUS, 'Status');
 *
 * @param {*} value The value to check.
 * @param {Array<*>} allowed Allowed values.
 * @param {string=} label Field name used in the error message.
 * @return {*} The same value, for chaining.
 * @throws {Error} VALIDATION_ERROR when the value is not allowed.
 */
function assertOneOf_(value, allowed, label) {
  if (!Array.isArray(allowed)) {
    throwError_(
      'assertOneOf_ needs an array of allowed values.',
      ERROR_CODES.SERVER_ERROR,
      { label: label || null }
    );
  }
  if (!isOneOf_(value, allowed)) {
    throwError_(
      (label || 'Value') + ' must be one of: ' + allowed.join(', ') + '.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: label || null, value: value, allowedValues: allowed }
    );
  }
  return value;
}

/**
 * Assert that a value is a usable email address.
 *
 * NOTE: this requires a non-blank value. For an OPTIONAL email field, check
 * isBlank_(value) first and only assert when a value was supplied.
 *
 * @param {*} value The value to check.
 * @param {string=} label Field name used in the error message.
 * @return {string} The trimmed email address.
 * @throws {Error} VALIDATION_ERROR when the value is blank or malformed.
 */
function assertEmail_(value, label) {
  const field = label || 'Email';
  if (isBlank_(value)) {
    throwError_(field + ' is required.', ERROR_CODES.VALIDATION_ERROR, { field: label || null });
  }
  if (!isValidEmail_(value)) {
    throwError_(
      field + ' is not a valid email address.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: label || null, value: toTrimmedString_(value) }
    );
  }
  return toTrimmedString_(value);
}