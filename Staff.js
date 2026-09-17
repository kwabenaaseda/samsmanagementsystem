/**
 * Staff.js
 * Staff records over the Staff sheet.
 *
 * SCHEMA ASSUMPTION: the Staff tab must contain exactly these columns (order is
 * irrelevant, names must match): Staff_ID, First_Name, Last_Name, Gender,
 * Date_of_Birth, Phone, Email, Address, Position, Department, Employment_Date,
 * Employment_Status, Salary_Amount, Salary_Frequency, Last_Salary_Paid_Date,
 * Next_Salary_Due_Date, Salary_Status, Notes.
 *
 * CONVENTIONS
 *   - Staff_ID is server-generated (generateId_('STF')) and immutable. Clients
 *     never supply or change it.
 *   - Staff are never physically deleted. Deactivation is a soft status change
 *     (Employment_Status='Inactive') that preserves the row for history.
 *   - Every write runs inside withScriptLock_.
 *   - Authorization uses requirePermission_ (centralized Phase 2 mechanism).
 *   - A missing or malformed sheet is a structural failure (SERVER_ERROR).
 *
 * ASSUMPTION (documented business rule): required fields on create are
 * First_Name, Last_Name, Email, Position and Employment_Date. Employment_Status
 * defaults to 'Active'. Salary_Frequency is validated against a fixed allow-list
 * when supplied. Salary_Amount, when supplied, must be a non-negative number.
 * All other columns are optional.
 */

var STAFF_COLUMNS = [
  'Staff_ID', 'First_Name', 'Last_Name', 'Gender', 'Date_of_Birth',
  'Phone', 'Email', 'Address', 'Position', 'Department',
  'Employment_Date', 'Employment_Status', 'Salary_Amount', 'Salary_Frequency',
  'Last_Salary_Paid_Date', 'Next_Salary_Due_Date', 'Salary_Status', 'Notes'
];

var REQUIRED_STAFF_FIELDS = ['First_Name', 'Last_Name', 'Email', 'Position', 'Employment_Date'];
var STAFF_ID_COLUMN = 'Staff_ID';
var STAFF_STATUS_COLUMN = 'Employment_Status';

/* ===========================================================================
 * Structural access
 * ========================================================================= */

/**
 * Get the Staff sheet, translating a missing tab (NOT_FOUND) into a
 * structural SERVER_ERROR.
 * @return {Sheet}
 */
function getStaffSheet_() {
  var sheet;
  try {
    sheet = getSheet_(CONFIG.SHEETS.STAFF);
  } catch (err) {
    if (err && err.code === ERROR_CODES.NOT_FOUND) {
      throwError_('Staff sheet "' + CONFIG.SHEETS.STAFF + '" is missing from the spreadsheet.',
        ERROR_CODES.SERVER_ERROR, { sheet: CONFIG.SHEETS.STAFF, reason: 'sheet-missing' });
    }
    throw err;
  }
  return sheet;
}

/**
 * Validate the Staff sheet exists and contains its full canonical schema.
 * @return {{sheet: Sheet, headers: string[]}}
 * @throws {Error} SERVER_ERROR when the sheet or any required column is missing.
 */
function getStaffHeaders_() {
  var sheet = getStaffSheet_();
  var headers = getHeaders_(sheet);
  if (headers.length === 0) {
    throwError_('Staff sheet has no header row; cannot map records to columns.',
      ERROR_CODES.SERVER_ERROR, { sheet: CONFIG.SHEETS.STAFF });
  }
  var missing = STAFF_COLUMNS.filter(function (c) { return headers.indexOf(c) === -1; });
  if (missing.length > 0) {
    throwError_('Staff sheet is missing recommended column(s): ' + missing.join(', ') + '.',
      ERROR_CODES.SERVER_ERROR, { sheet: CONFIG.SHEETS.STAFF, missingColumns: missing });
  }
  return { sheet: sheet, headers: headers };
}

/**
 * Read every staff row as header-keyed objects, skipping blank rows.
 * @return {Object[]}
 */
function readAllStaff_() {
  var ctx = getStaffHeaders_();
  var sheet = ctx.sheet;
  var headers = ctx.headers;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(1, 1, lastRow, headers.length).getValues();
  var records = [];
  for (var i = 1; i < values.length; i++) {
    if (isBlankRow_(values[i])) continue;
    records.push(rowToObject_(headers, values[i]));
  }
  return records;
}

/**
 * Look up a staff member by Staff_ID.
 * @return {{sheetRow: number, record: Object}|null}
 */
function findStaffById_(id) {
  return findRowById_(CONFIG.SHEETS.STAFF, id, STAFF_ID_COLUMN);
}

/**
 * Validate (without changing) the mutable staff fields. Defaults are applied by
 * callers, so this never inserts a value.
 */
function normalizeStaffRecord_(record) {
  if (!isBlank_(record[STAFF_STATUS_COLUMN])) {
    assertOneOf_(toTrimmedString_(record[STAFF_STATUS_COLUMN]), CONFIG.VALUES.STAFF_STATUS, STAFF_STATUS_COLUMN);
    record[STAFF_STATUS_COLUMN] = toTrimmedString_(record[STAFF_STATUS_COLUMN]);
  }
  if (!isBlank_(record.Email)) {
    assertEmail_(record.Email, 'Email');
    record.Email = toTrimmedString_(record.Email);
  }
  if ('Salary_Frequency' in record && !isBlank_(toTrimmedString_(record.Salary_Frequency))) {
    assertOneOf_(toTrimmedString_(record.Salary_Frequency), CONFIG.VALUES.SALARY_FREQUENCY, 'Salary_Frequency');
    record.Salary_Frequency = toTrimmedString_(record.Salary_Frequency);
  }
  if ('Salary_Amount' in record && !isBlank_(record.Salary_Amount)) {
    if (!isNonNegativeNumber_(record.Salary_Amount)) {
      throwError_('Salary_Amount must be a non-negative number.',
        ERROR_CODES.VALIDATION_ERROR, { field: 'Salary_Amount', received: record.Salary_Amount });
    }
    record.Salary_Amount = toNumber_(record.Salary_Amount);
  }
}

/* ==========================================================================
 * Actions
 * ========================================================================= */

function handleStaffList_(payload) {
  requirePermission_('STAFF.READ');
  var records = readAllStaff_();
  return success({ staff: records, count: records.length },
    'Retrieved ' + records.length + ' staff member(s)');
}

function handleStaffGet_(payload) {
  requirePermission_('STAFF.READ');
  var id = toTrimmedString_(payload && payload.Staff_ID);
  if (id === '') {
    throwError_('Staff_ID is required to look up staff.',
      ERROR_CODES.VALIDATION_ERROR, { field: STAFF_ID_COLUMN });
  }
  getStaffHeaders_();
  var found = findStaffById_(id);
  if (!found) {
    throwError_('No staff found for Staff_ID "' + id + '".',
      ERROR_CODES.NOT_FOUND, { [STAFF_ID_COLUMN]: id });
  }
  return success(found.record, 'Staff retrieved');
}

function handleStaffCreate_(payload) {
  requirePermission_('STAFF.CREATE');
  getStaffHeaders_();
  if (!isBlank_(payload && payload.Staff_ID)) {
    throwError_('Staff_ID is server-generated and cannot be supplied by the client.',
      ERROR_CODES.VALIDATION_ERROR, { field: STAFF_ID_COLUMN });
  }
  assertRequired_(payload, REQUIRED_STAFF_FIELDS);
  var record = Object.assign({}, payload);
  record.Staff_ID = generateId_('STF');
  if (isBlank_(record[STAFF_STATUS_COLUMN])) record[STAFF_STATUS_COLUMN] = CONFIG.STAFF_STATUS.ACTIVE;
  normalizeStaffRecord_(record);
  if (findStaffById_(record.Staff_ID)) {
    throwError_('A staff member with Staff_ID "' + record.Staff_ID + '" already exists.',
      ERROR_CODES.CONFLICT, { [STAFF_ID_COLUMN]: record.Staff_ID });
  }
  var created = appendRow_(CONFIG.SHEETS.STAFF, record);
  return success(created, 'Staff created');
}

function handleStaffUpdate_(payload) {
  requirePermission_('STAFF.UPDATE');
  var id = toTrimmedString_(payload && payload.Staff_ID);
  if (id === '') {
    throwError_('Staff_ID is required to update staff.',
      ERROR_CODES.VALIDATION_ERROR, { field: STAFF_ID_COLUMN });
  }
  // Staff_ID identifies the row only; it can never be changed by an update.
  var patch = Object.assign({}, payload);
  delete patch[STAFF_ID_COLUMN];

  // The read-modify-write must be atomic: hold the lock across the whole
  // sequence so two concurrent updates cannot lose each other's fields.
  return withScriptLock_(function () {
    var ctx = getStaffHeaders_();
    var existing = findStaffById_(id);
    if (!existing) {
      throwError_('No staff found for Staff_ID "' + id + '".',
        ERROR_CODES.NOT_FOUND, { [STAFF_ID_COLUMN]: id });
    }
    var merged = Object.assign({}, existing.record, patch);
    normalizeStaffRecord_(merged);
    // recordToValues_ rejects unknown columns -> VALIDATION_ERROR and aligns order.
    var values = recordToValues_(CONFIG.SHEETS.STAFF, ctx.headers, merged);
    ctx.sheet.getRange(existing.sheetRow, 1, 1, ctx.headers.length).setValues([values]);
    return success(rowToObject_(ctx.headers, values), 'Staff updated');
  });
}

function handleStaffDeactivate_(payload) {
  requirePermission_('STAFF.DEACTIVATE');
  var id = toTrimmedString_(payload && payload.Staff_ID);
  if (id === '') {
    throwError_('Staff_ID is required to deactivate staff.',
      ERROR_CODES.VALIDATION_ERROR, { field: STAFF_ID_COLUMN });
  }
  // Soft status change: never delete the row.
  return withScriptLock_(function () {
    var ctx = getStaffHeaders_();
    var existing = findStaffById_(id);
    if (!existing) {
      throwError_('No staff found for Staff_ID "' + id + '".',
        ERROR_CODES.NOT_FOUND, { [STAFF_ID_COLUMN]: id });
    }
    var merged = Object.assign({}, existing.record, {
      [STAFF_STATUS_COLUMN]: CONFIG.STAFF_STATUS.INACTIVE
    });
    var values = recordToValues_(CONFIG.SHEETS.STAFF, ctx.headers, merged);
    ctx.sheet.getRange(existing.sheetRow, 1, 1, ctx.headers.length).setValues([values]);
    return success(rowToObject_(ctx.headers, values), 'Staff deactivated');
  });
}
