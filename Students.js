/**
 * Students.js
 * Student records over the Students sheet.
 *
 * SCHEMA ASSUMPTION: the Students tab must contain exactly these columns (order
 * is irrelevant to the code, but the names must match): Student_ID, First_Name,
 * Last_Name, Gender, Date_of_Birth, Class, Parent_Guardian, Guardian_Phone,
 * Guardian_Email, Emergency_Contact_1_Name, Emergency_Contact_1_Phone,
 * Emergency_Contact_1_Relationship, Emergency_Contact_2_Name,
 * Emergency_Contact_2_Phone, Emergency_Contact_2_Relationship, Allergies,
 * Illnesses_Medical_Conditions, Physical_Defects_Special_Conditions,
 * Admission_Date, Status, Withdrawal_Date, Notes.
 *
 * CONVENTIONS
 *   - Student_ID is server-generated (generateId_('STU')) and immutable. Clients
 *     never supply or change it.
 *   - Students are never physically deleted. Withdrawal is a soft status change
 *     (Status='Withdrawn', Withdrawal_Date set) that preserves the row for
 *     historical records.
 *   - Every write runs inside withScriptLock_ (appendRow_ / writeStudentRow_).
 *   - Authorization uses the centralized Phase 2 mechanism (requirePermission_),
 *     so UNAUTHORIZED / FORBIDDEN are produced uniformly by the router.
 *   - A missing or malformed sheet is a structural failure (SERVER_ERROR), not a
 *     user error -- mirroring Auth.js.
 *
 * ASSUMPTION (documented business rule): the required fields on create are
 * First_Name, Last_Name, Class, Admission_Date and Parent_Guardian. All other
 * columns are optional; omitted optional fields are left blank.
 */

var STUDENT_COLUMNS = [
  'Student_ID', 'First_Name', 'Last_Name', 'Gender', 'Date_of_Birth',
  'Class', 'Parent_Guardian', 'Guardian_Phone', 'Guardian_Email',
  'Emergency_Contact_1_Name', 'Emergency_Contact_1_Phone',
  'Emergency_Contact_1_Relationship',
  'Emergency_Contact_2_Name', 'Emergency_Contact_2_Phone',
  'Emergency_Contact_2_Relationship',
  'Allergies', 'Illnesses_Medical_Conditions',
  'Physical_Defects_Special_Conditions',
  'Admission_Date', 'Status', 'Withdrawal_Date', 'Notes'
];

var REQUIRED_STUDENT_FIELDS = ['First_Name', 'Last_Name', 'Class', 'Admission_Date', 'Parent_Guardian'];
var STUDENT_ID_COLUMN = 'Student_ID';

/* ==========================================================================
 * Structural access
 * ======================================================================= */

/**
 * Get the Students sheet, translating a missing tab (NOT_FOUND) into a
 * structural SERVER_ERROR -- a missing sheet is a configuration problem, not a
 * user-not-found problem.
 * @return {Sheet}
 */
function getStudentsSheet_() {
  var sheet;
  try {
    sheet = getSheet_(CONFIG.SHEETS.STUDENTS);
  } catch (err) {
    if (err && err.code === ERROR_CODES.NOT_FOUND) {
      throwError_('Students sheet "' + CONFIG.SHEETS.STUDENTS + '" is missing from the spreadsheet.',
        ERROR_CODES.SERVER_ERROR, { sheet: CONFIG.SHEETS.STUDENTS, reason: 'sheet-missing' });
    }
    throw err;
  }
  return sheet;
}

/**
 * Validate the Students sheet exists and contains its full canonical schema.
 * @return {{sheet: Sheet, headers: string[]}}
 * @throws {Error} SERVER_ERROR when the sheet or any required column is missing.
 */
function getStudentsHeaders_() {
  var sheet = getStudentsSheet_();
  var headers = getHeaders_(sheet);
  if (headers.length === 0) {
    throwError_('Students sheet has no header row; cannot map records to columns.',
      ERROR_CODES.SERVER_ERROR, { sheet: CONFIG.SHEETS.STUDENTS });
  }
  var missing = STUDENT_COLUMNS.filter(function (c) { return headers.indexOf(c) === -1; });
  if (missing.length > 0) {
    throwError_('Students sheet is missing recommended column(s): ' + missing.join(', ') + '.',
      ERROR_CODES.SERVER_ERROR, { sheet: CONFIG.SHEETS.STUDENTS, missingColumns: missing });
  }
  return { sheet: sheet, headers: headers };
}

/**
 * Read every student row as header-keyed objects, skipping blank rows.
 * @return {Object[]}
 */
function readAllStudents_() {
  var ctx = getStudentsHeaders_();
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
 * Look up a student by Student_ID.
 * @return {{sheetRow: number, record: Object}|null}
 */
function findStudentById_(id) {
  return findRowById_(CONFIG.SHEETS.STUDENTS, id, STUDENT_ID_COLUMN);
}

/**
 * Validate (without changing) the mutable student fields. Defaults are applied
 * by callers, so this never inserts a value.
 */
function normalizeStudentRecord_(record) {
  if (!isBlank_(record.Status)) {
    assertOneOf_(toTrimmedString_(record.Status), CONFIG.VALUES.STUDENT_STATUS, 'Status');
    record.Status = toTrimmedString_(record.Status);
  }
  if (!isBlank_(record.Guardian_Email)) {
    assertEmail_(record.Guardian_Email, 'Guardian_Email');
    record.Guardian_Email = toTrimmedString_(record.Guardian_Email);
  }
}

/* ==========================================================================
 * Actions
 * ======================================================================= */

function handleStudentsList_(payload) {
  requirePermission_('STUDENTS.READ');
  var records = readAllStudents_();
  return success({ students: records, count: records.length },
    'Retrieved ' + records.length + ' student(s)');
}

function handleStudentsGet_(payload) {
  requirePermission_('STUDENTS.READ');
  var id = toTrimmedString_(payload && payload.Student_ID);
  if (id === '') {
    throwError_('Student_ID is required to look up a student.',
      ERROR_CODES.VALIDATION_ERROR, { field: STUDENT_ID_COLUMN });
  }
  getStudentsHeaders_();
  var found = findStudentById_(id);
  if (!found) {
    throwError_('No student found for Student_ID "' + id + '".',
      ERROR_CODES.NOT_FOUND, { [STUDENT_ID_COLUMN]: id });
  }
  return success(found.record, 'Student retrieved');
}

function handleStudentsCreate_(payload) {
  requirePermission_('STUDENTS.CREATE');
  getStudentsHeaders_();
  if (!isBlank_(payload && payload.Student_ID)) {
    throwError_('Student_ID is server-generated and cannot be supplied by the client.',
      ERROR_CODES.VALIDATION_ERROR, { field: STUDENT_ID_COLUMN });
  }
  assertRequired_(payload, REQUIRED_STUDENT_FIELDS);
  var record = Object.assign({}, payload);
  record.Student_ID = generateId_('STU');
  if (isBlank_(record.Status)) record.Status = CONFIG.STUDENT_STATUS.ACTIVE;
  normalizeStudentRecord_(record);
  if (findStudentById_(record.Student_ID)) {
    throwError_('A student with Student_ID "' + record.Student_ID + '" already exists.',
      ERROR_CODES.CONFLICT, { [STUDENT_ID_COLUMN]: record.Student_ID });
  }
  var created = appendRow_(CONFIG.SHEETS.STUDENTS, record);
  return success(created, 'Student created');
}

function handleStudentsUpdate_(payload) {
  requirePermission_('STUDENTS.UPDATE');
  var id = toTrimmedString_(payload && payload.Student_ID);
  if (id === '') {
    throwError_('Student_ID is required to update a student.',
      ERROR_CODES.VALIDATION_ERROR, { field: STUDENT_ID_COLUMN });
  }
  // Student_ID identifies the row only; it can never be changed by an update.
  var patch = Object.assign({}, payload);
  delete patch[STUDENT_ID_COLUMN];

  // The read-modify-write must be atomic: hold the lock across the whole
  // sequence so two concurrent updates cannot lose each other's fields.
  return withScriptLock_(function () {
    var ctx = getStudentsHeaders_();
    var existing = findStudentById_(id);
    if (!existing) {
      throwError_('No student found for Student_ID "' + id + '".',
        ERROR_CODES.NOT_FOUND, { [STUDENT_ID_COLUMN]: id });
    }
    var merged = Object.assign({}, existing.record, patch);
    normalizeStudentRecord_(merged);
    // recordToValues_ rejects unknown columns -> VALIDATION_ERROR and aligns order.
    var values = recordToValues_(CONFIG.SHEETS.STUDENTS, ctx.headers, merged);
    ctx.sheet.getRange(existing.sheetRow, 1, 1, ctx.headers.length).setValues([values]);
    return success(rowToObject_(ctx.headers, values), 'Student updated');
  });
}

function handleStudentsWithdraw_(payload) {
  requirePermission_('STUDENTS.WITHDRAW');
  var id = toTrimmedString_(payload && payload.Student_ID);
  if (id === '') {
    throwError_('Student_ID is required to withdraw a student.',
      ERROR_CODES.VALIDATION_ERROR, { field: STUDENT_ID_COLUMN });
  }
  // Soft-delete: never remove the row. Record the withdrawal date.
  return withScriptLock_(function () {
    var ctx = getStudentsHeaders_();
    var existing = findStudentById_(id);
    if (!existing) {
      throwError_('No student found for Student_ID "' + id + '".',
        ERROR_CODES.NOT_FOUND, { [STUDENT_ID_COLUMN]: id });
    }
    var merged = Object.assign({}, existing.record, {
      Status: CONFIG.STUDENT_STATUS.WITHDRAWN,
      Withdrawal_Date: formatDate_(now_())
    });
    var values = recordToValues_(CONFIG.SHEETS.STUDENTS, ctx.headers, merged);
    ctx.sheet.getRange(existing.sheetRow, 1, 1, ctx.headers.length).setValues([values]);
    return success(rowToObject_(ctx.headers, values), 'Student withdrawn');
  });
}
