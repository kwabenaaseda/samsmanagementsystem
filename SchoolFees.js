/**
 * SchoolFees.js
 *
 * Handles school fee OBLIGATIONS (fee accounts) and, via SchoolFeePayments.js,
 * the payment transactions recorded against them (Option 2 model).
 * Schema: Payment_ID, Student_ID, Academic_Year, Term, Amount_Due, Amount_Paid,
 *         Balance, Payment_Date, Payment_Method, Reference, Status, Recorded_By, Notes
 *
 * ONE ROW = ONE FEE ACCOUNT for a student / academic year / term:
 *   - Amount_Due   = the Fee Amount charged (editable; may be increased later).
 *   - Amount_Paid  = SERVER-DERIVED: sum of non-voided School_Fee_Payments
 *                    rows whose Fee_ID references this row. Never client-supplied.
 *   - Balance      = SERVER-DERIVED: Amount_Due - Amount_Paid (outstanding).
 *   - Payment_Date / Payment_Method / Reference are legacy compatibility
 *     columns from the single-row model; new flows record them on the
 *     School_Fee_Payments rows instead.
 *
 * Payment_ID format: SF-001, SF-002, ... (server-generated, immutable)
 * Student_ID must exist and not be withdrawn for new accounts
 * Void is soft correction: Status = 'Voided', record preserved
 * Payment_ID, Amount_Paid, Balance and Recorded_By are server-controlled: a
 * client-supplied value is ignored (or rejected for Payment_ID), and unknown payload columns are
 * rejected with VALIDATION_ERROR, matching the Phase 3 convention
 */

var SCHOOL_FEES_COLUMNS = [
  'Payment_ID', 'Student_ID', 'Academic_Year', 'Term', 'Amount_Due',
  'Amount_Paid', 'Balance', 'Payment_Date', 'Payment_Method', 'Reference',
  'Status', 'Recorded_By', 'Notes'
];

var SCHOOL_FEES_ID_COLUMN = 'Payment_ID';
var SCHOOL_FEES_STUDENT_COLUMN = 'Student_ID';

var SCHOOL_FEES_CREATE_REQUIRED = [
  'Student_ID', 'Academic_Year', 'Term', 'Amount_Due'
];

var SCHOOL_FEES_PAYMENT_METHODS = [
  'Cash', 'Bank Transfer', 'Mobile Money', 'Other'
];

var SCHOOL_FEES_STATUS_VALUES = [
  'Unpaid', 'Partial', 'Paid', 'Voided'
];

function getSchoolFeesSheet_() {
  try {
    return getSheet_(CONFIG.SHEETS.SCHOOL_FEES);
  } catch (err) {
    if (err && err.code === ERROR_CODES.NOT_FOUND) {
      throwError_(
        'School_Fees sheet is missing.',
        ERROR_CODES.SERVER_ERROR,
        { sheet: CONFIG.SHEETS.SCHOOL_FEES, reason: 'sheet-missing' }
      );
    }
    throw err;
  }
}

function getSchoolFeesHeaders_() {
  var sheet = getSchoolFeesSheet_();
  var headers = getHeaders_(sheet);
  if (headers.length === 0) {
    throwError_(
      'School_Fees sheet has no headers.',
      ERROR_CODES.SERVER_ERROR,
      { sheet: CONFIG.SHEETS.SCHOOL_FEES }
    );
  }
  var missing = SCHOOL_FEES_COLUMNS.filter(function (c) {
    return headers.indexOf(c) === -1;
  });
  if (missing.length > 0) {
    throwError_(
      'School_Fees sheet missing columns: ' + missing.join(', '),
      ERROR_CODES.SERVER_ERROR,
      { sheet: CONFIG.SHEETS.SCHOOL_FEES, missingColumns: missing }
    );
  }
  return headers;
}

function normalizeSchoolFeesRecord_(record) {
  record.Amount_Due = Number(record.Amount_Due);
  record.Amount_Paid = Number(record.Amount_Paid);
  record.Balance = record.Amount_Due - record.Amount_Paid;
  record.Payment_Date = toTrimmedString_(record.Payment_Date);
  record.Payment_Method = toTrimmedString_(record.Payment_Method);
  record.Reference = toTrimmedString_(record.Reference);
  record.Status = toTrimmedString_(record.Status);
  record.Student_ID = toTrimmedString_(record.Student_ID);
  record.Academic_Year = toTrimmedString_(record.Academic_Year);
  record.Term = toTrimmedString_(record.Term);
  record.Recorded_By = toTrimmedString_(record.Recorded_By);
  record.Notes = toTrimmedString_(record.Notes);
  return record;
}

/**
 * Reject payload keys that are not School_Fees columns.
 *
 * Mirrors the recordToValues_ rule Phase 3 relies on: a client may never invent
 * columns, and silently dropping a typo would hide a frontend/backend mismatch,
 * so every offending key is named in the error.
 *
 * @param {string[]} headers Canonical School_Fees headers.
 * @param {Object} payload Request payload.
 * @return {Object} The same payload, for chaining.
 * @throws {Error} VALIDATION_ERROR listing the unknownColumns.
 */
function assertKnownSchoolFeesFields_(headers, payload) {
  var unknown = Object.keys(payload).filter(function (key) {
    return headers.indexOf(key) === -1;
  });
  if (unknown.length > 0) {
    throwError_(
      'Unknown column(s) for sheet "' + CONFIG.SHEETS.SCHOOL_FEES + '": ' + unknown.join(', ') + '.',
      ERROR_CODES.VALIDATION_ERROR,
      { sheet: CONFIG.SHEETS.SCHOOL_FEES, unknownColumns: unknown, validColumns: headers }
    );
  }
  return payload;
}

/**
 * Derive Status from the amounts. Balance/Status are never client-supplied.
 * @param {number} amountDue Amount due.
 * @param {number} amountPaid Amount paid.
 * @return {string} One of SCHOOL_FEES_STATUS_VALUES.
 */
function deriveSchoolFeesStatus_(amountDue, amountPaid) {
  if (amountDue <= 0) return 'Paid';
  if (amountPaid >= amountDue) return 'Paid';
  if (amountPaid > 0) return 'Partial';
  return 'Unpaid';
}

/**
 * Re-derive Amount_Paid/Balance/Status for the given obligation records from
 * the payment ledger (School_Fee_Payments), reading the ledger ONCE for the
 * whole batch. The ledger is authoritative: a stale or hand-edited
 * Amount_Paid/Balance cell on the School_Fees row can never be served.
 *
 * A 'Voided' obligation keeps its Status (voiding an account is itself a
 * correction and must stay visible as such); live accounts get
 * Unpaid / Partial / Paid from the recomputed totals.
 */
function enrichSchoolFeeRecords_(records) {
  if (!records || records.length === 0) return records;
  var totals = {};
  try {
    var headers = getFeePaymentsHeaders_();
    var sheet = getFeePaymentsSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
      for (var i = 0; i < values.length; i++) {
        var p = normalizeFeePaymentRecord_(rowToObject_(headers, values[i]));
        if (p.Status === 'Voided') continue;
        totals[p.Fee_ID] = (totals[p.Fee_ID] || 0) + p.Amount;
      }
    }
  } catch (err) {
    // An absent School_Fee_Payments sheet simply means nothing paid yet.
    if (!(err && err.details && err.details.sheet === CONFIG.SHEETS.SCHOOL_FEE_PAYMENTS)) {
      throw err;
    }
  }
  for (var j = 0; j < records.length; j++) {
    var r = records[j];
    var paid = totals[r.Payment_ID] || 0;
    r.Amount_Paid = paid;
    r.Balance = r.Amount_Due - paid;
    if (toTrimmedString_(r.Status) !== 'Voided') {
      r.Status = deriveSchoolFeesStatus_(r.Amount_Due, paid);
    }
  }
  return records;
}

function nextSchoolFeesId_() {
  var sheet = getSchoolFeesSheet_();
  var headers = getHeaders_(sheet);
  var idIdx = headers.indexOf(SCHOOL_FEES_ID_COLUMN);
  var lastRow = sheet.getLastRow();
  var maxNum = 0;
  if (lastRow >= 2) {
    var values = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      var v = toTrimmedString_(values[i][0]);
      var m = v.match(/^SF-(\d+)$/);
      if (m) {
        var n = parseInt(m[1], 10);
        if (n > maxNum) maxNum = n;
      }
    }
  }
  return 'SF-' + String(maxNum + 1).padStart(3, '0');
}

function findSchoolFeeById_(id) {
  var found = findRowById_(CONFIG.SHEETS.SCHOOL_FEES, id, SCHOOL_FEES_ID_COLUMN);
  // Balance is only ever the server-side calculation, so a stale or hand-edited
  // Balance cell can never be served to (or trusted by) a caller.
  if (found) {
    normalizeSchoolFeesRecord_(found.record);
    enrichSchoolFeeRecords_([found.record]);
  }
  return found;
}

function readAllSchoolFees_(headers) {
  var sheet = getSchoolFeesSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var records = [];
  for (var i = 0; i < values.length; i++) {
    records.push(normalizeSchoolFeesRecord_(rowToObject_(headers, values[i])));
  }
  return enrichSchoolFeeRecords_(records);
}

function handleSchoolFeesList_(payload) {
  requirePermission_('SCHOOL_FEES.READ');
  var headers = getSchoolFeesHeaders_();
  var records = readAllSchoolFees_(headers);

  if (payload && payload.Student_ID) {
    var sid = toTrimmedString_(payload.Student_ID);
    records = records.filter(function (r) {
      return toTrimmedString_(r.Student_ID) === sid;
    });
  }
  if (payload && payload.Academic_Year) {
    var ay = toTrimmedString_(payload.Academic_Year);
    records = records.filter(function (r) {
      return toTrimmedString_(r.Academic_Year) === ay;
    });
  }
  if (payload && payload.Term) {
    var term = toTrimmedString_(payload.Term);
    records = records.filter(function (r) {
      return toTrimmedString_(r.Term) === term;
    });
  }
  if (payload && payload.Status) {
    var status = toTrimmedString_(payload.Status);
    records = records.filter(function (r) {
      return toTrimmedString_(r.Status) === status;
    });
  }
  return success(records, 'School fees retrieved');
}

function handleSchoolFeesGet_(payload) {
  requirePermission_('SCHOOL_FEES.READ');
  var id = toTrimmedString_(payload && payload.Payment_ID);
  if (!id) {
    throwError_(
      'Payment_ID is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Payment_ID' }
    );
  }
  var result = findSchoolFeeById_(id);
  if (!result) {
    throwError_(
      'School fee payment not found: ' + id,
      ERROR_CODES.NOT_FOUND,
      { Payment_ID: id }
    );
  }
  return success(result.record, 'School fee payment retrieved');
}

function validateStudentForPayment_(studentId, allowWithdrawn) {
  if (!studentId || toTrimmedString_(studentId) === '') {
    throwError_(
      'Student_ID is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Student_ID' }
    );
  }
  var student = findStudentById_(toTrimmedString_(studentId));
  if (!student) {
    throwError_(
      'Student not found: ' + studentId,
      ERROR_CODES.NOT_FOUND,
      { Student_ID: studentId }
    );
  }
  if (!allowWithdrawn) {
    var status = toTrimmedString_(student.record.Status);
    if (status === 'Withdrawn') {
      throwError_(
        'Cannot create payment for withdrawn student: ' + studentId,
        ERROR_CODES.VALIDATION_ERROR,
        { Student_ID: studentId, reason: 'student-withdrawn' }
      );
    }
  }
  return student;
}

function handleSchoolFeesCreate_(payload) {
  requirePermission_('SCHOOL_FEES.CREATE');
  if (!payload || typeof payload !== 'object') {
    throwError_(
      'Payload is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { reason: 'missing-payload' }
    );
  }
  if (payload.Payment_ID && toTrimmedString_(payload.Payment_ID) !== '') {
    throwError_(
      'Payment_ID cannot be set by client. It is server-generated.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Payment_ID' }
    );
  }
  assertRequired_(payload, SCHOOL_FEES_CREATE_REQUIRED);
  assertKnownSchoolFeesFields_(getSchoolFeesHeaders_(), payload);

  // Creating an account records an OBLIGATION only. Money received is a
  // School_Fee_Payments row created via feePayments.create against this
  // account, so legacy payment fields are rejected rather than silently
  // double-tracked on the account row.
  ['Amount_Paid', 'Payment_Method', 'Payment_Date'].forEach(function (field) {
    if (payload[field] !== undefined) {
      throwError_(
        field + ' belongs to payment transactions. Record money received via feePayments.create against this fee account.',
        ERROR_CODES.VALIDATION_ERROR,
        { field: field, reason: 'use-feePayments-create' }
      );
    }
  });

  var studentId = toTrimmedString_(payload.Student_ID);
  validateStudentForPayment_(studentId, false);

  var amountDue = Number(payload.Amount_Due);
  if (isNaN(amountDue) || amountDue < 0) {
    throwError_(
      'Amount_Due must be a non-negative number.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Amount_Due', value: payload.Amount_Due }
    );
  }

  // A brand-new obligation has received no money yet. Payments are recorded
  // exclusively through feePayments.create against this account.
  var amountPaid = 0;

  // Legacy compatibility columns from the single-row model. New flows record
  // method/date/reference on the School_Fee_Payments rows instead.
  var paymentMethod = '';
  var paymentDate = formatDate_(now_());

  var academicYear = toTrimmedString_(payload.Academic_Year);
  if (!academicYear) {
    throwError_(
      'Academic_Year is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Academic_Year' }
    );
  }

  var term = toTrimmedString_(payload.Term);
  if (!term) {
    throwError_(
      'Term is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Term' }
    );
  }

  var user = requireAuthentication_();
  var recordedBy = user.staffId || user.userId || 'SYSTEM';

  var record = {
    Student_ID: studentId,
    Academic_Year: academicYear,
    Term: term,
    Amount_Due: amountDue,
    Amount_Paid: amountPaid,
    Balance: amountDue - amountPaid,
    Payment_Date: paymentDate,
    Payment_Method: paymentMethod,
    Reference: toTrimmedString_(payload.Reference || ''),
    Status: deriveSchoolFeesStatus_(amountDue, amountPaid),
    Recorded_By: recordedBy,
    Notes: toTrimmedString_(payload.Notes || '')
  };

  return withScriptLock_(function () {
    var id = nextSchoolFeesId_();
    record.Payment_ID = id;
    var existing = findSchoolFeeById_(id);
    if (existing) {
      throwError_(
        'Duplicate Payment_ID generated: ' + id,
        ERROR_CODES.CONFLICT,
        { Payment_ID: id }
      );
    }
    var sheet = getSchoolFeesSheet_();
    var created = appendRow_(CONFIG.SHEETS.SCHOOL_FEES, record);
    return success(created, 'School fee account created');
  });
}

function handleSchoolFeesUpdate_(payload) {
  requirePermission_('SCHOOL_FEES.UPDATE');
  if (!payload || typeof payload !== 'object') {
    throwError_(
      'Payload is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { reason: 'missing-payload' }
    );
  }
  var id = toTrimmedString_(payload.Payment_ID);
  if (!id) {
    throwError_(
      'Payment_ID is required for update.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Payment_ID' }
    );
  }

  // The whole read-modify-write runs inside the critical section and the stored
  // record is re-read after the lock is taken: a void or update committed by
  // another user while this call waited for the lock must be visible before any
  // decision (NOT_FOUND, voided-payment, recalculation) is made.
  return withScriptLock_(function () {
    var result = findSchoolFeeById_(id);
    if (!result) {
      throwError_(
        'School fee payment not found: ' + id,
        ERROR_CODES.NOT_FOUND,
        { Payment_ID: id }
      );
    }

    if (toTrimmedString_(result.record.Status) === 'Voided') {
      throwError_(
        'Cannot update a voided payment: ' + id,
        ERROR_CODES.VALIDATION_ERROR,
        { Payment_ID: id, reason: 'voided-payment' }
      );
    }

    // Payment_ID is the lookup key, so it can never be changed: the row value is
    // rebuilt from the stored record below, never from the payload.
    // Reject invented columns (recordToValues_ semantics) before touching the row.
    assertKnownSchoolFeesFields_(getSchoolFeesHeaders_(), payload);

    var current = result.record;
    // Payment_ID is carried through explicitly: it identifies the row and is
    // immutable. Omitting it would let recordToValues_ blank the ID column and
    // orphan the very row being updated.
    var updated = { Payment_ID: current.Payment_ID };

    if (payload.Student_ID !== undefined) {
      var newSid = toTrimmedString_(payload.Student_ID);
      if (newSid && newSid !== current.Student_ID) {
        validateStudentForPayment_(newSid, false);
      }
      updated.Student_ID = newSid || current.Student_ID;
    } else {
      updated.Student_ID = current.Student_ID;
    }

    if (payload.Academic_Year !== undefined) {
      var ay = toTrimmedString_(payload.Academic_Year);
      if (!ay) {
        throwError_(
          'Academic_Year cannot be empty.',
          ERROR_CODES.VALIDATION_ERROR,
          { field: 'Academic_Year' }
        );
      }
      updated.Academic_Year = ay;
    } else {
      updated.Academic_Year = current.Academic_Year;
    }

    if (payload.Term !== undefined) {
      var term = toTrimmedString_(payload.Term);
      if (!term) {
        throwError_(
          'Term cannot be empty.',
          ERROR_CODES.VALIDATION_ERROR,
          { field: 'Term' }
        );
      }
      updated.Term = term;
    } else {
      updated.Term = current.Term;
    }

    var amountDue = current.Amount_Due;

    if (payload.Amount_Due !== undefined) {
      amountDue = Number(payload.Amount_Due);
      if (isNaN(amountDue) || amountDue < 0) {
        throwError_(
          'Amount_Due must be a non-negative number.',
          ERROR_CODES.VALIDATION_ERROR,
          { field: 'Amount_Due', value: payload.Amount_Due }
        );
      }
    }

    // Amount_Paid/Balance are server-derived from the payment ledger and can
    // never be set by a client. Increasing the Fee Amount is allowed (a fee
    // increase recalculates Outstanding without touching payment history);
    // reducing it below Total Paid would corrupt the account, so it is
    // rejected for MVP1 (no credit/overpayment concept).
    if (payload.Amount_Paid !== undefined) {
      throwError_(
        'Amount_Paid is server-derived from recorded payments and cannot be set directly.',
        ERROR_CODES.VALIDATION_ERROR,
        { field: 'Amount_Paid', reason: 'server-derived-field' }
      );
    }

    var totalPaid = computeFeeTotalPaid_(current.Payment_ID);
    if (amountDue < totalPaid) {
      throwError_(
        'Fee Amount (' + amountDue + ') cannot be less than Total Paid (' + totalPaid + ').',
        ERROR_CODES.VALIDATION_ERROR,
        { field: 'Amount_Due', value: amountDue, totalPaid: totalPaid, reason: 'fee-below-total-paid' }
      );
    }

    updated.Amount_Due = amountDue;
    updated.Amount_Paid = totalPaid;
    updated.Balance = amountDue - totalPaid;

    if (payload.Payment_Date !== undefined) {
      var pd = toTrimmedString_(payload.Payment_Date);
      if (!pd) {
        throwError_(
          'Payment_Date is required.',
          ERROR_CODES.VALIDATION_ERROR,
          { field: 'Payment_Date' }
        );
      }
      var parsed = toDate_(pd);
      if (!parsed) {
        throwError_(
          'Invalid Payment_Date: ' + pd,
          ERROR_CODES.VALIDATION_ERROR,
          { field: 'Payment_Date', value: pd }
        );
      }
      updated.Payment_Date = formatDate_(parsed);
    } else {
      updated.Payment_Date = current.Payment_Date;
    }

    if (payload.Payment_Method !== undefined) {
      var pm = toTrimmedString_(payload.Payment_Method);
      if (SCHOOL_FEES_PAYMENT_METHODS.indexOf(pm) === -1) {
        throwError_(
          'Invalid Payment_Method: ' + pm,
          ERROR_CODES.VALIDATION_ERROR,
          { field: 'Payment_Method', value: pm, allowedValues: SCHOOL_FEES_PAYMENT_METHODS }
        );
      }
      updated.Payment_Method = pm;
    } else {
      updated.Payment_Method = current.Payment_Method;
    }

    if (payload.Reference !== undefined) {
      updated.Reference = toTrimmedString_(payload.Reference);
    } else {
      updated.Reference = current.Reference;
    }

    if (payload.Status !== undefined) {
      var s = toTrimmedString_(payload.Status);
      if (SCHOOL_FEES_STATUS_VALUES.indexOf(s) === -1) {
        throwError_(
          'Invalid Status: ' + s,
          ERROR_CODES.VALIDATION_ERROR,
          { field: 'Status', value: s, allowedValues: SCHOOL_FEES_STATUS_VALUES }
        );
      }
      updated.Status = s;
    } else {
      updated.Status = deriveSchoolFeesStatus_(amountDue, totalPaid);
    }

    updated.Recorded_By = current.Recorded_By;

    if (payload.Notes !== undefined) {
      updated.Notes = toTrimmedString_(payload.Notes);
    } else {
      updated.Notes = current.Notes;
    }

    var sheet = getSchoolFeesSheet_();
    var headers = getSchoolFeesHeaders_();
    var values = recordToValues_(CONFIG.SHEETS.SCHOOL_FEES, headers, updated);
    sheet.getRange(result.sheetRow, 1, 1, headers.length).setValues([values]);
    var refreshed = findSchoolFeeById_(id);
    return success(refreshed.record, 'School fee payment updated');
  });
}

function handleSchoolFeesVoid_(payload) {
  requirePermission_('SCHOOL_FEES.VOID');
  if (!payload || typeof payload !== 'object') {
    throwError_(
      'Payload is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { reason: 'missing-payload' }
    );
  }
  var id = toTrimmedString_(payload.Payment_ID);
  if (!id) {
    throwError_(
      'Payment_ID is required for void.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Payment_ID' }
    );
  }

  // The already-voided decision is taken inside the critical section against a
  // freshly read record, so a void committed by another user while this call
  // waited for the lock is seen before we decide (and can never be overwritten).
  return withScriptLock_(function () {
    var result = findSchoolFeeById_(id);
    if (!result) {
      throwError_(
        'School fee payment not found: ' + id,
        ERROR_CODES.NOT_FOUND,
        { Payment_ID: id }
      );
    }

    if (toTrimmedString_(result.record.Status) === 'Voided') {
      throwError_(
        'Payment already voided: ' + id,
        ERROR_CODES.VALIDATION_ERROR,
        { Payment_ID: id, reason: 'already-voided' }
      );
    }

    var updated = Object.assign({}, result.record, { Status: 'Voided' });

    var sheet = getSchoolFeesSheet_();
    var headers = getSchoolFeesHeaders_();
    var values = recordToValues_(CONFIG.SHEETS.SCHOOL_FEES, headers, updated);
    sheet.getRange(result.sheetRow, 1, 1, headers.length).setValues([values]);
    var refreshed = findSchoolFeeById_(id);
    return success(refreshed.record, 'School fee payment voided');
  });
}
