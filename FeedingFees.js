/**
 * FeedingFees.js
 *
 * Handles feeding fee payment records.
 * Schema: Payment_ID, Student_ID, Academic_Year, Term, Amount_Due, Amount_Paid,
 *         Balance, Payment_Date, Payment_Method, Reference, Status, Recorded_By, Notes
 *
 * Payment_ID format: FF-001, FF-002, ... (server-generated, immutable)
 * Balance = Amount_Due - Amount_Paid (calculated server-side)
 * Amount_Paid must not exceed Amount_Due
 * Student_ID must exist and not be withdrawn for new payments
 * Void is soft correction: Status = 'Voided', record preserved
 * Payment_ID, Balance and Recorded_By are server-controlled: a client-supplied
 * value is ignored (or rejected for Payment_ID), and unknown payload columns are
 * rejected with VALIDATION_ERROR, matching the Phase 3 convention
 */

var FEEDING_FEES_COLUMNS = [
  'Payment_ID', 'Student_ID', 'Academic_Year', 'Term', 'Amount_Due',
  'Amount_Paid', 'Balance', 'Payment_Date', 'Payment_Method', 'Reference',
  'Status', 'Recorded_By', 'Notes'
];

var FEEDING_FEES_ID_COLUMN = 'Payment_ID';
var FEEDING_FEES_STUDENT_COLUMN = 'Student_ID';

var FEEDING_FEES_CREATE_REQUIRED = [
  'Student_ID', 'Academic_Year', 'Term', 'Amount_Due', 'Amount_Paid',
  'Payment_Method', 'Payment_Date'
];

var FEEDING_FEES_PAYMENT_METHODS = [
  'Cash', 'Bank Transfer', 'Mobile Money', 'Other'
];

var FEEDING_FEES_STATUS_VALUES = [
  'Unpaid', 'Partial', 'Paid', 'Voided'
];

function getFeedingFeesSheet_() {
  try {
    return getSheet_(CONFIG.SHEETS.FEEDING_FEES);
  } catch (err) {
    if (err && err.code === ERROR_CODES.NOT_FOUND) {
      throwError_(
        'Feeding_Fees sheet is missing.',
        ERROR_CODES.SERVER_ERROR,
        { sheet: CONFIG.SHEETS.FEEDING_FEES, reason: 'sheet-missing' }
      );
    }
    throw err;
  }
}

function getFeedingFeesHeaders_() {
  var sheet = getFeedingFeesSheet_();
  var headers = getHeaders_(sheet);
  if (headers.length === 0) {
    throwError_(
      'Feeding_Fees sheet has no headers.',
      ERROR_CODES.SERVER_ERROR,
      { sheet: CONFIG.SHEETS.FEEDING_FEES }
    );
  }
  var missing = FEEDING_FEES_COLUMNS.filter(function (c) {
    return headers.indexOf(c) === -1;
  });
  if (missing.length > 0) {
    throwError_(
      'Feeding_Fees sheet missing columns: ' + missing.join(', '),
      ERROR_CODES.SERVER_ERROR,
      { sheet: CONFIG.SHEETS.FEEDING_FEES, missingColumns: missing }
    );
  }
  return headers;
}

function normalizeFeedingFeesRecord_(record) {
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
 * Reject payload keys that are not Feeding_Fees columns.
 *
 * Mirrors the recordToValues_ rule Phase 3 relies on: a client may never invent
 * columns, and silently dropping a typo would hide a frontend/backend mismatch,
 * so every offending key is named in the error.
 *
 * @param {string[]} headers Canonical Feeding_Fees headers.
 * @param {Object} payload Request payload.
 * @return {Object} The same payload, for chaining.
 * @throws {Error} VALIDATION_ERROR listing the unknownColumns.
 */
function assertKnownFeedingFeesFields_(headers, payload) {
  var unknown = Object.keys(payload).filter(function (key) {
    return headers.indexOf(key) === -1;
  });
  if (unknown.length > 0) {
    throwError_(
      'Unknown column(s) for sheet "' + CONFIG.SHEETS.FEEDING_FEES + '": ' + unknown.join(', ') + '.',
      ERROR_CODES.VALIDATION_ERROR,
      { sheet: CONFIG.SHEETS.FEEDING_FEES, unknownColumns: unknown, validColumns: headers }
    );
  }
  return payload;
}

/**
 * Derive Status from the amounts. Balance/Status are never client-supplied.
 * @param {number} amountDue Amount due.
 * @param {number} amountPaid Amount paid.
 * @return {string} One of FEEDING_FEES_STATUS_VALUES.
 */
function deriveFeedingFeesStatus_(amountDue, amountPaid) {
  if (amountDue <= 0) return 'Paid';
  if (amountPaid >= amountDue) return 'Paid';
  if (amountPaid > 0) return 'Partial';
  return 'Unpaid';
}

function nextFeedingFeesId_() {
  var sheet = getFeedingFeesSheet_();
  var headers = getHeaders_(sheet);
  var idIdx = headers.indexOf(FEEDING_FEES_ID_COLUMN);
  var lastRow = sheet.getLastRow();
  var maxNum = 0;
  if (lastRow >= 2) {
    var values = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      var v = toTrimmedString_(values[i][0]);
      var m = v.match(/^FF-(\d+)$/);
      if (m) {
        var n = parseInt(m[1], 10);
        if (n > maxNum) maxNum = n;
      }
    }
  }
  return 'FF-' + String(maxNum + 1).padStart(3, '0');
}

function findFeedingFeeById_(id) {
  var found = findRowById_(CONFIG.SHEETS.FEEDING_FEES, id, FEEDING_FEES_ID_COLUMN);
  // Balance is only ever the server-side calculation, so a stale or hand-edited
  // Balance cell can never be served to (or trusted by) a caller.
  if (found) normalizeFeedingFeesRecord_(found.record);
  return found;
}

function readAllFeedingFees_(headers) {
  var sheet = getFeedingFeesSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var records = [];
  for (var i = 0; i < values.length; i++) {
    records.push(normalizeFeedingFeesRecord_(rowToObject_(headers, values[i])));
  }
  return records;
}

function handleFeedingFeesList_(payload) {
  requirePermission_('FEEDING_FEES.READ');
  var headers = getFeedingFeesHeaders_();
  var records = readAllFeedingFees_(headers);

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
  return success(records, 'Feeding fees retrieved');
}

function handleFeedingFeesGet_(payload) {
  requirePermission_('FEEDING_FEES.READ');
  var id = toTrimmedString_(payload && payload.Payment_ID);
  if (!id) {
    throwError_(
      'Payment_ID is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Payment_ID' }
    );
  }
  var result = findFeedingFeeById_(id);
  if (!result) {
    throwError_(
      'Feeding fee payment not found: ' + id,
      ERROR_CODES.NOT_FOUND,
      { Payment_ID: id }
    );
  }
  return success(result.record, 'Feeding fee payment retrieved');
}

function validateStudentForFeedingPayment_(studentId, allowWithdrawn) {
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

function handleFeedingFeesCreate_(payload) {
  requirePermission_('FEEDING_FEES.CREATE');
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
  assertRequired_(payload, FEEDING_FEES_CREATE_REQUIRED);
  assertKnownFeedingFeesFields_(getFeedingFeesHeaders_(), payload);

  var studentId = toTrimmedString_(payload.Student_ID);
  validateStudentForFeedingPayment_(studentId, false);

  var amountDue = Number(payload.Amount_Due);
  if (isNaN(amountDue) || amountDue < 0) {
    throwError_(
      'Amount_Due must be a non-negative number.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Amount_Due', value: payload.Amount_Due }
    );
  }

  var amountPaid = Number(payload.Amount_Paid);
  if (isNaN(amountPaid) || amountPaid < 0) {
    throwError_(
      'Amount_Paid must be a non-negative number.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Amount_Paid', value: payload.Amount_Paid }
    );
  }

  if (amountPaid > amountDue) {
    throwError_(
      'Amount_Paid (' + amountPaid + ') cannot exceed Amount_Due (' + amountDue + ').',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Amount_Paid', value: amountPaid, reason: 'amount-paid-exceeds-due' }
    );
  }

  var paymentMethod = toTrimmedString_(payload.Payment_Method);
  if (FEEDING_FEES_PAYMENT_METHODS.indexOf(paymentMethod) === -1) {
    throwError_(
      'Invalid Payment_Method: ' + paymentMethod,
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Payment_Method', value: paymentMethod, allowedValues: FEEDING_FEES_PAYMENT_METHODS }
    );
  }

  var paymentDate = toTrimmedString_(payload.Payment_Date);
  if (!paymentDate) {
    throwError_(
      'Payment_Date is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Payment_Date' }
    );
  }
  var parsedDate = toDate_(paymentDate);
  if (!parsedDate) {
    throwError_(
      'Invalid Payment_Date: ' + paymentDate,
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Payment_Date', value: paymentDate }
    );
  }
  paymentDate = formatDate_(parsedDate);

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
    Status: deriveFeedingFeesStatus_(amountDue, amountPaid),
    Recorded_By: recordedBy,
    Notes: toTrimmedString_(payload.Notes || '')
  };

  return withScriptLock_(function () {
    var id = nextFeedingFeesId_();
    record.Payment_ID = id;
    var existing = findFeedingFeeById_(id);
    if (existing) {
      throwError_(
        'Duplicate Payment_ID generated: ' + id,
        ERROR_CODES.CONFLICT,
        { Payment_ID: id }
      );
    }
    var sheet = getFeedingFeesSheet_();
    var created = appendRow_(CONFIG.SHEETS.FEEDING_FEES, record);
    return success(created, 'Feeding fee payment created');
  });
}

function handleFeedingFeesUpdate_(payload) {
  requirePermission_('FEEDING_FEES.UPDATE');
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
    var result = findFeedingFeeById_(id);
    if (!result) {
      throwError_(
        'Feeding fee payment not found: ' + id,
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
    assertKnownFeedingFeesFields_(getFeedingFeesHeaders_(), payload);

    var current = result.record;
    // Payment_ID is carried through explicitly: it identifies the row and is
    // immutable. Omitting it would let recordToValues_ blank the ID column and
    // orphan the very row being updated.
    var updated = { Payment_ID: current.Payment_ID };

    if (payload.Student_ID !== undefined) {
      var newSid = toTrimmedString_(payload.Student_ID);
      if (newSid && newSid !== current.Student_ID) {
        validateStudentForFeedingPayment_(newSid, false);
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
    var amountPaid = current.Amount_Paid;

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

    if (payload.Amount_Paid !== undefined) {
      amountPaid = Number(payload.Amount_Paid);
      if (isNaN(amountPaid) || amountPaid < 0) {
        throwError_(
          'Amount_Paid must be a non-negative number.',
          ERROR_CODES.VALIDATION_ERROR,
          { field: 'Amount_Paid', value: payload.Amount_Paid }
        );
      }
    }

    if (amountPaid > amountDue) {
      throwError_(
        'Amount_Paid (' + amountPaid + ') cannot exceed Amount_Due (' + amountDue + ').',
        ERROR_CODES.VALIDATION_ERROR,
        { field: 'Amount_Paid', value: amountPaid, reason: 'amount-paid-exceeds-due' }
      );
    }

    updated.Amount_Due = amountDue;
    updated.Amount_Paid = amountPaid;
    updated.Balance = amountDue - amountPaid;

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
      if (FEEDING_FEES_PAYMENT_METHODS.indexOf(pm) === -1) {
        throwError_(
          'Invalid Payment_Method: ' + pm,
          ERROR_CODES.VALIDATION_ERROR,
          { field: 'Payment_Method', value: pm, allowedValues: FEEDING_FEES_PAYMENT_METHODS }
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
      if (FEEDING_FEES_STATUS_VALUES.indexOf(s) === -1) {
        throwError_(
          'Invalid Status: ' + s,
          ERROR_CODES.VALIDATION_ERROR,
          { field: 'Status', value: s, allowedValues: FEEDING_FEES_STATUS_VALUES }
        );
      }
      updated.Status = s;
    } else {
      updated.Status = deriveFeedingFeesStatus_(amountDue, amountPaid);
    }

    updated.Recorded_By = current.Recorded_By;

    if (payload.Notes !== undefined) {
      updated.Notes = toTrimmedString_(payload.Notes);
    } else {
      updated.Notes = current.Notes;
    }

    var sheet = getFeedingFeesSheet_();
    var headers = getFeedingFeesHeaders_();
    var values = recordToValues_(CONFIG.SHEETS.FEEDING_FEES, headers, updated);
    sheet.getRange(result.sheetRow, 1, 1, headers.length).setValues([values]);
    var refreshed = findFeedingFeeById_(id);
    return success(refreshed.record, 'Feeding fee payment updated');
  });
}

function handleFeedingFeesVoid_(payload) {
  requirePermission_('FEEDING_FEES.VOID');
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
    var result = findFeedingFeeById_(id);
    if (!result) {
      throwError_(
        'Feeding fee payment not found: ' + id,
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

    var sheet = getFeedingFeesSheet_();
    var headers = getFeedingFeesHeaders_();
    var values = recordToValues_(CONFIG.SHEETS.FEEDING_FEES, headers, updated);
    sheet.getRange(result.sheetRow, 1, 1, headers.length).setValues([values]);
    var refreshed = findFeedingFeeById_(id);
    return success(refreshed.record, 'Feeding fee payment voided');
  });
}
