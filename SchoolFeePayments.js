/**
 * SchoolFeePayments.js
 *
 * Payment TRANSACTIONS against School_Fees fee obligations (accounts).
 *
 * MODEL (Option 2):
 *   - One School_Fees row = one fee OBLIGATION for a student/year/term.
 *     Amount_Due there is the Fee Amount; Amount_Paid/Balance on that row are
 *     SERVER-DERIVED aggregates over the non-voided payment rows below.
 *   - One School_Fee_Payments row = one money-received event. A second payment
 *     against the same obligation NEVER creates a second obligation.
 *
 * Schema: Payment_ID, Fee_ID, Amount, Payment_Method, Payment_Date,
 *         Reference, Recorded_By, Status, Notes
 *
 * Payment_ID format: FFP-001, FFP-002, ... (server-generated, immutable)
 * Fee_ID references School_Fees.Payment_ID.
 * Status: 'Paid' (money received) or 'Voided' (correction; excluded from
 *         Total Paid but never deleted, so history is preserved).
 *
 * RULES ENFORCED HERE:
 *   - Amount must be > 0 and <= the obligation's current Outstanding
 *     (Fee Amount - non-voided payments). Overpayment is rejected for MVP1.
 *   - Every read-modify-write runs inside withScriptLock_().
 *   - Recorded_By is server-set from the authenticated user.
 *
 * PERMISSIONS: reuses the existing SCHOOL_FEES.* codes:
 *   feePayments.list   -> SCHOOL_FEES.READ
 *   feePayments.create -> SCHOOL_FEES.CREATE
 *   feePayments.void   -> SCHOOL_FEES.VOID
 */

var SCHOOL_FEE_PAYMENTS_COLUMNS = [
  'Payment_ID', 'Fee_ID', 'Amount', 'Payment_Method', 'Payment_Date',
  'Reference', 'Recorded_By', 'Status', 'Notes'
];

var SCHOOL_FEE_PAYMENTS_ID_COLUMN = 'Payment_ID';

var SCHOOL_FEE_PAYMENTS_CREATE_REQUIRED = [
  'Fee_ID', 'Amount', 'Payment_Method', 'Payment_Date'
];

/** A payment transaction is either money received or a voided correction. */
var SCHOOL_FEE_PAYMENT_STATUS_VALUES = ['Paid', 'Voided'];

function getFeePaymentsSheet_() {
  try {
    return getSheet_(CONFIG.SHEETS.SCHOOL_FEE_PAYMENTS);
  } catch (err) {
    if (err && err.code === ERROR_CODES.NOT_FOUND) {
      throwError_(
        'School_Fee_Payments sheet is missing.',
        ERROR_CODES.SERVER_ERROR,
        { sheet: CONFIG.SHEETS.SCHOOL_FEE_PAYMENTS, reason: 'sheet-missing' }
      );
    }
    throw err;
  }
}

function getFeePaymentsHeaders_() {
  var sheet = getFeePaymentsSheet_();
  var headers = getHeaders_(sheet);
  if (headers.length === 0) {
    throwError_(
      'School_Fee_Payments sheet has no headers.',
      ERROR_CODES.SERVER_ERROR,
      { sheet: CONFIG.SHEETS.SCHOOL_FEE_PAYMENTS }
    );
  }
  var missing = SCHOOL_FEE_PAYMENTS_COLUMNS.filter(function (c) {
    return headers.indexOf(c) === -1;
  });
  if (missing.length > 0) {
    throwError_(
      'School_Fee_Payments sheet missing columns: ' + missing.join(', '),
      ERROR_CODES.SERVER_ERROR,
      { sheet: CONFIG.SHEETS.SCHOOL_FEE_PAYMENTS, missingColumns: missing }
    );
  }
  return headers;
}

function normalizeFeePaymentRecord_(record) {
  record.Amount = Number(record.Amount) || 0;
  record.Fee_ID = toTrimmedString_(record.Fee_ID);
  record.Payment_Method = toTrimmedString_(record.Payment_Method);
  record.Payment_Date = toTrimmedString_(record.Payment_Date);
  record.Reference = toTrimmedString_(record.Reference);
  record.Status = toTrimmedString_(record.Status);
  record.Recorded_By = toTrimmedString_(record.Recorded_By);
  record.Notes = toTrimmedString_(record.Notes);
  return record;
}

/** Reject payload keys that are not School_Fee_Payments columns. */
function assertKnownFeePaymentFields_(headers, payload) {
  var unknown = [];
  Object.keys(payload).forEach(function (key) {
    if (key === '__auth') return;
    if (headers.indexOf(key) === -1) unknown.push(key);
  });
  if (unknown.length > 0) {
    throwError_(
      'Unknown column(s) for sheet "' + CONFIG.SHEETS.SCHOOL_FEE_PAYMENTS + '": ' + unknown.join(', ') + '.',
      ERROR_CODES.VALIDATION_ERROR,
      { sheet: CONFIG.SHEETS.SCHOOL_FEE_PAYMENTS, unknownColumns: unknown, validColumns: headers }
    );
  }
  return payload;
}

function findFeePaymentById_(id) {
  var found = findRowById_(
    CONFIG.SHEETS.SCHOOL_FEE_PAYMENTS,
    id,
    SCHOOL_FEE_PAYMENTS_ID_COLUMN
  );
  if (found) normalizeFeePaymentRecord_(found.record);
  return found;
}

/**
 * All payment rows for one fee obligation (voided included; callers filter).
 */
function readAllFeePaymentsForFee_(feeId, headers) {
  var sheet = getFeePaymentsSheet_();
  var lastRow = sheet.getLastRow();
  var want = toTrimmedString_(feeId);
  var out = [];
  if (lastRow < 2 || headers.length === 0) return out;
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  for (var i = 0; i < values.length; i++) {
    var rec = normalizeFeePaymentRecord_(rowToObject_(headers, values[i]));
    if (rec.Fee_ID === want) out.push(rec);
  }
  return out;
}

/**
 * Total Paid for an obligation = sum of NON-VOIDED payment amounts.
 * This is the only authoritative source for the account's paid figure.
 */
function computeFeeTotalPaid_(feeId) {
  // The payments sheet may not exist yet (fresh deployments before the tab is
  // created). An absent ledger simply means nothing has been paid.
  try {
    var headers = getFeePaymentsHeaders_();
  } catch (err) {
    if (err && err.details && err.details.sheet === CONFIG.SHEETS.SCHOOL_FEE_PAYMENTS) {
      return 0;
    }
    throw err;
  }
  var payments = readAllFeePaymentsForFee_(feeId, headers);
  var total = 0;
  for (var i = 0; i < payments.length; i++) {
    if (payments[i].Status !== 'Voided') total += payments[i].Amount;
  }
  return total;
}

/**
 * Re-derive the obligation's aggregate fields from its payment rows and
 * persist them on the School_Fees row (Amount_Paid/Balance are compatibility
 * mirrors of the payment ledger; the ledger is authoritative).
 * Must be called INSIDE withScriptLock_ after any payment change.
 */
function refreshFeeAccountTotals_(feeId) {
  var totalPaid = computeFeeTotalPaid_(feeId);
  var result = findSchoolFeeById_(feeId);
  if (!result) return null;
  var record = result.record;
  record.Amount_Paid = totalPaid;
  record.Balance = record.Amount_Due - totalPaid;
  record.Status = deriveSchoolFeesStatus_(record.Amount_Due, totalPaid);
  var sheet = getSchoolFeesSheet_();
  var headers = getSchoolFeesHeaders_();
  sheet.getRange(result.sheetRow, 1, 1, headers.length).setValues([
    recordToValues_(CONFIG.SHEETS.SCHOOL_FEES, headers, record)
  ]);
  return findSchoolFeeById_(feeId).record;
}

function nextFeePaymentId_() {
  var sheet = getFeePaymentsSheet_();
  var headers = getHeaders_(sheet);
  var idIdx = headers.indexOf(SCHOOL_FEE_PAYMENTS_ID_COLUMN);
  var lastRow = sheet.getLastRow();
  var maxNum = 0;
  if (lastRow >= 2 && idIdx !== -1) {
    var values = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      var v = toTrimmedString_(values[i][0]);
      var m = v.match(/^FFP-(\d+)$/);
      if (m) {
        var n = parseInt(m[1], 10);
        if (n > maxNum) maxNum = n;
      }
    }
  }
  return 'FFP-' + String(maxNum + 1).padStart(3, '0');
}

/* ==========================================================================
 * Handlers
 * ======================================================================== */

/**
 * feePayments.list
 * Optional filter: Fee_ID. Each returned record is one payment transaction.
 * With Fee_ID, the response also carries the account's server-derived
 * totals so the UI never computes authoritative balances itself.
 */
function handleFeePaymentsList_(payload) {
  requirePermission_('SCHOOL_FEES.READ');
  var headers = getFeePaymentsHeaders_();
  var payments = [];
  var sheet = getFeePaymentsSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (var i = 0; i < values.length; i++) {
      payments.push(normalizeFeePaymentRecord_(rowToObject_(headers, values[i])));
    }
  }

  var feeId = payload ? toTrimmedString_(payload.Fee_ID) : '';
  if (feeId) {
    payments = payments.filter(function (p) { return p.Fee_ID === feeId; });

    var account = findSchoolFeeById_(feeId);
    if (!account) {
      throwError_(
        'Fee account not found: ' + feeId,
        ERROR_CODES.NOT_FOUND,
        { Fee_ID: feeId }
      );
    }
    var rec = account.record;
    var totalPaid = computeFeeTotalPaid_(feeId);
    return success(
      {
        payments: payments,
        account: {
          Fee_ID: rec.Payment_ID,
          Student_ID: rec.Student_ID,
          Academic_Year: rec.Academic_Year,
          Term: rec.Term,
          Fee_Amount: rec.Amount_Due,
          Total_Paid: totalPaid,
          Outstanding: rec.Amount_Due - totalPaid,
          Status: deriveSchoolFeesStatus_(rec.Amount_Due, totalPaid)
        }
      },
      'Fee payments retrieved'
    );
  }

  return success(payments, 'Fee payments retrieved');
}

/**
 * feePayments.create
 *
 * Records ONE money-received event against an existing fee obligation.
 * Server-side inside the script lock:
 *   1. load the obligation (reject unknown / voided),
 *   2. Total Paid = sum of non-voided payments for that Fee_ID,
 *   3. Outstanding = Fee Amount - Total Paid,
 *   4. reject Amount <= 0 or Amount > Outstanding,
 *   5. append exactly one School_Fee_Payments row,
 *   6. recalculate and persist the account aggregates,
 *   7. return { payment, account }.
 */
function handleFeePaymentsCreate_(payload) {
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
  assertRequired_(payload, SCHOOL_FEE_PAYMENTS_CREATE_REQUIRED);

  return withScriptLock_(function () {
    var feeId = toTrimmedString_(payload.Fee_ID);
    var fee = findSchoolFeeById_(feeId);
    if (!fee) {
      throwError_(
        'Fee account not found: ' + feeId,
        ERROR_CODES.NOT_FOUND,
        { Fee_ID: feeId }
      );
    }
    if (toTrimmedString_(fee.record.Status) === 'Voided') {
      throwError_(
        'Cannot record a payment against a voided fee account: ' + feeId,
        ERROR_CODES.VALIDATION_ERROR,
        { Fee_ID: feeId, reason: 'voided-fee-account' }
      );
    }

    var amount = Number(payload.Amount);
    if (isNaN(amount) || amount <= 0) {
      throwError_(
        'Amount must be a positive number.',
        ERROR_CODES.VALIDATION_ERROR,
        { field: 'Amount', value: payload.Amount }
      );
    }

    var totalPaid = computeFeeTotalPaid_(feeId);
    var outstanding = fee.record.Amount_Due - totalPaid;
    if (amount > outstanding) {
      throwError_(
        'Payment Amount (' + amount + ') cannot exceed the outstanding balance (' + outstanding + ').',
        ERROR_CODES.VALIDATION_ERROR,
        { field: 'Amount', value: amount, outstanding: outstanding, reason: 'payment-exceeds-outstanding' }
      );
    }

    var paymentMethod = toTrimmedString_(payload.Payment_Method);
    if (SCHOOL_FEES_PAYMENT_METHODS.indexOf(paymentMethod) === -1) {
      throwError_(
        'Invalid Payment_Method: ' + paymentMethod,
        ERROR_CODES.VALIDATION_ERROR,
        { field: 'Payment_Method', value: paymentMethod, allowedValues: SCHOOL_FEES_PAYMENT_METHODS }
      );
    }

    var paymentDate = toTrimmedString_(payload.Payment_Date);
    var parsedDate = toDate_(paymentDate);
    if (!parsedDate) {
      throwError_(
        'Invalid Payment_Date: ' + paymentDate,
        ERROR_CODES.VALIDATION_ERROR,
        { field: 'Payment_Date', value: paymentDate }
      );
    }
    paymentDate = formatDate_(parsedDate);

    var user = requireAuthentication_();
    var recordedBy = user.staffId || user.userId || 'SYSTEM';

    var headers = getFeePaymentsHeaders_();
    assertKnownFeePaymentFields_(headers, payload);

    var payment = {
      Fee_ID: feeId,
      Amount: amount,
      Payment_Method: paymentMethod,
      Payment_Date: paymentDate,
      Reference: toTrimmedString_(payload.Reference || ''),
      Status: 'Paid',
      Recorded_By: recordedBy,
      Notes: toTrimmedString_(payload.Notes || '')
    };
    payment.Payment_ID = nextFeePaymentId_();

    if (findFeePaymentById_(payment.Payment_ID)) {
      throwError_(
        'Duplicate Payment_ID generated: ' + payment.Payment_ID,
        ERROR_CODES.CONFLICT,
        { Payment_ID: payment.Payment_ID }
      );
    }

    var created = appendRow_(CONFIG.SHEETS.SCHOOL_FEE_PAYMENTS, payment);

    // Step 6: recompute the obligation's aggregates from the ledger.
    var record = refreshFeeAccountTotals_(feeId);
    var account = Object.assign({}, record, {
      Total_Paid: record.Amount_Paid,
      Outstanding: record.Balance
    });

    return success(
      { payment: created, account: account },
      'Fee payment recorded'
    );
  });
}

/**
 * feePayments.void
 *
 * Soft correction: the payment row is NEVER deleted. Status becomes 'Voided'
 * and the row is excluded from Total Paid, so:
 *   Outstanding = Fee Amount - (non-voided payments)
 * The financial history stays intact.
 */
function handleFeePaymentsVoid_(payload) {
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

  return withScriptLock_(function () {
    var result = findFeePaymentById_(id);
    if (!result) {
      throwError_(
        'Fee payment not found: ' + id,
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
    var sheet = getFeePaymentsSheet_();
    var headers = getFeePaymentsHeaders_();
    sheet.getRange(result.sheetRow, 1, 1, headers.length).setValues([
      recordToValues_(CONFIG.SHEETS.SCHOOL_FEE_PAYMENTS, headers, updated)
    ]);

    // Recompute the obligation's aggregates now that this payment no longer
    // counts toward Total Paid.
    var record = refreshFeeAccountTotals_(updated.Fee_ID);
    var account = Object.assign({}, record, {
      Total_Paid: record.Amount_Paid,
      Outstanding: record.Balance
    });

    return success(
      { payment: updated, account: account },
      'Fee payment voided'
    );
  });
}
