/**
 * Stationery.js
 *
 * Handles stationery sales/payment records and fulfillment.
 *
 * Schema:
 * Transaction_ID, Student_ID, Item_ID, Quantity_Purchased, Unit_Price,
 * Total, Amount_Paid, Balance, Payment_Date, Payment_Method, Reference,
 * Fulfillment_Status, Quantity_Given, Quantity_Remaining, Given_By,
 * Given_Date, Recorded_By, Notes
 *
 * Transaction_ID is server-generated: ST-001, ST-002, ...
 *
 * Inventory is NOT reduced at purchase time. Stock is reduced when stationery
 * is fulfilled, because physical issuance is the inventory event.
 */

var STATIONERY_COLUMNS = [
  'Transaction_ID', 'Student_ID', 'Item_ID', 'Quantity_Purchased',
  'Unit_Price', 'Total', 'Amount_Paid', 'Balance', 'Payment_Date',
  'Payment_Method', 'Reference', 'Fulfillment_Status', 'Quantity_Given',
  'Quantity_Remaining', 'Given_By', 'Given_Date', 'Recorded_By', 'Notes'
];

var STATIONERY_ID_COLUMN = 'Transaction_ID';

var STATIONERY_PAYMENT_METHODS = [
  'Cash', 'Bank Transfer', 'Mobile Money', 'Other'
];

var STATIONERY_FULFILLMENT_STATUS = [
  'Pending', 'Partial', 'Fulfilled', 'Voided'
];

function getStationerySheet_() {
  try {
    return getSheet_(CONFIG.SHEETS.STATIONERY);
  } catch (err) {
    if (err && err.code === ERROR_CODES.NOT_FOUND) {
      throwError_(
        'Stationery sheet is missing.',
        ERROR_CODES.SERVER_ERROR,
        { sheet: CONFIG.SHEETS.STATIONERY, reason: 'sheet-missing' }
      );
    }
    throw err;
  }
}

function getStationeryHeaders_() {
  var sheet = getStationerySheet_();
  var headers = getHeaders_(sheet);

  if (headers.length === 0) {
    throwError_(
      'Stationery sheet has no headers.',
      ERROR_CODES.SERVER_ERROR,
      { sheet: CONFIG.SHEETS.STATIONERY }
    );
  }

  var missing = STATIONERY_COLUMNS.filter(function (column) {
    return headers.indexOf(column) === -1;
  });

  if (missing.length > 0) {
    throwError_(
      'Stationery sheet missing columns: ' + missing.join(', '),
      ERROR_CODES.SERVER_ERROR,
      {
        sheet: CONFIG.SHEETS.STATIONERY,
        missingColumns: missing
      }
    );
  }

  return headers;
}

function normalizeStationeryRecord_(record) {
  record.Transaction_ID = toTrimmedString_(record.Transaction_ID);
  record.Student_ID = toTrimmedString_(record.Student_ID);
  record.Item_ID = toTrimmedString_(record.Item_ID);
  record.Quantity_Purchased = Number(record.Quantity_Purchased);
  record.Unit_Price = Number(record.Unit_Price);
  record.Total = record.Quantity_Purchased * record.Unit_Price;
  record.Amount_Paid = Number(record.Amount_Paid);
  record.Balance = record.Total - record.Amount_Paid;
  record.Payment_Date = toTrimmedString_(record.Payment_Date);
  record.Payment_Method = toTrimmedString_(record.Payment_Method);
  record.Reference = toTrimmedString_(record.Reference);
  record.Fulfillment_Status = toTrimmedString_(record.Fulfillment_Status);
  record.Quantity_Given = Number(record.Quantity_Given || 0);
  record.Quantity_Remaining =
    record.Quantity_Purchased - record.Quantity_Given;
  record.Given_By = toTrimmedString_(record.Given_By);
  record.Given_Date = toTrimmedString_(record.Given_Date);
  record.Recorded_By = toTrimmedString_(record.Recorded_By);
  record.Notes = toTrimmedString_(record.Notes);
  return record;
}

function assertKnownStationeryFields_(headers, payload) {
  var unknown = Object.keys(payload).filter(function (key) {
    return headers.indexOf(key) === -1;
  });

  if (unknown.length > 0) {
    throwError_(
      'Unknown column(s) for sheet "' +
        CONFIG.SHEETS.STATIONERY +
        '": ' +
        unknown.join(', ') +
        '.',
      ERROR_CODES.VALIDATION_ERROR,
      {
        sheet: CONFIG.SHEETS.STATIONERY,
        unknownColumns: unknown,
        validColumns: headers
      }
    );
  }

  return payload;
}

function nextStationeryId_() {
  var sheet = getStationerySheet_();
  var headers = getHeaders_(sheet);
  var idIdx = headers.indexOf(STATIONERY_ID_COLUMN);
  var lastRow = sheet.getLastRow();
  var maxNum = 0;

  if (lastRow >= 2) {
    var values = sheet
      .getRange(2, idIdx + 1, lastRow - 1, 1)
      .getValues();

    for (var i = 0; i < values.length; i++) {
      var value = toTrimmedString_(values[i][0]);
      var match = value.match(/^ST-(\d+)$/);

      if (match) {
        var number = parseInt(match[1], 10);
        if (number > maxNum) maxNum = number;
      }
    }
  }

  return 'ST-' + String(maxNum + 1).padStart(3, '0');
}

function findStationeryById_(id) {
  var found = findRowById_(
    CONFIG.SHEETS.STATIONERY,
    id,
    STATIONERY_ID_COLUMN
  );

  if (found) {
    normalizeStationeryRecord_(found.record);
  }

  return found;
}

function readAllStationery_(headers) {
  var sheet = getStationerySheet_();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  var values = sheet
    .getRange(2, 1, lastRow - 1, headers.length)
    .getValues();

  return values.map(function (row) {
    return normalizeStationeryRecord_(
      rowToObject_(headers, row)
    );
  });
}

function validateStationeryStudent_(studentId) {
  var student = findStudentById_(studentId);

  if (!student) {
    throwError_(
      'Student not found: ' + studentId,
      ERROR_CODES.NOT_FOUND,
      { Student_ID: studentId }
    );
  }

  if (toTrimmedString_(student.record.Status) === 'Withdrawn') {
    throwError_(
      'Cannot create stationery transaction for withdrawn student: ' +
        studentId,
      ERROR_CODES.VALIDATION_ERROR,
      {
        Student_ID: studentId,
        reason: 'student-withdrawn'
      }
    );
  }

  return student;
}

function validateStationeryItem_(itemId) {
  var item = findInventoryById_(itemId);

  if (!item) {
    throwError_(
      'Inventory item not found: ' + itemId,
      ERROR_CODES.NOT_FOUND,
      { Item_ID: itemId }
    );
  }

  return item;
}

function handleStationeryList_(payload) {
  requirePermission_('STATIONERY.READ');

  var headers = getStationeryHeaders_();
  var records = readAllStationery_(headers);

  if (payload && payload.Student_ID) {
    var studentId = toTrimmedString_(payload.Student_ID);
    records = records.filter(function (record) {
      return record.Student_ID === studentId;
    });
  }

  if (payload && payload.Item_ID) {
    var itemId = toTrimmedString_(payload.Item_ID);
    records = records.filter(function (record) {
      return record.Item_ID === itemId;
    });
  }

  if (payload && payload.Fulfillment_Status) {
    var status = toTrimmedString_(payload.Fulfillment_Status);
    records = records.filter(function (record) {
      return record.Fulfillment_Status === status;
    });
  }

  return success(records, 'Stationery transactions retrieved');
}

function handleStationeryCreate_(payload) {
  requirePermission_('STATIONERY.CREATE');

  if (!payload || typeof payload !== 'object') {
    throwError_(
      'Payload is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { reason: 'missing-payload' }
    );
  }

  var headers = getStationeryHeaders_();

  assertKnownStationeryFields_(headers, payload);

  if (payload.Transaction_ID) {
    throwError_(
      'Transaction_ID cannot be set by client. It is server-generated.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Transaction_ID' }
    );
  }

  var studentId = toTrimmedString_(payload.Student_ID);
  var itemId = toTrimmedString_(payload.Item_ID);

  if (!studentId) {
    throwError_(
      'Student_ID is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Student_ID' }
    );
  }

  if (!itemId) {
    throwError_(
      'Item_ID is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Item_ID' }
    );
  }

  validateStationeryStudent_(studentId);
  var item = validateStationeryItem_(itemId);

  var quantity = Number(payload.Quantity_Purchased);

  if (isNaN(quantity) || quantity <= 0) {
    throwError_(
      'Quantity_Purchased must be greater than zero.',
      ERROR_CODES.VALIDATION_ERROR,
      {
        field: 'Quantity_Purchased',
        value: payload.Quantity_Purchased
      }
    );
  }

  var unitPrice = Number(item.record.Selling_Price);

  if (isNaN(unitPrice) || unitPrice < 0) {
    throwError_(
      'Inventory item has an invalid Selling_Price.',
      ERROR_CODES.SERVER_ERROR,
      {
        Item_ID: itemId,
        value: item.record.Selling_Price
      }
    );
  }

  var total = quantity * unitPrice;
  var amountPaid = Number(payload.Amount_Paid);

  if (isNaN(amountPaid) || amountPaid < 0) {
    throwError_(
      'Amount_Paid must be a non-negative number.',
      ERROR_CODES.VALIDATION_ERROR,
      {
        field: 'Amount_Paid',
        value: payload.Amount_Paid
      }
    );
  }

  if (amountPaid > total) {
    throwError_(
      'Amount_Paid (' +
        amountPaid +
        ') cannot exceed Total (' +
        total +
        ').',
      ERROR_CODES.VALIDATION_ERROR,
      {
        field: 'Amount_Paid',
        value: amountPaid,
        reason: 'amount-paid-exceeds-total'
      }
    );
  }

  var paymentMethod = toTrimmedString_(payload.Payment_Method);

  if (STATIONERY_PAYMENT_METHODS.indexOf(paymentMethod) === -1) {
    throwError_(
      'Invalid Payment_Method: ' + paymentMethod,
      ERROR_CODES.VALIDATION_ERROR,
      {
        field: 'Payment_Method',
        value: paymentMethod,
        allowedValues: STATIONERY_PAYMENT_METHODS
      }
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
      {
        field: 'Payment_Date',
        value: paymentDate
      }
    );
  }

  paymentDate = formatDate_(parsedDate);

  var user = requireAuthentication_();
  var recordedBy = user.staffId || user.userId || 'SYSTEM';

  var record = {
    Student_ID: studentId,
    Item_ID: itemId,
    Quantity_Purchased: quantity,
    Unit_Price: unitPrice,
    Total: total,
    Amount_Paid: amountPaid,
    Balance: total - amountPaid,
    Payment_Date: paymentDate,
    Payment_Method: paymentMethod,
    Reference: toTrimmedString_(payload.Reference || ''),
    Fulfillment_Status: 'Pending',
    Quantity_Given: 0,
    Quantity_Remaining: quantity,
    Given_By: '',
    Given_Date: '',
    Recorded_By: recordedBy,
    Notes: toTrimmedString_(payload.Notes || '')
  };

  return withScriptLock_(function () {
    var id = nextStationeryId_();
    record.Transaction_ID = id;

    if (findStationeryById_(id)) {
      throwError_(
        'Duplicate Transaction_ID generated: ' + id,
        ERROR_CODES.CONFLICT,
        { Transaction_ID: id }
      );
    }

    var created = appendRow_(
      CONFIG.SHEETS.STATIONERY,
      record
    );

    return success(
      created,
      'Stationery transaction created'
    );
  });
}

function handleStationeryFulfill_(payload) {
  requirePermission_('STATIONERY.FULFILL');

  if (!payload || typeof payload !== 'object') {
    throwError_(
      'Payload is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { reason: 'missing-payload' }
    );
  }

  var id = toTrimmedString_(payload.Transaction_ID);

  if (!id) {
    throwError_(
      'Transaction_ID is required for fulfillment.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Transaction_ID' }
    );
  }

  var quantityToGive = Number(payload.Quantity_Given);

  if (isNaN(quantityToGive) || quantityToGive <= 0) {
    throwError_(
      'Quantity_Given must be greater than zero.',
      ERROR_CODES.VALIDATION_ERROR,
      {
        field: 'Quantity_Given',
        value: payload.Quantity_Given
      }
    );
  }

  return withScriptLock_(function () {
    var result = findStationeryById_(id);

    if (!result) {
      throwError_(
        'Stationery transaction not found: ' + id,
        ERROR_CODES.NOT_FOUND,
        { Transaction_ID: id }
      );
    }

    var current = result.record;

    if (current.Fulfillment_Status === 'Voided') {
      throwError_(
        'Cannot fulfill a voided stationery transaction: ' + id,
        ERROR_CODES.VALIDATION_ERROR,
        {
          Transaction_ID: id,
          reason: 'voided-transaction'
        }
      );
    }

    var remaining = Number(current.Quantity_Remaining);

    if (quantityToGive > remaining) {
      throwError_(
        'Cannot fulfill ' +
          quantityToGive +
          ' item(s). Only ' +
          remaining +
          ' remain.',
        ERROR_CODES.VALIDATION_ERROR,
        {
          Transaction_ID: id,
          requested: quantityToGive,
          remaining: remaining
        }
      );
    }

    var item = validateStationeryItem_(current.Item_ID);
    var currentStock = Number(item.record.Current_Quantity);

    if (isNaN(currentStock) || currentStock < quantityToGive) {
      throwError_(
        'Insufficient inventory for item: ' + current.Item_ID,
        ERROR_CODES.VALIDATION_ERROR,
        {
          Item_ID: current.Item_ID,
          available: currentStock,
          requested: quantityToGive
        }
      );
    }

    var user = requireAuthentication_();
    var givenBy = user.staffId || user.userId || 'SYSTEM';
    var givenDate = formatDate_(new Date());

    var newGiven =
      Number(current.Quantity_Given) + quantityToGive;

    var newRemaining =
      Number(current.Quantity_Purchased) - newGiven;

    var newStatus =
      newRemaining <= 0 ? 'Fulfilled' : 'Partial';

    // Inventory.js provides the locked stock operation.
    var stockResult = adjustInventoryStockOut_(
      current.Item_ID,
      quantityToGive,
      'Stationery fulfillment ' + id,
      id
    );

    if (!stockResult || !stockResult.success) {
      throwError_(
        stockResult && stockResult.message
          ? stockResult.message
          : 'Inventory stock-out failed.',
        stockResult && stockResult.error
          ? stockResult.error.code
          : ERROR_CODES.SERVER_ERROR,
        stockResult && stockResult.error
          ? stockResult.error.details
          : null
      );
    }

    current.Quantity_Given = newGiven;
    current.Quantity_Remaining = newRemaining;
    current.Fulfillment_Status = newStatus;
    current.Given_By = givenBy;
    current.Given_Date = givenDate;

    var headers = getStationeryHeaders_();
    var values = recordToValues_(
      CONFIG.SHEETS.STATIONERY,
      headers,
      current
    );

    var sheet = getStationerySheet_();

    sheet
      .getRange(result.sheetRow, 1, 1, headers.length)
      .setValues([values]);

    var refreshed = findStationeryById_(id);

    return success(
      refreshed.record,
      'Stationery fulfillment recorded'
    );
  });
}
