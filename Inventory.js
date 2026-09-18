/**
 * Inventory.js
 *
 * Phase 5: stock items and the stock movement ledger.
 *
 * SCHEMA
 *   Inventory (CONFIG.SHEETS.INVENTORY):
 *     Item_ID, Item_Name, Category, Unit, Selling_Price, Current_Quantity,
 *     Minimum_Stock_Level, Status
 *   Inventory_Movements (CONFIG.SHEETS.INVENTORY_MOVEMENTS):
 *     Movement_ID, Item_ID, Movement_Type, Quantity, Date, Reason,
 *     Recorded_By, Notes
 *
 * IDS
 *   Item_ID     server-generated and immutable: ITM-001, ITM-002, ...
 *   Movement_ID server-generated and append-only: MOV-001, MOV-002, ...
 *   A client-supplied Item_ID / Movement_ID is REJECTED (VALIDATION_ERROR),
 *   never trusted, exactly like Payment_ID in the Phase 4B fee modules.
 *
 * SERVER-CONTROLLED COLUMNS
 *   Status        DERIVED from Current_Quantity and Minimum_Stock_Level on
 *                 every read and every write, so it can never drift from the
 *                 quantity. Accepting a client value would let the sheet claim
 *                 'In Stock' while the quantity says otherwise.
 *   Date          the server clock (script timezone), not the client's.
 *   Recorded_By   the authenticated user (Staff_ID, else User_ID), not the
 *                 client's.
 *   Unknown payload fields are rejected with VALIDATION_ERROR and
 *   details.unknownColumns + details.validColumns, matching the Phase 3 and
 *   Phase 4B convention.
 *
 * STOCK RULES
 *   stockIn  increases Current_Quantity.
 *   stockOut decreases Current_Quantity and can NEVER drive it below zero.
 *   Every quantity change appends exactly one Inventory_Movements row INSIDE the
 *   same critical section as the item update, so the ledger and the item row can
 *   never disagree. The ledger tab is validated before the item row is touched,
 *   so a stock change can never be applied without its movement row. Stock on
 *   hand is therefore reconstructable from the ledger.
 *   Movement_Type distinguishes the direction unambiguously as the machine
 *   tokens 'STOCK_IN' / 'STOCK_OUT' (Config.js: machine tokens are UPPER_SNAKE).
 *
 * LOCKING
 *   create / stockIn / stockOut run the whole read -> validate -> write ->
 *   movement sequence inside withScriptLock_() (Utils.js). There is no second
 *   locking mechanism: the one script lock is reused.
 *
 *   Stationery.js calls adjustInventoryStockOut_() from INSIDE its own
 *   withScriptLock_ section (handleStationeryFulfill_). That helper therefore
 *   runs through withInventoryLock_(), which re-enters the critical section only
 *   when this execution does NOT already hold the script lock (Lock.hasLock()).
 *   Standalone calls stay protected; the nested call from stationery.fulfill
 *   can neither deadlock nor wait on itself.
 *
 * Audit_Log writes are deliberately absent: that is Phase 7.
 */

var INVENTORY_COLUMNS = [
  'Item_ID', 'Item_Name', 'Category', 'Unit', 'Selling_Price',
  'Current_Quantity', 'Minimum_Stock_Level', 'Status'
];

var INVENTORY_MOVEMENT_COLUMNS = [
  'Movement_ID', 'Item_ID', 'Movement_Type', 'Quantity', 'Date', 'Reason',
  'Recorded_By', 'Notes'
];

var INVENTORY_ID_COLUMN = 'Item_ID';
var INVENTORY_MOVEMENT_ID_COLUMN = 'Movement_ID';

var INVENTORY_ID_PREFIX = 'ITM';
var INVENTORY_MOVEMENT_ID_PREFIX = 'MOV';

/** Movement_Type stored values. Machine tokens, so UPPER_SNAKE. */
var INVENTORY_MOVEMENT_TYPES = {
  STOCK_IN: 'STOCK_IN',
  STOCK_OUT: 'STOCK_OUT'
};

/** Status vocabulary produced by deriveInventoryStatus_. Stored Title Case. */
var INVENTORY_STATUS_VALUES = {
  IN_STOCK: 'In Stock',
  LOW_STOCK: 'Low Stock',
  OUT_OF_STOCK: 'Out of Stock'
};

/**
 * Fields inventory.create accepts. 'Status' is accepted and then derived, and
 * Item_ID is deliberately absent: a client may never set its own ID.
 */
var INVENTORY_CREATE_ALLOWED = [
  'Item_Name', 'Category', 'Unit', 'Selling_Price', 'Current_Quantity',
  'Minimum_Stock_Level', 'Status'
];

/** Fields inventory.stockIn / inventory.stockOut accept. */
var INVENTORY_STOCK_ALLOWED = ['Item_ID', 'Quantity', 'Reason', 'Notes'];

/** Filters inventory.list accepts. */
var INVENTORY_LIST_FILTERS = ['Item_ID', 'Category', 'Status'];

/** Filters inventory.movements accepts. */
var INVENTORY_MOVEMENT_FILTERS = ['Item_ID', 'Movement_Type', 'Date'];

function getInventorySheet_() {
  try {
    return getSheet_(CONFIG.SHEETS.INVENTORY);
  } catch (err) {
    if (err && err.code === ERROR_CODES.NOT_FOUND) {
      throwError_(
        'Inventory sheet is missing.',
        ERROR_CODES.SERVER_ERROR,
        { sheet: CONFIG.SHEETS.INVENTORY, reason: 'sheet-missing' }
      );
    }
    throw err;
  }
}

function getInventoryHeaders_() {
  var sheet = getInventorySheet_();
  var headers = getHeaders_(sheet);

  if (headers.length === 0) {
    throwError_(
      'Inventory sheet has no headers.',
      ERROR_CODES.SERVER_ERROR,
      { sheet: CONFIG.SHEETS.INVENTORY }
    );
  }

  var missing = INVENTORY_COLUMNS.filter(function (column) {
    return headers.indexOf(column) === -1;
  });

  if (missing.length > 0) {
    throwError_(
      'Inventory sheet missing columns: ' + missing.join(', '),
      ERROR_CODES.SERVER_ERROR,
      { sheet: CONFIG.SHEETS.INVENTORY, missingColumns: missing }
    );
  }

  return headers;
}

function getMovementsSheet_() {
  try {
    return getSheet_(CONFIG.SHEETS.INVENTORY_MOVEMENTS);
  } catch (err) {
    if (err && err.code === ERROR_CODES.NOT_FOUND) {
      throwError_(
        'Inventory_Movements sheet is missing.',
        ERROR_CODES.SERVER_ERROR,
        { sheet: CONFIG.SHEETS.INVENTORY_MOVEMENTS, reason: 'sheet-missing' }
      );
    }
    throw err;
  }
}

function getMovementsHeaders_() {
  var sheet = getMovementsSheet_();
  var headers = getHeaders_(sheet);

  if (headers.length === 0) {
    throwError_(
      'Inventory_Movements sheet has no headers.',
      ERROR_CODES.SERVER_ERROR,
      { sheet: CONFIG.SHEETS.INVENTORY_MOVEMENTS }
    );
  }

  var missing = INVENTORY_MOVEMENT_COLUMNS.filter(function (column) {
    return headers.indexOf(column) === -1;
  });

  if (missing.length > 0) {
    throwError_(
      'Inventory_Movements sheet missing columns: ' + missing.join(', '),
      ERROR_CODES.SERVER_ERROR,
      { sheet: CONFIG.SHEETS.INVENTORY_MOVEMENTS, missingColumns: missing }
    );
  }

  return headers;
}

/**
 * Derive Status from the quantity, so Status can never contradict the numbers.
 *
 *   quantity <= 0                        -> 'Out of Stock'
 *   quantity <= Minimum_Stock_Level      -> 'Low Stock'
 *   otherwise                            -> 'In Stock'
 *
 * A blank or 0 minimum means "no reorder threshold", so only 0 is low.
 *
 * @param {number} quantity Current_Quantity.
 * @param {number} minimum Minimum_Stock_Level.
 * @return {string} One of INVENTORY_STATUS_VALUES.
 */
function deriveInventoryStatus_(quantity, minimum) {
  var quantityNumber = Number(quantity);
  var minimumNumber = Number(minimum);

  if (isNaN(quantityNumber)) quantityNumber = 0;
  if (isNaN(minimumNumber)) minimumNumber = 0;

  if (quantityNumber <= 0) return INVENTORY_STATUS_VALUES.OUT_OF_STOCK;
  if (quantityNumber <= minimumNumber) return INVENTORY_STATUS_VALUES.LOW_STOCK;
  return INVENTORY_STATUS_VALUES.IN_STOCK;
}

/**
 * Normalise a raw Inventory row the same way Students / Staff / School_Fees /
 * Feeding_Fees do: strings trimmed, numbers coerced, and derived columns
 * recalculated (here Status) instead of trusted from the sheet.
 */
function normalizeInventoryRecord_(record) {
  record.Item_ID = toTrimmedString_(record.Item_ID);
  record.Item_Name = toTrimmedString_(record.Item_Name);
  record.Category = toTrimmedString_(record.Category);
  record.Unit = toTrimmedString_(record.Unit);
  record.Selling_Price = Number(record.Selling_Price);
  record.Current_Quantity = Number(record.Current_Quantity);
  record.Minimum_Stock_Level = Number(record.Minimum_Stock_Level);
  record.Status = deriveInventoryStatus_(
    record.Current_Quantity,
    record.Minimum_Stock_Level
  );
  return record;
}

/** Normalise a raw Inventory_Movements row; Quantity is always numeric. */
function normalizeMovementRecord_(record) {
  record.Movement_ID = toTrimmedString_(record.Movement_ID);
  record.Item_ID = toTrimmedString_(record.Item_ID);
  record.Movement_Type = toTrimmedString_(record.Movement_Type);
  record.Quantity = Number(record.Quantity);
  record.Date = toTrimmedString_(record.Date);
  record.Reason = toTrimmedString_(record.Reason);
  record.Recorded_By = toTrimmedString_(record.Recorded_By);
  record.Notes = toTrimmedString_(record.Notes);
  return record;
}
/**
 * Reject payload keys that are not part of `allowed`.
 *
 * Same contract as assertKnownSchoolFeesFields_ / assertKnownStationeryFields_:
 * VALIDATION_ERROR with details.unknownColumns and details.validColumns, so a
 * typo can never be silently ignored. `sheet` names the tab the accepted fields
 * belong to; `action` names the action, because create and the stock actions
 * accept different subsets of those columns.
 *
 * @param {string} action Action name for the error details.
 * @param {string} sheetName The sheet the accepted fields belong to.
 * @param {Array<string>} allowed Allowed header names.
 * @param {Object} payload The caller payload.
 * @throws {Error} VALIDATION_ERROR listing the unknownColumns.
 */
function assertKnownInventoryFields_(action, sheetName, allowed, payload) {
  var unknown = Object.keys(payload).filter(function (key) {
    return allowed.indexOf(key) === -1;
  });

  if (unknown.length > 0) {
    throwError_(
      'Unknown column(s) for sheet "' + sheetName + '": ' + unknown.join(', ') + '.',
      ERROR_CODES.VALIDATION_ERROR,
      {
        action: action,
        sheet: sheetName,
        unknownColumns: unknown,
        validColumns: allowed.slice()
      }
    );
  }
}

/**
 * Coerce a payload value to a number, rejecting blanks, non-numeric strings,
 * NaN, Infinity and null.
 *
 * @param {*} value Raw payload value.
 * @param {string} field Field name for the error details.
 * @return {number} The numeric value.
 */
function parseInventoryNumber_(value, field) {
  if (value === null || value === undefined || typeof value === 'boolean') {
    throwError_(
      field + ' must be a number.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: field, value: value }
    );
  }

  if (typeof value === 'string' && toTrimmedString_(value) === '') {
    throwError_(
      field + ' must be a number.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: field, value: value }
    );
  }

  var number = Number(value);

  if (!isFinite(number)) {
    throwError_(
      field + ' must be a number.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: field, value: value }
    );
  }

  return number;
}

/**
 * Parse and validate a money / quantity / threshold field.
 *
 * @param {*} value Raw payload value.
 * @param {string} field Field name.
 * @param {Object=} opts `defaultValue` used when the value is null/undefined,
 *     `integer` requires a whole number, `positive` requires > 0 (default is
 *     >= 0). A defaulted field skips the range check and is returned as is.
 * @return {number} The accepted number.
 */
function parseInventoryAmount_(value, field, opts) {
  var options = opts || {};

  if (value === null || value === undefined) {
    if (options.defaultValue !== undefined) return options.defaultValue;
    throwError_(
      field + ' is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: field }
    );
  }

  var number = parseInventoryNumber_(value, field);

  if (options.integer && Math.floor(number) !== number) {
    throwError_(
      field + ' must be a whole number.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: field, value: value }
    );
  }

  if (options.positive) {
    if (number <= 0) {
      throwError_(
        field + ' must be greater than zero.',
        ERROR_CODES.VALIDATION_ERROR,
        { field: field, value: value }
      );
    }
  } else if (number < 0) {
    throwError_(
      field + ' cannot be negative.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: field, value: value }
    );
  }

  return number;
}
/**
 * Validate the caller's stock payload for stockIn / stockOut and return the
 * values the locked section will use. The permission check happens here, so a
 * caller is always authorized before anything is read.
 *
 * @param {Object} payload Caller payload.
 * @param {string} action 'inventory.stockIn' or 'inventory.stockOut'.
 * @param {string} permission Permission code (INVENTORY.ADJUST).
 * @return {{itemId: string, quantity: number, reason: string, notes: string}}
 */
function parseInventoryStockRequest_(payload, action, permission) {
  requirePermission_(permission);
  assertKnownInventoryFields_(
    action,
    CONFIG.SHEETS.INVENTORY_MOVEMENTS,
    INVENTORY_STOCK_ALLOWED,
    payload
  );

  var itemId = toTrimmedString_(payload.Item_ID);
  var itemIdSupplied = payload.Item_ID !== undefined &&
    payload.Item_ID !== null &&
    !isBlank_(payload.Item_ID);

  if (itemIdSupplied && itemId === '') {
    throwError_(
      'Item_ID must be a non-empty string.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Item_ID', value: payload.Item_ID }
    );
  }

  if (itemId === '') {
    throwError_(
      'Item_ID is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Item_ID' }
    );
  }

  if (Object.prototype.hasOwnProperty.call(payload, 'Item_ID') &&
      payload.Item_ID !== null &&
      typeof payload.Item_ID !== 'string') {
    throwError_(
      'Item_ID must be a string.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Item_ID', value: payload.Item_ID }
    );
  }

  if (payload.Quantity === undefined || payload.Quantity === null) {
    throwError_(
      'Quantity is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Quantity' }
    );
  }

  var quantity = parseInventoryAmount_(payload.Quantity, 'Quantity', {
    positive: true,
    integer: true
  });

  return {
    itemId: itemId,
    quantity: quantity,
    reason: toTrimmedString_(payload.Reason || ''),
    notes: toTrimmedString_(payload.Notes || '')
  };
}

/** The actor recorded in Recorded_By: Staff_ID, else User_ID. Never the client. */
function inventoryActor_() {
  var user = requireAuthentication_();
  return user.staffId || user.userId || 'SYSTEM';
}
/**
 * Next server-generated Item_ID: ITM-001, ITM-002, ... (max + 1, mirroring
 * nextStationeryId_ / nextFeedingFeesId_). Called inside the lock so two
 * callers can never be handed the same ID.
 *
 * @return {string} The next Item_ID.
 */
function nextInventoryItemId_() {
  var sheet = getInventorySheet_();
  var headers = getHeaders_(sheet);
  var idIdx = headers.indexOf(INVENTORY_ID_COLUMN);
  var lastRow = sheet.getLastRow();
  var maxNum = 0;

  if (lastRow >= 2) {
    var values = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues();

    for (var i = 0; i < values.length; i++) {
      var match = toTrimmedString_(values[i][0]).match(/^ITM-(\d+)$/);
      if (match) {
        var number = parseInt(match[1], 10);
        if (number > maxNum) maxNum = number;
      }
    }
  }

  return INVENTORY_ID_PREFIX + '-' + String(maxNum + 1).padStart(3, '0');
}

/**
 * Next server-generated Movement_ID: MOV-001, MOV-002, ... (max + 1). Called
 * inside the lock, which is what keeps the ledger's IDs collision-free.
 */
function nextInventoryMovementId_() {
  var sheet = getMovementsSheet_();
  var headers = getHeaders_(sheet);
  var idIdx = headers.indexOf(INVENTORY_MOVEMENT_ID_COLUMN);
  var lastRow = sheet.getLastRow();
  var maxNum = 0;

  if (lastRow >= 2) {
    var values = sheet.getRange(2, idIdx + 1, lastRow - 1, 1).getValues();

    for (var i = 0; i < values.length; i++) {
      var match = toTrimmedString_(values[i][0]).match(/^MOV-(\d+)$/);
      if (match) {
        var number = parseInt(match[1], 10);
        if (number > maxNum) maxNum = number;
      }
    }
  }

  return INVENTORY_MOVEMENT_ID_PREFIX + '-' + String(maxNum + 1).padStart(3, '0');
}

/**
 * Find an Inventory item by Item_ID.
 *
 * @param {string} itemId Item_ID to match.
 * @return {{sheetRow: number, record: Object}|null} Null when not found.
 * @throws {Error} SERVER_ERROR when the Inventory tab is absent.
 */
function findInventoryById_(itemId) {
  getInventorySheet_();

  var found = findRowById_(
    CONFIG.SHEETS.INVENTORY,
    itemId,
    INVENTORY_ID_COLUMN
  );

  if (found) normalizeInventoryRecord_(found.record);

  return found;
}

/**
 * Find a single Inventory_Movements row by Movement_ID.
 *
 * @param {string} movementId Movement_ID to match.
 * @return {{sheetRow: number, record: Object}|null} Null when not found.
 */
function findMovementById_(movementId) {
  getMovementsSheet_();

  var found = findRowById_(
    CONFIG.SHEETS.INVENTORY_MOVEMENTS,
    movementId,
    INVENTORY_MOVEMENT_ID_COLUMN
  );

  if (found) normalizeMovementRecord_(found.record);

  return found;
}

/** Read every Inventory item, normalised (Status derived). */
function readAllInventory_(headers) {
  var sheet = getInventorySheet_();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  return values.map(function (row) {
    return normalizeInventoryRecord_(rowToObject_(headers, row));
  });
}

/** Read every Inventory_Movements row, normalised, in sheet (chronological) order. */
function readAllMovements_(headers) {
  var sheet = getMovementsSheet_();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  return values.map(function (row) {
    return normalizeMovementRecord_(rowToObject_(headers, row));
  });
}
/**
 * inventory.list
 *
 * Requires INVENTORY.READ. Optional exact-match filters: Item_ID, Category,
 * Status (the derived status). An empty result is an empty array, never an
 * error.
 */
function handleInventoryList_(payload) {
  requirePermission_('INVENTORY.READ');

  var filters = payload || {};
  assertKnownInventoryFields_(
    'inventory.list',
    CONFIG.SHEETS.INVENTORY,
    INVENTORY_LIST_FILTERS,
    filters
  );

  var headers = getInventoryHeaders_();
  var records = readAllInventory_(headers);

  if (!isBlank_(filters.Item_ID)) {
    var itemId = toTrimmedString_(filters.Item_ID);
    records = records.filter(function (record) {
      return record.Item_ID === itemId;
    });
  }

  if (!isBlank_(filters.Category)) {
    var category = toTrimmedString_(filters.Category);
    records = records.filter(function (record) {
      return record.Category === category;
    });
  }

  if (!isBlank_(filters.Status)) {
    var status = toTrimmedString_(filters.Status);
    records = records.filter(function (record) {
      return record.Status === status;
    });
  }

  return success(records, 'Inventory items retrieved');
}

/**
 * inventory.movements
 *
 * Requires INVENTORY.READ. Optional exact-match filters: Item_ID,
 * Movement_Type (STOCK_IN / STOCK_OUT) and Date (YYYY-MM-DD). Newest movement
 * is last, because the ledger is returned in sheet order.
 */
function handleInventoryMovements_(payload) {
  var filters = payload || {};
  assertKnownInventoryFields_(
    'inventory.movements',
    CONFIG.SHEETS.INVENTORY_MOVEMENTS,
    INVENTORY_MOVEMENT_FILTERS,
    filters
  );

  if (!isBlank_(filters.Movement_Type)) {
    var rawType = toTrimmedString_(filters.Movement_Type).toUpperCase();
    if (rawType !== INVENTORY_MOVEMENT_TYPES.STOCK_IN &&
        rawType !== INVENTORY_MOVEMENT_TYPES.STOCK_OUT) {
      throwError_(
        'Invalid Movement_Type: ' + filters.Movement_Type,
        ERROR_CODES.VALIDATION_ERROR,
        {
          field: 'Movement_Type',
          value: filters.Movement_Type,
          validValues: [
            INVENTORY_MOVEMENT_TYPES.STOCK_IN,
            INVENTORY_MOVEMENT_TYPES.STOCK_OUT
          ]
        }
      );
    }
  }

  requirePermission_('INVENTORY.READ');

  var headers = getMovementsHeaders_();
  var records = readAllMovements_(headers);

  if (!isBlank_(filters.Item_ID)) {
    var itemId = toTrimmedString_(filters.Item_ID);
    records = records.filter(function (record) {
      return record.Item_ID === itemId;
    });
  }

  if (!isBlank_(filters.Movement_Type)) {
    var movementType = toTrimmedString_(filters.Movement_Type).toUpperCase();
    records = records.filter(function (record) {
      return record.Movement_Type === movementType;
    });
  }

  if (!isBlank_(filters.Date)) {
    var date = toTrimmedString_(filters.Date);
    records = records.filter(function (record) {
      return record.Date === date;
    });
  }

  return success(records, 'Inventory movements retrieved');
}
/**
 * inventory.create
 *
 * Requires INVENTORY.CREATE. Item_ID is server-generated (ITM-nnn) and is not an
 * accepted field. Current_Quantity and Minimum_Stock_Level default to 0 when
 * omitted. Status is derived, never taken from the caller.
 */
function handleInventoryCreate_(payload) {
  requirePermission_('INVENTORY.CREATE');

  if (!payload || typeof payload !== 'object') {
    throwError_(
      'Payload is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { reason: 'missing-payload' }
    );
  }

  assertKnownInventoryFields_(
    'inventory.create',
    CONFIG.SHEETS.INVENTORY,
    INVENTORY_CREATE_ALLOWED,
    payload
  );

  var itemName = toTrimmedString_(payload.Item_Name);

  if (itemName === '') {
    throwError_(
      'Item_Name is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { field: 'Item_Name', value: payload.Item_Name }
    );
  }

  var sellingPrice = parseInventoryAmount_(
    payload.Selling_Price,
    'Selling_Price',
    {}
  );

  var currentQuantity = parseInventoryAmount_(
    payload.Current_Quantity,
    'Current_Quantity',
    { defaultValue: 0, integer: true }
  );

  var minimumStockLevel = parseInventoryAmount_(
    payload.Minimum_Stock_Level,
    'Minimum_Stock_Level',
    { defaultValue: 0, integer: true }
  );

  // A client-supplied Status is ignored outright: Status is derived from the
  // quantities so the sheet can never claim stock it does not have.
  var record = {
    Item_Name: itemName,
    Category: toTrimmedString_(payload.Category || ''),
    Unit: toTrimmedString_(payload.Unit || ''),
    Selling_Price: sellingPrice,
    Current_Quantity: currentQuantity,
    Minimum_Stock_Level: minimumStockLevel,
    Status: deriveInventoryStatus_(currentQuantity, minimumStockLevel)
  };

  return withScriptLock_(function () {
    var itemId = nextInventoryItemId_();
    record.Item_ID = itemId;

    if (findInventoryById_(itemId)) {
      throwError_(
        'Duplicate Item_ID generated: ' + itemId,
        ERROR_CODES.CONFLICT,
        { Item_ID: itemId }
      );
    }

    var created = normalizeInventoryRecord_(
      appendRow_(CONFIG.SHEETS.INVENTORY, record)
    );

    return success(created, 'Inventory item created');
  });
}
/**
 * Append one Inventory_Movements row for a stock change.
 *
 * MUST be called while the caller holds the critical section, so the ID this
 * function allocates cannot collide with a concurrent caller's. Inventory never
 * changes without a movement row; the two writes are one atomic operation.
 *
 * @param {string} movementType INVENTORY_MOVEMENT_TYPES value.
 * @param {{itemId: string, quantity: number, reason: string, notes: string}} request
 * @return {Object} The created movement record, normalised.
 */
function recordInventoryMovement_(movementType, request) {
  if (movementType !== INVENTORY_MOVEMENT_TYPES.STOCK_IN &&
      movementType !== INVENTORY_MOVEMENT_TYPES.STOCK_OUT) {
    throwError_(
      'Invalid inventory movement type: ' + movementType,
      ERROR_CODES.SERVER_ERROR,
      {
        movementType: movementType,
        validValues: [
          INVENTORY_MOVEMENT_TYPES.STOCK_IN,
          INVENTORY_MOVEMENT_TYPES.STOCK_OUT
        ]
      }
    );
  }

  var movementId = nextInventoryMovementId_();

  var movement = {
    Movement_ID: movementId,
    Item_ID: request.itemId,
    Movement_Type: movementType,
    Quantity: request.quantity,
    Date: formatDate_(now_()),
    Reason: request.reason,
    Recorded_By: inventoryActor_(),
    Notes: request.notes
  };

  if (findMovementById_(movementId)) {
    throwError_(
      'Duplicate Movement_ID generated: ' + movementId,
      ERROR_CODES.CONFLICT,
      { Movement_ID: movementId }
    );
  }

  var sheet = getMovementsSheet_();
  var headers = getMovementsHeaders_();
  var values = recordToValues_(
    CONFIG.SHEETS.INVENTORY_MOVEMENTS,
    headers,
    movement
  );

  var targetRow = sheet.getLastRow() + 1;

  sheet
    .getRange(targetRow, 1, 1, headers.length)
    .setValues([values]);

  return normalizeMovementRecord_(
    rowToObject_(headers, values)
  );
}

/**
 * Apply one stock change: read -> validate -> item update -> movement write.
 *
 * The caller MUST already hold the critical section (withScriptLock_ or
 * withInventoryLock_): the read and both writes have to be atomic, otherwise a
 * concurrent stock-out could be validated against a stale quantity and drive
 * Current_Quantity negative.
 *
 * @param {string} movementType INVENTORY_MOVEMENT_TYPES value.
 * @param {{itemId: string, quantity: number, reason: string, notes: string}} request
 * @param {string} message Response message.
 * @return {Object} success() envelope with the refreshed item and the movement.
 */
function applyInventoryStockChange_(movementType, request, message) {
  // Fail before anything is written: every stock change MUST produce a ledger
  // row, so the movements sheet (and its columns) is proven usable up front.
  // Otherwise a missing/incomplete ledger tab would leave the item's quantity
  // changed with no movement recorded.
  getMovementsHeaders_();

  var found = findInventoryById_(request.itemId);

  if (!found) {
    throwError_(
      'Inventory item not found: ' + request.itemId,
      ERROR_CODES.NOT_FOUND,
      { Item_ID: request.itemId }
    );
  }

  var current = found.record;
  var previousQuantity = Number(current.Current_Quantity);

  if (!isFinite(previousQuantity)) previousQuantity = 0;

  var newQuantity = movementType === INVENTORY_MOVEMENT_TYPES.STOCK_IN
    ? previousQuantity + request.quantity
    : previousQuantity - request.quantity;

  // The only guard against negative stock, and it reads the quantity acquired
  // under this very lock.
  if (newQuantity < 0) {
    throwError_(
      'Insufficient stock for item ' + request.itemId + ': requested ' +
        request.quantity + ', available ' + previousQuantity + '.',
      ERROR_CODES.VALIDATION_ERROR,
      {
        Item_ID: request.itemId,
        requested: request.quantity,
        available: previousQuantity,
        Movement_Type: movementType
      }
    );
  }

  current.Current_Quantity = newQuantity;
  current.Status = deriveInventoryStatus_(
    newQuantity,
    current.Minimum_Stock_Level
  );

  var headers = getInventoryHeaders_();
  var values = recordToValues_(
    CONFIG.SHEETS.INVENTORY,
    headers,
    current
  );

  var sheet = getInventorySheet_();
  sheet.getRange(found.sheetRow, 1, 1, headers.length).setValues([values]);

  var movement = recordInventoryMovement_(movementType, request);
  var refreshed = findInventoryById_(request.itemId);

  return success(
    {
      item: refreshed.record,
      movement: movement,
      Previous_Quantity: previousQuantity,
      Current_Quantity: newQuantity
    },
    message
  );
}
/**
 * Run `fn` inside the project's critical section, re-entering it only when this
 * execution does not already hold the script lock.
 *
 * This exists for ONE reason: Stationery.handleStationeryFulfill_ calls
 * adjustInventoryStockOut_ while it already holds withScriptLock_. Apps Script's
 * script lock is not re-entrant, so a second tryLock() from the same execution
 * fails and withScriptLock_ would answer "the system is busy" -- a
 * self-inflicted CONFLICT. Lock.hasLock() tells us whether THIS execution owns
 * the lock, so the nested call simply continues inside the section it already
 * holds.
 *
 * No second locking mechanism is introduced: it is the same LockService script
 * lock, acquired exactly once per execution. When hasLock() is unavailable (or
 * false) this degrades to plain withScriptLock_, exactly like every other
 * mutating path in the project.
 *
 * @param {Function} fn Work to run while the lock is held.
 * @return {*} Whatever `fn` returns.
 */
function withInventoryLock_(fn) {
  var lock = LockService.getScriptLock();

  if (typeof lock.hasLock === 'function' && lock.hasLock()) {
    return fn();
  }

  return withScriptLock_(fn);
}

/**
 * inventory.stockIn
 *
 * Requires INVENTORY.ADJUST. Increases Current_Quantity and appends one
 * STOCK_IN movement. Validation runs before the lock; the read, the quantity
 * update and the movement write all happen inside it.
 */
function handleInventoryStockIn_(payload) {
  if (!payload || typeof payload !== 'object') {
    throwError_(
      'Payload is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { reason: 'missing-payload' }
    );
  }

  var request = parseInventoryStockRequest_(
    payload,
    'inventory.stockIn',
    'INVENTORY.ADJUST'
  );

  return withScriptLock_(function () {
    return applyInventoryStockChange_(
      INVENTORY_MOVEMENT_TYPES.STOCK_IN,
      request,
      'Inventory stock-in recorded'
    );
  });
}

/**
 * inventory.stockOut
 *
 * Requires INVENTORY.ADJUST. Decreases Current_Quantity (never below zero) and
 * appends one STOCK_OUT movement.
 */
function handleInventoryStockOut_(payload) {
  if (!payload || typeof payload !== 'object') {
    throwError_(
      'Payload is required.',
      ERROR_CODES.VALIDATION_ERROR,
      { reason: 'missing-payload' }
    );
  }

  var request = parseInventoryStockRequest_(
    payload,
    'inventory.stockOut',
    'INVENTORY.ADJUST'
  );

  return withScriptLock_(function () {
    return applyInventoryStockChange_(
      INVENTORY_MOVEMENT_TYPES.STOCK_OUT,
      request,
      'Inventory stock-out recorded'
    );
  });
}

/**
 * Stationery.js entry point (stationery.fulfill): take `quantity` of `itemId`
 * out of stock and record the movement against `reference`.
 *
 * Contract used by Stationery.handleStationeryFulfill_:
 *
 *   var stockResult = adjustInventoryStockOut_(
 *     current.Item_ID, quantityToGive,
 *     'Stationery fulfillment ' + id, id
 *   );
 *   if (!stockResult || !stockResult.success) { ...throw... }
 *
 * so it returns an envelope and NEVER throws for a business failure -- the
 * caller inspects `success` and re-raises the code and details itself.
 *
 * The failure envelope nests the code under `error` ({ code, message, details })
 * because that is exactly what Stationery.handleStationeryFulfill_ reads. It is
 * deliberately NOT failure(), whose `error` is the flat code string.
 *
 * LOCKING: called from inside Stationery's withScriptLock_ section, so it goes
 * through withInventoryLock_, which re-enters the already-held section instead
 * of trying (and failing) to acquire the script lock a second time. Called
 * standalone it takes the lock itself. Either way the read -> validate ->
 * quantity update -> movement write is atomic.
 *
 * The permission check is deliberately omitted: authorization belongs to the
 * calling action (stationery.fulfill checks STATIONERY.FULFILL before reaching
 * here), and stationery staff have no reason to hold INVENTORY.ADJUST.
 *
 * @param {string} itemId Item_ID to take stock from.
 * @param {number} quantity Whole units to remove (must be > 0).
 * @param {string=} reason Free-text reason stored on the movement.
 * @param {string=} reference Source document, e.g. the ST-nnn transaction ID.
 * @return {Object} success() envelope, or { success: false, message, error:
 *     { code, message, details } } so the caller can re-throw verbatim.
 */
function adjustInventoryStockOut_(itemId, quantity, reason, reference) {
  var resolvedId = toTrimmedString_(itemId);
  var referenceText = toTrimmedString_(reference);
  var body = toTrimmedString_(reason) || 'Stock out';

  if (referenceText !== '') {
    body = body + ' (' + referenceText + ')';
  }

  var request = {
    itemId: resolvedId,
    quantity: Number(quantity),
    reason: body,
    notes: ''
  };

  try {
    if (resolvedId === '') {
      throwError_(
        'Item_ID is required for a stock-out.',
        ERROR_CODES.VALIDATION_ERROR,
        { field: 'Item_ID' }
      );
    }

    request.quantity = parseInventoryAmount_(quantity, 'Quantity', {
      positive: true,
      integer: true
    });

    return withInventoryLock_(function () {
      return applyInventoryStockChange_(
        INVENTORY_MOVEMENT_TYPES.STOCK_OUT,
        request,
        'Inventory stock-out recorded'
      );
    });
  } catch (err) {
    var message = err && err.message
      ? err.message
      : 'Inventory stock-out failed.';
    var code = err && err.code ? err.code : ERROR_CODES.SERVER_ERROR;
    var details = err && err.details ? err.details : null;

    return {
      success: false,
      message: message,
      error: {
        code: code,
        message: message,
        details: details
      }
    };
  }
}