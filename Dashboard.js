/**
 * Dashboard.js
 *
 * Phase 6: dashboard summary aggregate.
 *
 * SCHEMA ASSUMPTIONS:
 *   Students: Status column exists (ACTIVE / WITHDRAWN / ...)
 *   Staff: Employment_Status column exists (ACTIVE / INACTIVE / ...)
 *   School_Fees: fee obligation rows (Amount_Due = Fee Amount)
 *   School_Fee_Payments: payment ledger (Amount, Status, Fee_ID → School_Fees)
 *   Feeding_Fees: Amount_Paid column exists
 *   Inventory: Status column exists (derived from Current_Quantity vs Minimum_Stock_Level)
 *
 * PERMISSION: DASHBOARD.READ (gated via requirePermission_ in Router.js).
 *
 * RESPONSE SHAPE:
 *   {
 *     activeStudents: number,        // Students where Status = 'Active'
 *     activeStaff: number,           // Staff where Employment_Status = 'Active'
 *     schoolFeesCollected: number,   // sum of Amount_Paid across School_Fees
 *     feedingFeesCollected: number,  // sum of Amount_Paid across Feeding_Fees
 *     lowStockItems: number,         // Inventory where Status != 'In Stock'
 *     recentPayments: Array<{
 *       type: 'schoolFees' | 'feedingFees',
 *       Payment_ID: string,
 *       Student_ID: string,
 *       Amount_Paid: number,
 *       Payment_Date: string
 *     }>  // last 5 across both fee sheets, most recent first
 *   }
 *
 * NOTE: No date-range filtering (current term/year) yet — sums everything.
 *       That is a future refinement, not built now.
 */

/* ==========================================================================
 * Aggregations
 * ======================================================================== */

/**
 * Count rows where a given column equals a trimmed value.
 * @param {string} sheetName Sheet/tab name.
 * @param {string} column Header to check.
 * @param {string} value Expected trimmed value.
 * @return {number} Matching data-row count.
 */
function countWhere_(sheetName, column, value) {
  var sheet = getSheet_(sheetName);
  var headers = getHeaders_(sheet);
  if (headers.length === 0) return 0;
  var colIndex = headers.indexOf(column);
  if (colIndex === -1) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var values = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  var wanted = toTrimmedString_(value);
  var count = 0;
  for (var i = 0; i < values.length; i++) {
    if (toTrimmedString_(values[i][0]) === wanted) count++;
  }
  return count;
}


/**
 * Sum a numeric column across every data row.
 * Blank / non-numeric cells are treated as 0.
 * @param {string} sheetName
 * @param {string} column Header to sum.
 * @return {number}
 */
function sumColumn_(sheetName, column) {
  var sheet = getSheet_(sheetName);
  var headers = getHeaders_(sheet);
  if (headers.length === 0) return 0;
  var colIndex = headers.indexOf(column);
  if (colIndex === -1) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var values = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  var total = 0;
  for (var i = 0; i < values.length; i++) {
    var n = Number(values[i][0]);
    if (!isNaN(n)) total += n;
  }
  return total;
}

/**
 * Sum the Amount column across NON-VOIDED School_Fee_Payments rows.
 * This is "money actually received" — fee obligations are never counted.
 * An absent payments sheet (fresh deployment) counts as zero collected.
 * @return {number}
 */
function sumSchoolFeePayments_() {
  try {
    var sheet = getSheet_(CONFIG.SHEETS.SCHOOL_FEE_PAYMENTS);
  } catch (err) {
    // Missing tab: no payments recorded yet, not a service failure.
    return 0;
  }
  var headers = getHeaders_(sheet);
  var amountIdx = headers.indexOf('Amount');
  var statusIdx = headers.indexOf('Status');
  if (amountIdx === -1) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var total = 0;
  for (var i = 0; i < values.length; i++) {
    if (statusIdx !== -1 && toTrimmedString_(values[i][statusIdx]) === 'Voided') continue;
    var n = Number(values[i][amountIdx]);
    if (!isNaN(n)) total += n;
  }
  return total;
}

/**
 * Count rows where the Status column does NOT equal 'In Stock'.
 * Used to report low / out-of-stock items for the dashboard.
 * @param {string} sheetName
 * @return {number}
 */
function countNonInStock_(sheetName) {
  var sheet = getSheet_(sheetName);
  var headers = getHeaders_(sheet);
  if (headers.length === 0) return 0;
  var colIndex = headers.indexOf('Status');
  if (colIndex === -1) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var values = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  var count = 0;
  for (var i = 0; i < values.length; i++) {
    if (toTrimmedString_(values[i][0]) !== 'In Stock') count++;
  }
  return count;
}

/**
 * Read the two fee sheets, merge into one flat list enriched with a `type`
 * discriminator, sort by Payment_Date descending, and return the 5 most recent.
 * @return {{type: string, Payment_ID: string, Student_ID: string, Amount_Paid: number, Payment_Date: string}[]}
 */
function recentPayments_(schoolFeesSheetName, feedingFeesSheetName) {
  var ffSheet = getSheet_(feedingFeesSheetName);

  var ffHeaders = getHeaders_(ffSheet);

  var ffLastRow = ffSheet.getLastRow();

  var all = [];

  // School fee PAYMENTS (transactions). Student_ID is resolved through the
  // referenced fee obligation, so the ledger stays free of duplicated data.
  try {
    var fpSheet = getSheet_(CONFIG.SHEETS.SCHOOL_FEE_PAYMENTS);
    var fpHeaders = getHeaders_(fpSheet);
    var fpLastRow = fpSheet.getLastRow();
    if (fpLastRow >= 2 && fpHeaders.length > 0) {
      var feeIdIdx = fpHeaders.indexOf('Fee_ID');
      if (feeIdIdx !== -1) {
        var fpValues = fpSheet.getRange(2, 1, fpLastRow - 1, fpHeaders.length).getValues();
        var feeStudent = {};
        for (var k = 0; k < fpValues.length; k++) {
          var fpRec = rowToObject_(fpHeaders, fpValues[k]);
          if (toTrimmedString_(fpRec.Status) === 'Voided') continue;
          var feeRow = findRowById_(
            CONFIG.SHEETS.SCHOOL_FEES,
            toTrimmedString_(fpRec.Fee_ID),
            'Payment_ID'
          );
          var sid = feeRow ? toTrimmedString_(feeRow.record.Student_ID) : '';
          feeStudent[toTrimmedString_(fpRec.Payment_ID)] = sid;
          all.push({
            type: 'schoolFees',
            Payment_ID: toTrimmedString_(fpRec.Payment_ID),
            Student_ID: sid,
            Amount_Paid: Number(fpRec.Amount) || 0,
            Payment_Date: toTrimmedString_(fpRec.Payment_Date)
          });
        }
      }
    }
  } catch (errPayments) {
    // Missing School_Fee_Payments tab: skip school-fee payments gracefully.
  }

  // Feeding_Fees rows
  if (ffLastRow >= 2 && ffHeaders.length > 0) {
    var ffValues = ffSheet.getRange(2, 1, ffLastRow - 1, ffHeaders.length).getValues();
    for (var i = 0; i < ffValues.length; i++) {
      var row = ffValues[i];
      if (isBlankRow_(row)) continue;
      var rec = rowToObject_(ffHeaders, row);
      all.push({
        type: 'feedingFees',
        Payment_ID: toTrimmedString_(rec.Payment_ID),
        Student_ID: toTrimmedString_(rec.Student_ID),
        Amount_Paid: Number(rec.Amount_Paid) || 0,
        Payment_Date: toTrimmedString_(rec.Payment_Date)
      });
    }
  }

  // Sort by Payment_Date descending (most recent first). Rows with no date
  // sort to the end so recent payments with dates are always preferred.
  all.sort(function (a, b) {
    var da = toDate_(a.Payment_Date);
    var db = toDate_(b.Payment_Date);
    if (!da && !db) return 0;
    if (!da) return 1;
    if (!db) return -1;
    return db.getTime() - da.getTime();
  });

  return all.slice(0, 5);
}

/* ==========================================================================
 * Handler
 * ======================================================================== */

/**
 * POST (or GET) dashboard.summary
 *
 * Aggregates counts and sums across Students, Staff, School_Fees,
 * Feeding_Fees, and Inventory in a single call. Requires DASHBOARD.READ.
 *
 * @param {Object} payload Not used by this action.
 * @param {Object} request Full request envelope from parseRequest_().
 * @return {Object} success envelope with DashboardSummary shape.
 */
function handleDashboardSummary_(payload, request) {
  requirePermission_('DASHBOARD.READ');

  var activeStudents = countWhere_(CONFIG.SHEETS.STUDENTS, 'Status', CONFIG.STUDENT_STATUS.ACTIVE);
  var activeStaff = countWhere_(CONFIG.SHEETS.STAFF, 'Employment_Status', CONFIG.STAFF_STATUS.ACTIVE);
  var schoolFeesCollected = sumSchoolFeePayments_();
  var feedingFeesCollected = sumColumn_(CONFIG.SHEETS.FEEDING_FEES, 'Amount_Paid');
  var lowStockItems = countNonInStock_(CONFIG.SHEETS.INVENTORY);
  var recentPayments = recentPayments_(CONFIG.SHEETS.SCHOOL_FEES, CONFIG.SHEETS.FEEDING_FEES);

  return success({
    activeStudents: activeStudents,
    activeStaff: activeStaff,
    schoolFeesCollected: schoolFeesCollected,
    feedingFeesCollected: feedingFeesCollected,
    lowStockItems: lowStockItems,
    recentPayments: recentPayments
  }, 'Dashboard summary retrieved');
}
