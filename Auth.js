/**
 * Auth.js
 * Google-identity authentication for the School Management System API.
 *
 * MECHANISM: Google identity + Users-sheet allowlist. No passwords are
 * stored and no custom sessions or tokens are issued. The caller identity
 * comes ONLY from Session.getActiveUser().getEmail(), never from request
 * payload, query parameters, or headers.
 *
 * CAVEAT: getActiveUser() can return blank in some web-app contexts. A
 * blank email is UNAUTHENTICATED (fail closed). Live availability under
 * executeAs USER_DEPLOYING + access MYSELF must be confirmed via auth.me.
 *
 * SCHEMA ASSUMPTION: Users header is
 * User_ID, Staff_ID, Email, Role, Status, Last_Login. Role holds a role
 * NAME matched by Role_Name semantics. No header is renamed or created.
 */

var AUTH_NO_IDENTITY_MESSAGE =
  'Google identity could not be determined in this execution context. ' +
  'Confirm Session.getActiveUser().getEmail() on the deployed web app.';

function getAuthenticatedEmail_() {
  try {
    if (typeof Session === 'undefined' || !Session.getActiveUser) return '';
    var user = Session.getActiveUser();
    if (!user || typeof user.getEmail !== 'function') return '';
    return toTrimmedString_(user.getEmail());
  } catch (err) {
    return '';
  }
}

function normalizeRoleKey_(role) {
  return toTrimmedString_(role).toUpperCase();
}

function findActiveUserByEmail_(email) {
  var sheetName = CONFIG.SHEETS.USERS;
  var sheet;
  try {
    sheet = getSheet_(sheetName);
  } catch (err) {
    if (err && err.code === ERROR_CODES.NOT_FOUND) {
      throwError_('Users sheet is missing or unconfigured; authentication requires an existing Users tab.',
        ERROR_CODES.SERVER_ERROR, { sheet: sheetName, reason: 'sheet-missing' });
    }
    throw err;
  }
  var headers = getHeaders_(sheet);
  if (headers.length === 0) {
    throwError_('Sheet "' + sheetName + '" has no header row.',
      ERROR_CODES.SERVER_ERROR, { sheet: sheetName });
  }
  var required = ['User_ID', 'Email', 'Role', 'Status'];
  var missing = required.filter(function (c) { return headers.indexOf(c) === -1; });
  if (missing.length > 0) {
    throwError_('Sheet "' + sheetName + '" missing column(s): ' + missing.join(', ') + '.',
      ERROR_CODES.SERVER_ERROR,
      { sheet: sheetName, missingColumns: missing, availableColumns: headers });
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throwError_('No account found for this Google identity.',
      ERROR_CODES.NOT_FOUND, { email: email });
  }
  var eIdx = headers.indexOf('Email');
  var sIdx = headers.indexOf('Status');
  var emails = sheet.getRange(2, eIdx + 1, lastRow - 1, 1).getValues();
  var statuses = sheet.getRange(2, sIdx + 1, lastRow - 1, 1).getValues();
  var wanted = String(email).trim().toLowerCase();
  var matches = [];
  for (var i = 0; i < emails.length; i++) {
    if (toTrimmedString_(emails[i][0]).toLowerCase() !== wanted) continue;
    if (toTrimmedString_(statuses[i][0]).toLowerCase() !== 'active') continue;
    matches.push(i + 2);
  }
  if (matches.length === 0) {
    throwError_('No account found for this Google identity.',
      ERROR_CODES.NOT_FOUND, { email: email });
  }
  if (matches.length > 1) {
    throwError_('Multiple active accounts share this email. Contact an administrator.',
      ERROR_CODES.CONFLICT, { email: email, matchingRows: matches });
  }
  var row = matches[0];
  var vals = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  return { sheetRow: row, record: rowToObject_(headers, vals) };
}

function getCurrentUser_() {
  var email = getAuthenticatedEmail_();
  if (email === '') {
    return { user: null, error: appError_(AUTH_NO_IDENTITY_MESSAGE,
      ERROR_CODES.UNAUTHORIZED, { reason: 'no-google-identity' }) };
  }
  var found;
  try {
    found = findActiveUserByEmail_(email);
  } catch (err) {
    if (err && err.code === ERROR_CODES.NOT_FOUND) {
      return { user: null, error: appError_(
        'No active account found for ' + email + '. Contact an administrator.',
        ERROR_CODES.UNAUTHORIZED, { reason: 'no-matching-user', email: email }) };
    }
    return { user: null, error: err };
  }
  var r = found.record;
  var user = { userId: toTrimmedString_(r.User_ID), staffId: toTrimmedString_(r.Staff_ID),
    email: toTrimmedString_(r.Email), role: toTrimmedString_(r.Role) };
  updateLastLogin_(found.sheetRow, r);
  return { user: user, error: null };
}

function requireAuthentication_() {
  var result = getCurrentUser_();
  if (result.error) throw result.error;
  return result.user;
}

function updateLastLogin_(sheetRow, record) {
  try {
    var sheet = getSheetOrNull_(CONFIG.SHEETS.USERS);
    if (!sheet) return;
    var headers = getHeaders_(sheet);
    if (headers.indexOf('Last_Login') === -1) return;
    var today = formatDate_(now_());
    if (today === '') return;
    var current = toTrimmedString_(record.Last_Login);
    var currentDay = '';
    if (current !== '') {
      var parsed = toDate_(current);
      currentDay = parsed ? formatDate_(parsed) : current;
    }
    if (currentDay === today) return;
    setCellValue_(CONFIG.SHEETS.USERS, sheetRow, 'Last_Login', formatDateTime_(now_()));
  } catch (err) { /* bookkeeping must never break login */ }
}

function handleAuthMe_() {
  var user = requireAuthentication_();
  return success({ userId: user.userId, staffId: user.staffId,
    email: user.email, role: user.role }, 'Authenticated');
}

function handleAuthCheck_(payload) {
  var user = requireAuthentication_();
  var body = (payload && typeof payload === 'object' && !Array.isArray(payload)) ? payload : {};
  if (body.permission === undefined || body.permission === null || body.permission === '') {
    return success({ authenticated: true,
      user: { userId: user.userId, staffId: user.staffId, email: user.email, role: user.role },
      permission: null, allowed: true }, 'Authenticated');
  }
  if (typeof body.permission !== 'string' || body.permission.trim() === '') {
    throwError_('"permission" must be a non-empty string like "STUDENTS.READ".',
      ERROR_CODES.VALIDATION_ERROR, { received: body.permission });
  }
  var permission = body.permission.trim().toUpperCase();
  assertValidPermissionFormat_(permission);
  var allowed = hasPermission_(user, permission);
  return success({ authenticated: true,
    user: { userId: user.userId, staffId: user.staffId, email: user.email, role: user.role },
    permission: permission, allowed: allowed },
    allowed ? 'Permission granted' : 'Permission denied');
}
