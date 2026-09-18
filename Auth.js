/**
 * Auth.js
 * Google-identity authentication for the School Management System API.
 *
 * MECHANISM: Google identity + Users-sheet allowlist. No passwords are
 * stored and no custom sessions or tokens are issued by SAMS itself.
 *
 * Identity is resolved in one of two ways, in priority order:
 *
 * 1. OAUTH TOKEN (primary, staff logins): the frontend obtains an OAuth
 *    access token via Google Identity Services and sends it in the request
 *    body's reserved `__auth.access_token` field (threaded here by Router.js).
 *    The token is verified SERVER-SIDE against Google's own userinfo endpoint
 *    (openidconnect.googleapis.com/v1/userinfo). The verified email from that
 *    call is the caller's identity. A caller-supplied email is NEVER trusted.
 *
 * 2. SESSION IDENTITY (fallback): Session.getActiveUser().getEmail(), the
 *    original mechanism, still works for owner-execution contexts.
 *
 * Google answers "who are you?"; the Users sheet answers "may you use SAMS,
 * and as what role?". A valid Google token belonging to an email that is not
 * an Active row in Users is UNAUTHORIZED, never entry.
 *
 * The access token is used transiently for the userinfo call and is never
 * written to any sheet or log.
 *
 * SCHEMA ASSUMPTION: Users header is
 * User_ID, Staff_ID, Email, Role, Status, Last_Login. Role holds a role
 * NAME matched by Role_Name semantics. No header is renamed or created.
 */

var AUTH_NO_IDENTITY_MESSAGE =
  'No Google identity could be verified for this request. Sign in with Google from the app.';

/**
 * Request-scoped OAuth access token, set by Router.js handleRequest_ before
 * dispatch and cleared afterwards. Apps Script executes each request in its
 * own context, so a module-level slot is safe.
 */
var REQUEST_AUTH_TOKEN_ = null;

/** Google's OpenID Connect userinfo endpoint (first-class, not debug tokeninfo). */
var GOOGLE_USERINFO_ENDPOINT_ = 'https://openidconnect.googleapis.com/v1/userinfo';

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

/**
 * Resolve the caller's email by verifying the OAuth access token against
 * Google's userinfo endpoint, entirely server-side.
 *
 * @param {string} token The bearer access token from __auth.access_token.
 * @return {string} The verified email address.
 * @throws UNAUTHORIZED with a distinct `reason` detail for every failure mode.
 */
function resolveCallerEmailFromToken_(token) {
  if (typeof UrlFetchApp === 'undefined' || !UrlFetchApp || typeof UrlFetchApp.fetch !== 'function') {
    throwError_('Token verification is unavailable in this execution context.',
      ERROR_CODES.UNAUTHORIZED, { reason: 'token-verification-unavailable' });
  }

  var response;
  try {
    response = UrlFetchApp.fetch(GOOGLE_USERINFO_ENDPOINT_, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
    });
  } catch (err) {
    throwError_('Could not verify the Google identity token.',
      ERROR_CODES.UNAUTHORIZED, { reason: 'token-verification-failed' });
  }

  var code = response && typeof response.getResponseCode === 'function'
    ? response.getResponseCode() : 0;
  if (code === 401 || code === 403) {
    // Expired, revoked, or issued for a different client/audience.
    throwError_('The Google sign-in token is invalid or has expired. Sign in again.',
      ERROR_CODES.UNAUTHORIZED, { reason: 'token-invalid' });
  }
  if (code !== 200) {
    throwError_('The Google identity token could not be verified.',
      ERROR_CODES.UNAUTHORIZED, { reason: 'token-verification-failed' });
  }

  var info;
  try {
    info = JSON.parse(response.getContentText());
  } catch (err) {
    throwError_('The Google identity token could not be verified.',
      ERROR_CODES.UNAUTHORIZED, { reason: 'token-verification-failed' });
  }

  if (!info || info.email_verified !== true || !info.email) {
    throwError_('The Google account has no verified email address.',
      ERROR_CODES.UNAUTHORIZED, { reason: 'email-unverified' });
  }
  return toTrimmedString_(info.email);
}

/**
 * Resolve the caller's email: verified OAuth token first, session identity as
 * fallback. Returns '' when neither yields an identity.
 */
function getCallerEmail_() {
  if (REQUEST_AUTH_TOKEN_) return resolveCallerEmailFromToken_(REQUEST_AUTH_TOKEN_);
  return getAuthenticatedEmail_();
}

/** Test/dispatch hook: install or clear the request-scoped token. */
function setRequestAuthToken_(token) {
  REQUEST_AUTH_TOKEN_ = (typeof token === 'string' && token.trim() !== '') ? token.trim() : null;
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
  var email;
  try {
    email = getCallerEmail_();
  } catch (err) {
    // Token resolution failures (unverified email, invalid/expired token,
    // userinfo outage) are auth outcomes, not crashes: report them as a
    // structured error result so every caller sees { user, error }.
    return { user: null, error: err };
  }
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
