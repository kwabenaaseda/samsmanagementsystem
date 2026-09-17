/**
 * Permissions.js
 * Centralized authorization for the School Management System API.
 *
 * Permission format: MODULE.ACTION in UPPER_SNAKE form (e.g.
 * 'STUDENTS.READ'). Default-deny: unknown permissions, unknown roles, and
 * inactive roles all deny. No role check lives outside this file.
 *
 * PHASE 4A: grants come from the spreadsheet, not from code. The flow is
 *
 *   Users.Role (role name) -> Roles -> Role_Permissions -> Permissions
 *
 * There is deliberately NO hard-coded "Admin allows everything" rule and no
 * in-code mapping table. If a role has no active Role_Permissions rows, it
 * has no permissions -- including Admin.
 *
 * SCHEMA (preferred, created by setupRolePermissions()):
 *   Role_Permissions: Role_Permission_ID, Role_ID, Permission_ID, Status
 *   Roles / Permissions keep their existing schemas; the reader ADAPTS to
 *   their actual columns by detecting the role-name column
 *   (Role_Name | Role | Name) and the permission-code column
 *   (Permission_Name | Permission | Permission_Code | Code | Name). A
 *   missing Role_ID / Permission_ID column is tolerated: the role name /
 *   permission name itself is then used as the join key.
 *
 * FAIL-SAFE RULES (never silently allow):
 *   - Missing Roles / Permissions / Role_Permissions sheet -> SERVER_ERROR.
 *   - A mapping row whose Permission_ID has no Permissions row -> SERVER_ERROR
 *     (orphan reference -- surface it, do not ignore it).
 *   - Duplicate (Role_ID, Permission_ID) rows with the SAME effective status
 *     are de-duplicated deterministically; with CONFLICTING statuses
 *     (Active vs Inactive) -> CONFLICT so the bad row is fixed.
 *   - Mapping rows with a non-'Active' Status grant nothing (inactive deny).
 */

var ROLE_PERMISSION_COLUMNS = ['Role_Permission_ID', 'Role_ID', 'Permission_ID', 'Status'];
var ROLE_NAME_COLUMN_ALIASES = ['Role_Name', 'Role', 'Name'];
var PERMISSION_CODE_COLUMN_ALIASES = ['Permission_Name', 'Permission', 'Permission_Code', 'Code', 'Name'];

function assertValidPermissionFormat_(permission) {
  if (typeof permission !== 'string' || !/^[A-Z][A-Z0-9_]*\.[A-Z][A-Z0-9_]*$/.test(permission)) {
    throwError_('Permission must look like "MODULE.ACTION" (e.g. "STUDENTS.READ").',
      ERROR_CODES.VALIDATION_ERROR, { received: permission });
  }
}

/**
 * Read the Roles sheet and detect its join columns.
 * @return {{sheet: Sheet, headers: string[], nameColumn: string, idColumn: string|null}}
 * @throws {Error} SERVER_ERROR when the sheet or its columns are unusable.
 */
function getRolesContext_() {
  var sheetName = CONFIG.SHEETS.ROLES;
  var sheet;
  try {
    sheet = getSheet_(sheetName);
  } catch (err) {
    if (err && err.code === ERROR_CODES.NOT_FOUND) {
      throwError_('Roles sheet "' + sheetName + '" is missing; authorization cannot resolve roles.',
        ERROR_CODES.SERVER_ERROR, { sheet: sheetName, reason: 'sheet-missing' });
    }
    throw err;
  }
  var headers = getHeaders_(sheet);
  var nameColumn = null;
  for (var i = 0; i < ROLE_NAME_COLUMN_ALIASES.length; i++) {
    if (headers.indexOf(ROLE_NAME_COLUMN_ALIASES[i]) !== -1) {
      nameColumn = ROLE_NAME_COLUMN_ALIASES[i];
      break;
    }
  }
  var idColumn = headers.indexOf('Role_ID') !== -1 ? 'Role_ID' : null;
  if (!nameColumn) {
    throwError_('Roles sheet "' + sheetName + '" has no role-name column (looked for: ' +
      ROLE_NAME_COLUMN_ALIASES.join(', ') + ').',
      ERROR_CODES.SERVER_ERROR, { sheet: sheetName, availableColumns: headers });
  }
  return { sheet: sheet, headers: headers, nameColumn: nameColumn, idColumn: idColumn };
}

/**
 * Find a role row by name (case-insensitive), falling back to Role_ID.
 * @return {{sheetRow: number, record: Object}|null}
 */
function findRoleRow_(roles, roleKey) {
  var wanted = toTrimmedString_(roleKey).toLowerCase();
  if (wanted === '') return null;
  var lastRow = roles.sheet.getLastRow();
  if (lastRow < 2) return null;
  var values = roles.sheet.getRange(1, 1, lastRow, roles.headers.length).getValues();
  for (var i = 1; i < values.length; i++) {
    if (isBlankRow_(values[i])) continue;
    var record = rowToObject_(roles.headers, values[i]);
    var name = toTrimmedString_(record[roles.nameColumn]).toLowerCase();
    if (name === wanted) return { sheetRow: i + 1, record: record };
    if (roles.idColumn && toTrimmedString_(record[roles.idColumn]).toLowerCase() === wanted) {
      return { sheetRow: i + 1, record: record };
    }
  }
  return null;
}

/**
 * Read every Role_Permissions row as header-keyed records, skipping blanks.
 * @return {Object[]}
 * @throws {Error} SERVER_ERROR when the sheet is missing or lacks join columns.
 */
function readRolePermissionRows_() {
  var sheetName = CONFIG.SHEETS.ROLE_PERMISSIONS;
  var sheet;
  try {
    sheet = getSheet_(sheetName);
  } catch (err) {
    if (err && err.code === ERROR_CODES.NOT_FOUND) {
      throwError_('Role_Permissions sheet "' + sheetName + '" is missing; run setupRolePermissions() from the Apps Script editor to create and seed it.',
        ERROR_CODES.SERVER_ERROR, { sheet: sheetName, reason: 'sheet-missing' });
    }
    throw err;
  }
  var headers = getHeaders_(sheet);
  var required = ['Role_ID', 'Permission_ID'];
  var missing = required.filter(function (c) { return headers.indexOf(c) === -1; });
  if (missing.length > 0) {
    throwError_('Role_Permissions sheet "' + sheetName + '" is missing column(s): ' + missing.join(', ') + '.',
      ERROR_CODES.SERVER_ERROR, { sheet: sheetName, missingColumns: missing, availableColumns: headers });
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(1, 1, lastRow, headers.length).getValues();
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    if (isBlankRow_(values[i])) continue;
    rows.push(rowToObject_(headers, values[i]));
  }
  return rows;
}

/**
 * Read the Permissions sheet into lookup maps. Joins by Permission_ID when
 * that column exists, otherwise by the permission code itself (schema
 * adaptation -- see the file header).
 * @return {{byId: Object<string,string>, byCode: Object<string,string>, idColumn: string|null, codeColumn: string, codes: string[]}}
 * @throws {Error} SERVER_ERROR when the sheet or its columns are unusable.
 */
function getPermissionsIndex_() {
  var sheetName = CONFIG.SHEETS.PERMISSIONS;
  var sheet;
  try {
    sheet = getSheet_(sheetName);
  } catch (err) {
    if (err && err.code === ERROR_CODES.NOT_FOUND) {
      throwError_('Permissions sheet "' + sheetName + '" is missing; authorization cannot resolve permission names.',
        ERROR_CODES.SERVER_ERROR, { sheet: sheetName, reason: 'sheet-missing' });
    }
    throw err;
  }
  var headers = getHeaders_(sheet);
  var codeColumn = null;
  for (var i = 0; i < PERMISSION_CODE_COLUMN_ALIASES.length; i++) {
    if (headers.indexOf(PERMISSION_CODE_COLUMN_ALIASES[i]) !== -1) {
      codeColumn = PERMISSION_CODE_COLUMN_ALIASES[i];
      break;
    }
  }
  if (!codeColumn) {
    throwError_('Permissions sheet "' + sheetName + '" has no permission-name column (looked for: ' +
      PERMISSION_CODE_COLUMN_ALIASES.join(', ') + ').',
      ERROR_CODES.SERVER_ERROR, { sheet: sheetName, availableColumns: headers });
  }
  var idColumn = headers.indexOf('Permission_ID') !== -1 ? 'Permission_ID' : null;
  var lastRow = sheet.getLastRow();
  var byId = {};
  var byCode = {};
  var idByCode = {};
  if (lastRow >= 2) {
    var values = sheet.getRange(1, 1, lastRow, headers.length).getValues();
    for (var r = 1; r < values.length; r++) {
      if (isBlankRow_(values[r])) continue;
      var record = rowToObject_(headers, values[r]);
      var code = toTrimmedString_(record[codeColumn]).toUpperCase();
      if (code === '') continue;
      byCode[code] = code;
      var id = idColumn ? toTrimmedString_(record[idColumn]) : code;
      if (id !== '') {
        byId[id.toLowerCase()] = code;
        idByCode[code] = id;
      }
    }
  }
  return {
    byId: byId,
    byCode: byCode,
    idByCode: idByCode,
    idColumn: idColumn,
    codeColumn: codeColumn,
    codes: Object.keys(byCode).sort()
  };
}

/**
 * Resolve the permission codes granted to a role through Role_Permissions.
 * This is THE source of authorization -- there is no in-code fallback.
 *
 * @param {string} roleKey Role name (or Role_ID) as held in Users.Role.
 * @return {string[]} Granted permission codes; empty when the role is
 *     unknown or has no active mapping rows (default-deny).
 * @throws {Error} SERVER_ERROR for structural problems; CONFLICT for
 *     duplicate mapping rows with conflicting statuses.
 */
function resolveRolePermissions_(roleKey) {
  var roleName = toTrimmedString_(roleKey);
  if (roleName === '') return [];

  var roles = getRolesContext_();
  var role = findRoleRow_(roles, roleName);
  if (!role) return []; // unknown role -> no permissions (default-deny)

  var roleId = toTrimmedString_(role.record[roles.idColumn || roles.nameColumn]);
  var rows = readRolePermissionRows_();

  // Deterministic duplicate handling: the same (Role_ID, Permission_ID)
  // twice with the same effective status de-duplicates; conflicting
  // statuses surface as CONFLICT so the sheet gets fixed instead of guessed.
  var effective = {};
  var conflicts = [];
  for (var i = 0; i < rows.length; i++) {
    if (toTrimmedString_(rows[i].Role_ID).toLowerCase() !== roleId.toLowerCase()) continue;
    var permissionId = toTrimmedString_(rows[i].Permission_ID);
    if (permissionId === '') continue; // a blank join key can never grant
    var active = toTrimmedString_(rows[i].Status).toLowerCase() === 'active';
    if (!effective.hasOwnProperty(permissionId)) {
      effective[permissionId] = active;
      continue;
    }
    if (effective[permissionId] !== active && conflicts.indexOf(permissionId) === -1) {
      conflicts.push(permissionId);
    }
  }
  if (conflicts.length > 0) {
    throwError_('Role "' + roleName + '" has duplicate Role_Permissions rows with conflicting Status for Permission_ID(s): ' +
      conflicts.join(', ') + '.',
      ERROR_CODES.CONFLICT, { sheet: CONFIG.SHEETS.ROLE_PERMISSIONS, role: roleName, permissionIds: conflicts });
  }

  var activeIds = Object.keys(effective).filter(function (id) { return effective[id]; });
  if (activeIds.length === 0) return [];

  var permissions = getPermissionsIndex_();
  var codes = [];
  var orphans = [];
  for (i = 0; i < activeIds.length; i++) {
    var code = permissions.byId[activeIds[i].toLowerCase()];
    if (code === undefined) {
      orphans.push(activeIds[i]);
      continue;
    }
    codes.push(code);
  }
  if (orphans.length > 0) {
    throwError_('Role_Permissions grants Permission_ID(s) that do not exist in the Permissions sheet: ' +
      orphans.join(', ') + '.',
      ERROR_CODES.SERVER_ERROR,
      { sheet: CONFIG.SHEETS.PERMISSIONS, role: roleName, permissionIds: orphans, reason: 'orphan-mapping' });
  }
  return codes;
}

function hasPermission_(userContext, permission) {
  if (!userContext || typeof userContext !== 'object') return false;
  if (typeof permission !== 'string' || !/^[A-Z][A-Z0-9_]*\.[A-Z][A-Z0-9_]*$/.test(permission)) {
    return false;
  }
  var grants = resolveRolePermissions_(normalizeRoleKey_(userContext.role));
  for (var i = 0; i < grants.length; i++) {
    if (grants[i] === '*') return true;
    if (grants[i] === permission) return true;
  }
  return false;
}

function requirePermission_(permission) {
  assertValidPermissionFormat_(permission);
  var user = requireAuthentication_();
  if (!hasPermission_(user, permission)) {
    throwError_('Permission denied: ' + permission + ' is required.',
      ERROR_CODES.FORBIDDEN, { permission: permission, role: user.role });
  }
  return user;
}

/* ==========================================================================
 * Setup / seed (Phase 4A)
 * ======================================================================== */

/**
 * Create and seed the Role_Permissions sheet. Run ONCE from the Apps Script
 * editor (Run > setupRolePermissions). It is deliberately NOT an API action:
 * it is an owner-run setup step, documented in README.md, so there is no
 * hidden in-code mapping -- everything it writes is visible in the sheet.
 *
 * What it does (idempotent -- safe to re-run):
 *   1. Creates the Role_Permissions tab with the canonical columns if absent.
 *   2. Adds a Permissions row for every code in CONFIG.PERMISSION_CODES that
 *      the Permissions sheet does not already have (matched by name).
 *   3. Maps the Admin role (matched by name in the Roles sheet, then by
 *      Role_ID) to EVERY permission code with Status 'Active'.
 *   4. Grants NOTHING to any other role -- deny-by-default. Granting
 *      Teacher/Accountant/etc. is a deliberate later decision made by
 *      editing the Role_Permissions sheet, not by this seed.
 *
 * It does NOT create or alter the Roles / Permissions sheets; if either is
 * missing or unreadable it fails with SERVER_ERROR instead of guessing.
 *
 * @return {Object} Report of what already existed and what was created.
 */
function setupRolePermissions() {
  return withScriptLock_(function () {
    var report = {
      rolePermissionsSheetCreated: false,
      permissionsAdded: [],
      mappingsCreated: [],
      permissionsAlreadyPresent: 0,
      mappingsAlreadyPresent: 0
    };
    var spreadsheet = getSpreadsheet_();

    // 1. Role_Permissions sheet with the canonical columns.
    var sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.ROLE_PERMISSIONS);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CONFIG.SHEETS.ROLE_PERMISSIONS);
      report.rolePermissionsSheetCreated = true;
    }
    var headers = getHeaders_(sheet);
    if (headers.length === 0) {
      sheet.getRange(1, 1, 1, ROLE_PERMISSION_COLUMNS.length).setValues([ROLE_PERMISSION_COLUMNS]);
      headers = ROLE_PERMISSION_COLUMNS.slice();
    }
    var missing = ROLE_PERMISSION_COLUMNS.filter(function (c) { return headers.indexOf(c) === -1; });
    if (missing.length > 0) {
      throwError_('Role_Permissions sheet exists but is missing column(s): ' + missing.join(', ') + '.',
        ERROR_CODES.SERVER_ERROR, { sheet: CONFIG.SHEETS.ROLE_PERMISSIONS, missingColumns: missing });
    }

    // 2. Ensure every canonical permission code has a Permissions row.
    var permissions = getPermissionsIndex_();
    CONFIG.PERMISSION_CODES.forEach(function (code) {
      if (permissions.byCode[code.toUpperCase()]) {
        report.permissionsAlreadyPresent += 1;
        return;
      }
      var record = {};
      if (permissions.idColumn) record[permissions.idColumn] = generateId_('PERM');
      record[permissions.codeColumn] = code;
      if (getHeaders_(getSheet_(CONFIG.SHEETS.PERMISSIONS)).indexOf('Status') !== -1) {
        record.Status = 'Active';
      }
      appendRow_(CONFIG.SHEETS.PERMISSIONS, record);
      report.permissionsAdded.push(code);
    });

    // 3. Map the Admin role to every permission code.
    var roles = getRolesContext_();
    var admin = findRoleRow_(roles, 'Admin');
    if (!admin) {
      throwError_('No "Admin" role row was found in the Roles sheet; the seed cannot map permissions.',
        ERROR_CODES.SERVER_ERROR, { sheet: CONFIG.SHEETS.ROLES, reason: 'admin-role-missing' });
    }
    var adminId = toTrimmedString_(admin.record[roles.idColumn || roles.nameColumn]);

    // Re-read after step 2 so freshly added rows are included.
    permissions = getPermissionsIndex_();
    var existing = {};
    readRolePermissionRows_().forEach(function (row) {
      if (toTrimmedString_(row.Role_ID).toLowerCase() !== adminId.toLowerCase()) return;
      var id = toTrimmedString_(row.Permission_ID).toLowerCase();
      if (id !== '') existing[id] = true;
    });

    permissions.codes.forEach(function (code) {
      var permissionId = permissions.idColumn ? permissions.idByCode[code] : code;
      if (!permissionId) return; // unreachable for rows indexed above
      if (existing[permissionId.toLowerCase()]) {
        report.mappingsAlreadyPresent += 1;
        return;
      }
      appendRow_(CONFIG.SHEETS.ROLE_PERMISSIONS, {
        Role_Permission_ID: generateId_('RP'),
        Role_ID: adminId,
        Permission_ID: permissionId,
        Status: 'Active'
      });
      report.mappingsCreated.push(code);
    });

    return report;
  });
}
