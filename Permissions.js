/**
 * Permissions.js
 * Centralized authorization for the School Management System API.
 *
 * Permission format: MODULE.ACTION in UPPER_SNAKE form (e.g.
 * 'STUDENTS.READ'). Default-deny: unknown permissions, unknown roles, and
 * inactive roles all deny. No role check lives outside this file.
 *
 * TEMPORARY MAPPING (Phase 2): the spreadsheet has NO Role_Permissions
 * mapping table, so role grants live in TEMP_ROLE_PERMISSIONS_ below. Only
 * the Admin role is granted access; every other role name resolves to NO
 * permissions. This is intentionally deny-by-default, NOT a statement
 * about what teachers or accountants should eventually be able to do.
 *
 * REPLACEMENT PATH: when the Role_Permissions sheet exists, replace the
 * body of resolveRolePermissions_() with a sheet read. hasPermission_(),
 * requirePermission_(), Auth.js, and Router.js are unaffected.
 */

var TEMP_ROLE_PERMISSIONS_ = {
  // TEMPORARY (Phase 2): no Role_Permissions sheet exists, so only the
  // Admin role is granted access. Every other role resolves to NO
  // permissions until the Role_Permissions sheet lands. This is
  // intentionally deny-by-default, NOT a statement about what teachers
  // or accountants should eventually be able to do.
  ADMIN: ['*']
};

function assertValidPermissionFormat_(permission) {
  if (typeof permission !== 'string' || !/^[A-Z][A-Z0-9_]*\.[A-Z][A-Z0-9_]*$/.test(permission)) {
    throwError_('Permission must look like "MODULE.ACTION" (e.g. "STUDENTS.READ").',
      ERROR_CODES.VALIDATION_ERROR, { received: permission });
  }
}

function resolveRolePermissions_(roleKey) {
  var key = toTrimmedString_(roleKey).toUpperCase();
  if (key === '' || !TEMP_ROLE_PERMISSIONS_.hasOwnProperty(key)) return [];
  return TEMP_ROLE_PERMISSIONS_[key].slice();
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
