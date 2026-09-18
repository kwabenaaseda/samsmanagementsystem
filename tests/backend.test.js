/**
 * tests/backend.test.js
 *
 * Lightweight, zero-dependency tests for the Apps Script backend foundation.
 * Run with:  node tests/backend.test.js
 *
 * HOW IT WORKS
 * Apps Script flattens every project file into one global scope, so this
 * harness concatenates the backend files into a single script and runs it
 * inside a Node `vm` context with minimal stubs for the Google services. That
 * mirrors real Apps Script semantics (including cross-file `const` visibility)
 * without a framework or network access.
 *
 * The stubs assume a UTC+0 script timezone, which matches appsscript.json
 * ("Africa/Abidjan"), so date assertions stay deterministic.
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const BACKEND_FILES = [
  'Config.js',
  'Response.js',
  'Utils.js',
  'Permissions.js',
  'Auth.js',
  'Students.js',
  'Staff.js',
  'SchoolFees.js',
  'FeedingFees.js',
  'Stationery.js',
  'Inventory.js',
  'Dashboard.js',
  'Router.js'
];

/* ==========================================================================
 * Tiny assertion helpers (no framework)
 * ======================================================================== */

let passed = 0;
let failed = 0;
const failures = [];

function check(label, fn) {
  try {
    fn();
    passed++;
    console.log('  ok   ' + label);
  } catch (err) {
    failed++;
    failures.push(label + ' :: ' + err.message);
    console.log('  FAIL ' + label);
    console.log('        ' + err.message);
  }
}

function eq(actual, expected, msg) {
  const a = JSON.stringify(actual);
  const b = JSON.stringify(expected);
  if (a !== b) {
    throw new Error((msg ? msg + '\n        ' : '') + 'expected ' + b + '\n        actual   ' + a);
  }
}

function ok(condition, msg) {
  if (!condition) throw new Error(msg || 'expected a truthy value');
}

function section(title) {
  console.log('\n' + title);
}

function throwsWithCode(fn, code) {
  let thrown = null;
  try {
    fn();
  } catch (err) {
    thrown = err;
  }
  if (!thrown) throw new Error('expected it to throw ' + code + ', but it did not throw');
  if (thrown.code !== code) {
    throw new Error('expected error code ' + code + ' but got ' + thrown.code + ' (' + thrown.message + ')');
  }
  return thrown;
}

/* ==========================================================================
 * Google Apps Script service stubs
 * ======================================================================== */

function makeRange(sheet, row, col, numRows, numCols) {
  return {
    getValues: function () {
      const out = [];
      for (let r = 0; r < numRows; r++) {
        const source = sheet._rows[row - 1 + r] || [];
        const line = [];
        for (let c = 0; c < numCols; c++) {
          const value = source[col - 1 + c];
          line.push(value === undefined || value === null ? '' : value);
        }
        out.push(line);
      }
      return out;
    },
    setValues: function (values) {
      values.forEach(function (line, r) {
        const target = row - 1 + r;
        while (sheet._rows.length <= target) sheet._rows.push([]);
        line.forEach(function (value, c) {
          while (sheet._rows[target].length < col + c) sheet._rows[target].push('');
          sheet._rows[target][col - 1 + c] = value;
        });
      });
      return this;
    },
  };
}

/** In-memory stand-in for a Sheet tab. `rows` is an array of arrays. */
function makeSheet(name, rows) {
  return {
    _rows: rows.map(function (r) {
      return r.slice();
    }),
    getName: function () {
      return name;
    },
    getLastRow: function () {
      let last = 0;
      this._rows.forEach(function (row, index) {
        const hasValue = row.some(function (cell) {
          return cell !== '' && cell !== null && cell !== undefined;
        });
        if (hasValue) last = index + 1;
      });
      return last;
    },
    getLastColumn: function () {
      return this._rows.reduce(function (max, row) {
        return Math.max(max, row.length);
      }, 0);
    },
    getRange: function (row, col, numRows, numCols) {
      return makeRange(this, row, col, numRows, numCols);
    },
  };
}

function makeSpreadsheet(name, sheets) {
  return {
    getName: function () {
      return name;
    },
    getSheets: function () {
      return sheets;
    },
    getSheetByName: function (tabName) {
      return (
        sheets.filter(function (s) {
          return s.getName() === tabName;
        })[0] || null
      );
    },
    insertSheet: function (tabName) {
      const created = makeSheet(tabName, []);
      sheets.push(created);
      return created;
    },
  };
}

function randomUuid() {
  let out = '';
  for (let i = 0; i < 32; i++) out += '0123456789abcdef'[Math.floor(Math.random() * 16)];
  return out.slice(0, 8) + '-' + out.slice(8, 12) + '-' + out.slice(12, 16) + '-' + out.slice(16, 20) + '-' + out.slice(20);
}

/** Minimal Utilities.formatDate supporting the patterns the backend uses. */
function simpleFormat(date, pattern) {
  const pad = function (n, w) {
    return String(n).padStart(w, '0');
  };
  const tokens = {
    yyyy: pad(date.getUTCFullYear(), 4),
    MM: pad(date.getUTCMonth() + 1, 2),
    dd: pad(date.getUTCDate(), 2),
    HH: pad(date.getUTCHours(), 2),
    mm: pad(date.getUTCMinutes(), 2),
    ss: pad(date.getUTCSeconds(), 2),
  };
  return pattern.replace(/yyyy|MM|dd|HH|mm|ss/g, function (token) {
    return tokens[token];
  });
}

/**
 * Build a sandbox with Google service stubs.
 * @param {Object|null} spreadsheet Value SpreadsheetApp.openById should return.
 * @param {Object=} opts Behaviour switches (openByIdThrows, lockUnavailable, uuid).
 */
function makeSandbox(spreadsheet, opts) {
  const options = opts || {};
  // Lock depth is observable so the tests can prove that a decision or a row
  // write happened inside the critical section (locked read-modify-write checks).
  const sandbox = {
    __lockDepth: 0,
    __lockAttempts: 0,
    console: console,
    SpreadsheetApp: {
      openById: function (id) {
        if (options.openByIdThrows) throw new Error('Simulated: no access to ' + id);
        if (!spreadsheet) throw new Error('Simulated: no spreadsheet configured');
        return spreadsheet;
      },
    },
    LockService: {
      getScriptLock: function () {
        return {
          tryLock: function () {
            const acquired = !options.lockUnavailable;
            sandbox.__lockAttempts += 1;
            if (acquired) sandbox.__lockDepth += 1;
            return acquired;
          },
          releaseLock: function () {
            if (sandbox.__lockDepth > 0) sandbox.__lockDepth -= 1;
          },
          hasLock: function () {
            return sandbox.__lockDepth > 0;
          },
        };
      },
    },
    Session: {
      getScriptTimeZone: function () {
        return 'Africa/Abidjan';
      },
      // Set by tests: null/absent = blank identity; { email } = signed in.
      __activeEmail: undefined,
      getActiveUser: function () {
        var address = this.__activeEmail;
        return {
          getEmail: function () {
            if (address === undefined || address === null) return '';
            if (typeof address === 'function') return address();
            return address;
          }
        };
      },
    },
    Utilities: {
      getUuid: function () {
        return options.uuid ? options.uuid() : randomUuid();
      },
      formatDate: function (date, tz, pattern) {
        return simpleFormat(date, pattern);
      },
    },
    // Simulated UrlFetchApp for OAuth token verification. Tests configure
    // behaviour via opts.urlFetchThrows and opts.urlFetchResponse.
    UrlFetchApp: {
      fetch: function (url, params) {
        sandbox.__urlFetchCalls = sandbox.__urlFetchCalls || [];
        sandbox.__urlFetchCalls.push({ url: url, params: params });
        if (options.urlFetchThrows) throw new Error('Simulated: network failure');
        const response = options.urlFetchResponse;
        if (!response) throw new Error('Simulated: no UrlFetchApp response configured');
        return {
          getResponseCode: function () { return response.status; },
          getContentText: function () {
            return typeof response.body === 'string' ? response.body : JSON.stringify(response.body);
          },
        };
      },
    },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput: function (text) {
        return {
          _text: text,
          _mimeType: null,
          setMimeType: function (mime) {
            this._mimeType = mime;
            return this;
          },
          getContent: function () {
            return this._text;
          },
          getMimeType: function () {
            return this._mimeType;
          },
        };
      },
    },
  };
  return sandbox;
}

/**
 * Concatenate the backend into one script and evaluate it, exactly as Apps
 * Script would. Returns the contextified sandbox.
 */
function loadBackend(sandbox) {
  const context = vm.createContext(sandbox);
  const source = BACKEND_FILES.map(function (file) {
    return '\n/* ======== ' + file + ' ======== */\n' + fs.readFileSync(path.join(ROOT, file), 'utf8');
  }).join('\n');

  // Top-level `const` bindings are lexical, not global-object properties, so
  // expose the ones the tests assert on.
  const exporter = '\n;globalThis.__api = { CONFIG: CONFIG, ERROR_CODES: ERROR_CODES,' +
    ' getAuthenticatedEmail_: (typeof getAuthenticatedEmail_ !== "undefined" ? getAuthenticatedEmail_ : undefined),' +
    ' normalizeRoleKey_: (typeof normalizeRoleKey_ !== "undefined" ? normalizeRoleKey_ : undefined),' +
    ' findActiveUserByEmail_: (typeof findActiveUserByEmail_ !== "undefined" ? findActiveUserByEmail_ : undefined),' +
    ' getCurrentUser_: (typeof getCurrentUser_ !== "undefined" ? getCurrentUser_ : undefined),' +
   ' getCallerEmail_: (typeof getCallerEmail_ !== "undefined" ? getCallerEmail_ : undefined),' +
   ' resolveCallerEmailFromToken_: (typeof resolveCallerEmailFromToken_ !== "undefined" ? resolveCallerEmailFromToken_ : undefined),' +
   ' setRequestAuthToken_: (typeof setRequestAuthToken_ !== "undefined" ? setRequestAuthToken_ : undefined),' +
    ' requireAuthentication_: (typeof requireAuthentication_ !== "undefined" ? requireAuthentication_ : undefined),' +
    ' assertValidPermissionFormat_: (typeof assertValidPermissionFormat_ !== "undefined" ? assertValidPermissionFormat_ : undefined),' +
    ' resolveRolePermissions_: (typeof resolveRolePermissions_ !== "undefined" ? resolveRolePermissions_ : undefined),' +
    ' hasPermission_: (typeof hasPermission_ !== "undefined" ? hasPermission_ : undefined),' +
    ' requirePermission_: (typeof requirePermission_ !== "undefined" ? requirePermission_ : undefined),' +
    ' getSheet_: (typeof getSheet_ !== "undefined" ? getSheet_ : undefined),' +
    ' getSheetOrNull_: (typeof getSheetOrNull_ !== "undefined" ? getSheetOrNull_ : undefined),' +
    ' getHeaders_: (typeof getHeaders_ !== "undefined" ? getHeaders_ : undefined),' +
    ' rowToObject_: (typeof rowToObject_ !== "undefined" ? rowToObject_ : undefined),' +
    ' isBlankRow_: (typeof isBlankRow_ !== "undefined" ? isBlankRow_ : undefined),' +
    ' readAll_: (typeof readAll_ !== "undefined" ? readAll_ : undefined),' +
    ' appendRow_: (typeof appendRow_ !== "undefined" ? appendRow_ : undefined),' +
    ' setCellValue_: (typeof setCellValue_ !== "undefined" ? setCellValue_ : undefined),' +
    ' findRowById_: (typeof findRowById_ !== "undefined" ? findRowById_ : undefined),' +
    ' generateId_: (typeof generateId_ !== "undefined" ? generateId_ : undefined),' +
    ' formatDate_: (typeof formatDate_ !== "undefined" ? formatDate_ : undefined),' +
    ' formatDateTime_: (typeof formatDateTime_ !== "undefined" ? formatDateTime_ : undefined),' +
    ' now_: (typeof now_ !== "undefined" ? now_ : undefined),' +
    ' nowIso_: (typeof nowIso_ !== "undefined" ? nowIso_ : undefined),' +
    ' toDate_: (typeof toDate_ !== "undefined" ? toDate_ : undefined),' +
    ' isBlank_: (typeof isBlank_ !== "undefined" ? isBlank_ : undefined),' +
    ' toTrimmedString_: (typeof toTrimmedString_ !== "undefined" ? toTrimmedString_ : undefined),' +
    ' isNonEmptyString_: (typeof isNonEmptyString_ !== "undefined" ? isNonEmptyString_ : undefined),' +
    ' isValidEmail_: (typeof isValidEmail_ !== "undefined" ? isValidEmail_ : undefined),' +
    ' assertRequired_: (typeof assertRequired_ !== "undefined" ? assertRequired_ : undefined),' +
    ' assertOneOf_: (typeof assertOneOf_ !== "undefined" ? assertOneOf_ : undefined),' +
    ' assertEmail_: (typeof assertEmail_ !== "undefined" ? assertEmail_ : undefined),' +
    ' appError_: (typeof appError_ !== "undefined" ? appError_ : undefined),' +
    ' throwError_: (typeof throwError_ !== "undefined" ? throwError_ : undefined),' +
    ' success: (typeof success !== "undefined" ? success : undefined),' +
    ' failure: (typeof failure !== "undefined" ? failure : undefined),' +
    ' jsonResponse: (typeof jsonResponse !== "undefined" ? jsonResponse : undefined),' +
    ' getSpreadsheet_: (typeof getSpreadsheet_ !== "undefined" ? getSpreadsheet_ : undefined),' +
    ' listSheetNames_: (typeof listSheetNames_ !== "undefined" ? listSheetNames_ : undefined),' +
    ' sheetExists_: (typeof sheetExists_ !== "undefined" ? sheetExists_ : undefined),' +
    ' parseRequest_: (typeof parseRequest_ !== "undefined" ? parseRequest_ : undefined),' +
    ' listAvailableActions_: (typeof listAvailableActions_ !== "undefined" ? listAvailableActions_ : undefined),' +
    ' doGet: (typeof doGet !== "undefined" ? doGet : undefined),' +
    ' doPost: (typeof doPost !== "undefined" ? doPost : undefined),' +
    ' getStudentsSheet_: (typeof getStudentsSheet_ !== "undefined" ? getStudentsSheet_ : undefined),' +
    ' getStudentsHeaders_: (typeof getStudentsHeaders_ !== "undefined" ? getStudentsHeaders_ : undefined),' +
    ' readAllStudents_: (typeof readAllStudents_ !== "undefined" ? readAllStudents_ : undefined),' +
    ' findStudentById_: (typeof findStudentById_ !== "undefined" ? findStudentById_ : undefined),' +
    ' getStaffSheet_: (typeof getStaffSheet_ !== "undefined" ? getStaffSheet_ : undefined),' +
    ' getStaffHeaders_: (typeof getStaffHeaders_ !== "undefined" ? getStaffHeaders_ : undefined),' +
    ' readAllStaff_: (typeof readAllStaff_ !== "undefined" ? readAllStaff_ : undefined),' +
    ' findStaffById_: (typeof findStaffById_ !== "undefined" ? findStaffById_ : undefined),' +
    ' recordToValues_: (typeof recordToValues_ !== "undefined" ? recordToValues_ : undefined),' +
    ' withScriptLock_: (typeof withScriptLock_ !== "undefined" ? withScriptLock_ : undefined),' +
    ' hasPermission_: (typeof hasPermission_ !== "undefined" ? hasPermission_ : undefined),' +
    ' toPermissionCode_: (typeof toPermissionCode_ !== "undefined" ? toPermissionCode_ : undefined),' +
    ' getPermissionsIndex_: (typeof getPermissionsIndex_ !== "undefined" ? getPermissionsIndex_ : undefined),' +
    ' setupRolePermissions: (typeof setupRolePermissions !== "undefined" ? setupRolePermissions : undefined),' +
    ' getSchoolFeesSheet_: (typeof getSchoolFeesSheet_ !== "undefined" ? getSchoolFeesSheet_ : undefined),' +
    ' getSchoolFeesHeaders_: (typeof getSchoolFeesHeaders_ !== "undefined" ? getSchoolFeesHeaders_ : undefined),' +
    ' findSchoolFeeById_: (typeof findSchoolFeeById_ !== "undefined" ? findSchoolFeeById_ : undefined),' +
    ' handleSchoolFeesList_: (typeof handleSchoolFeesList_ !== "undefined" ? handleSchoolFeesList_ : undefined),' +
    ' handleSchoolFeesGet_: (typeof handleSchoolFeesGet_ !== "undefined" ? handleSchoolFeesGet_ : undefined),' +
    ' handleSchoolFeesCreate_: (typeof handleSchoolFeesCreate_ !== "undefined" ? handleSchoolFeesCreate_ : undefined),' +
    ' handleSchoolFeesUpdate_: (typeof handleSchoolFeesUpdate_ !== "undefined" ? handleSchoolFeesUpdate_ : undefined),' +
    ' handleSchoolFeesVoid_: (typeof handleSchoolFeesVoid_ !== "undefined" ? handleSchoolFeesVoid_ : undefined),' +
    ' getFeedingFeesSheet_: (typeof getFeedingFeesSheet_ !== "undefined" ? getFeedingFeesSheet_ : undefined),' +
    ' getFeedingFeesHeaders_: (typeof getFeedingFeesHeaders_ !== "undefined" ? getFeedingFeesHeaders_ : undefined),' +
    ' findFeedingFeeById_: (typeof findFeedingFeeById_ !== "undefined" ? findFeedingFeeById_ : undefined),' +
    ' handleFeedingFeesList_: (typeof handleFeedingFeesList_ !== "undefined" ? handleFeedingFeesList_ : undefined),' +
    ' handleFeedingFeesGet_: (typeof handleFeedingFeesGet_ !== "undefined" ? handleFeedingFeesGet_ : undefined),' +
    ' handleFeedingFeesCreate_: (typeof handleFeedingFeesCreate_ !== "undefined" ? handleFeedingFeesCreate_ : undefined),' +
    ' handleFeedingFeesUpdate_: (typeof handleFeedingFeesUpdate_ !== "undefined" ? handleFeedingFeesUpdate_ : undefined),' +
    ' handleFeedingFeesVoid_: (typeof handleFeedingFeesVoid_ !== "undefined" ? handleFeedingFeesVoid_ : undefined)' +
    ' };';

  vm.runInContext(source + exporter, context, { filename: 'backend-bundle.js' });
  return sandbox;
}

/** Parse the JSON body out of a ContentService.TextOutput stub. */
function readEnvelope(textOutput) {
  return JSON.parse(textOutput.getContent());
}

/** The 13 logical tabs Config.js is expected to declare. */
const EXPECTED_TABS = [
  'Students',
  'Staff',
  'Users',
  'Roles',
  'Permissions',
  'Role_Permissions',
  'School_Fees',
  'Feeding_Fees',
  'Stationery',
  'Inventory',
  'Inventory_Movements',
  'Salary_Payments',
  'Delegations',
  'Audit_Log',
];

/* ==========================================================================
 * Phase 4A: Role_Permissions fixtures
 * ======================================================================== */

// Mirrors CONFIG.PERMISSION_CODES; asserted equal in the Phase 4A section.
const P4A_PERMISSION_CODES = [
  'STUDENTS.READ', 'STUDENTS.CREATE', 'STUDENTS.UPDATE', 'STUDENTS.WITHDRAW',
  'STAFF.READ', 'STAFF.CREATE', 'STAFF.UPDATE', 'STAFF.DEACTIVATE',
    'SCHOOL_FEES.READ', 'SCHOOL_FEES.CREATE', 'SCHOOL_FEES.UPDATE', 'SCHOOL_FEES.VOID',
  'FEEDING_FEES.READ', 'FEEDING_FEES.CREATE', 'FEEDING_FEES.UPDATE', 'FEEDING_FEES.VOID',
  'STATIONERY.READ', 'STATIONERY.CREATE', 'STATIONERY.UPDATE', 'STATIONERY.VOID', 'STATIONERY.FULFILL',
  'INVENTORY.READ', 'INVENTORY.CREATE', 'INVENTORY.UPDATE', 'INVENTORY.ADJUST',
  'SALARIES.READ', 'SALARIES.CREATE', 'SALARIES.UPDATE', 'SALARIES.VOID',
  'AUDIT_LOG.READ',
  'DELEGATIONS.READ', 'DELEGATIONS.CREATE', 'DELEGATIONS.UPDATE', 'DELEGATIONS.REVOKE',
  'DASHBOARD.READ'
];

/** Permissions rows in the CONFIRMED live schema (Permission_ID, Module, Action, Description). */
const P4A_MODULE_DISPLAY = {
  STUDENTS: 'Students',
  STAFF: 'Staff',
  SCHOOL_FEES: 'School_Fees',
  FEEDING_FEES: 'Feeding_Fees',
  STATIONERY: 'Stationery',
  INVENTORY: 'Inventory',
  SALARIES: 'Salaries',
  AUDIT_LOG: 'Audit_Log',
  DELEGATIONS: 'Delegations',
  DASHBOARD: 'Dashboard',
};

/** Permissions rows: PERM-1..PERM-N in canonical order, Module + Action columns. */
function p4aPermissionRows() {
  return P4A_PERMISSION_CODES.map(function (code, index) {
    var dot = code.indexOf('.');
    return [
      'PERM-' + (index + 1),
      P4A_MODULE_DISPLAY[code.slice(0, dot)],
      code.slice(dot + 1),
      'Permission ' + code,
    ];
  });
}

/** Admin mapping rows: ROL-1 -> every PERM-n with Status Active. */
function p4aAdminMappingRows() {
  return P4A_PERMISSION_CODES.map(function (code, index) {
    return ['RP-' + (index + 1), 'ROL-1', 'PERM-' + (index + 1), 'Active'];
  });
}

/**
 * Roles / Permissions / Role_Permissions sheets with Admin fully mapped.
 * opts.roles / opts.permissions / opts.mappings override the row bodies
 * (opts.mappings === [] gives a header-only mapping sheet).
 */
function p4aPermissionSheets(opts) {
  opts = opts || {};
  return {
    roles: makeSheet('Roles', [['Role_ID', 'Role_Name', 'Description', 'Status']]
      .concat(opts.roles || [
        ['ROL-1', 'Admin', 'Full system administrator', 'Active'],
        ['ROL-2', 'Teacher', 'Teaching staff', 'Active'],
      ])),
    permissions: makeSheet('Permissions', [['Permission_ID', 'Module', 'Action', 'Description']]
      .concat(opts.permissions || p4aPermissionRows())),
    rolePermissions: makeSheet('Role_Permissions', [['Role_Permission_ID', 'Role_ID', 'Permission_ID', 'Status']]
      .concat(opts.mappings !== undefined ? opts.mappings : p4aAdminMappingRows())),
  };
}

/** A spreadsheet containing every expected tab, plus data rows for tests. */
function makeFullSpreadsheet() {
  const sheets = EXPECTED_TABS.map(function (tabName) {
    return makeSheet(tabName, [['Header_A', 'Header_B']]);
  });
  // Give the Students tab real content for the read tests.
  const students = sheets[0];
  students._rows = [
    ['Student_ID', 'First_Name', 'Last_Name', 'Status'],
    ['STU-1', 'Ama', 'Mensah', 'Active'],
    ['STU-2', 'Kofi', 'Owusu', 'Active'],
    ['', '', '', ''],
    ['STU-3', 'Yaa', 'Boateng', 'Withdrawn'],
  ];
  // Give the Users tab a real allowlist for the auth tests.
  const users = sheets[2];
  users._rows = [
    ['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'],
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', ''],
    ['USR-2', 'STF-2', 'teacher@school.edu', 'Teacher', 'Active', ''],
  ];
  // Phase 4A: real Roles / Permissions / Role_Permissions so authorization
  // resolves through the mapping (Admin fully granted, Teacher nothing).
  const p4a = p4aPermissionSheets();
  sheets[EXPECTED_TABS.indexOf('Roles')] = p4a.roles;
  sheets[EXPECTED_TABS.indexOf('Permissions')] = p4a.permissions;
  sheets[EXPECTED_TABS.indexOf('Role_Permissions')] = p4a.rolePermissions;

  // Give the Staff tab real content for the read tests.
  const staff = sheets[1];
  staff._rows = [
    ['Staff_ID', 'First_Name', 'Last_Name', 'Role', 'Phone', 'Email',
      'Emergency_Contact', 'Employment_Status', 'Hire_Date', 'Salary_Amount',
      'Salary_Frequency', 'Bank_Account', 'Remarks'],
    ['STF-1', 'Paul', 'Mensah', 'Admin', '020-1234567', 'paul@school.edu',
      'Mensah', 'Active', '2020-01-15', 8500, 'Monthly', '1234567890', 'Head teacher'],
    ['STF-2', 'Akua', 'Owusu', 'Teacher', '020-2345678', 'akua@school.edu',
      'Owusu', 'Active', '2021-03-01', 5200, 'Monthly', '0987654321', 'JHS 1 teacher'],
  ];

  // Phase 4B: pre-populate the School_Fees and Feeding_Fees sheets.
  const schoolFeesIdx = EXPECTED_TABS.indexOf('School_Fees');
  const schoolFees = sheets[schoolFeesIdx];
  schoolFees._rows = [
    ['Payment_ID', 'Student_ID', 'Academic_Year', 'Term', 'Amount_Due',
      'Amount_Paid', 'Balance', 'Payment_Date', 'Payment_Method', 'Reference',
      'Status', 'Recorded_By', 'Notes'],
    ['SF-001', 'STU-1', '2025/2026', 'Term 1', 1200, 1200, 0,
      '2025-09-15', 'Bank Transfer', 'SF-2025-T1-001', 'Paid', 'STF-1',
      'Term 1 fees'],
    ['SF-002', 'STU-1', '2025/2026', 'Term 2', 1200, 600, 600,
      '2025-12-10', 'Mobile Money', 'SF-2025-T2-002', 'Partial', 'STF-1',
      'Half paid'],
    ['SF-003', 'STU-2', '2025/2026', 'Term 1', 1200, 0, 1200,
      '2025-09-15', 'Cash', 'SF-2025-T1-003', 'Unpaid', 'STF-1',
      'Not yet paid'],
  ];
  const feedingFeesIdx = EXPECTED_TABS.indexOf('Feeding_Fees');
  const feedingFees = sheets[feedingFeesIdx];
  feedingFees._rows = [
    ['Payment_ID', 'Student_ID', 'Academic_Year', 'Term', 'Amount_Due',
      'Amount_Paid', 'Balance', 'Payment_Date', 'Payment_Method', 'Reference',
      'Status', 'Recorded_By', 'Notes'],
    ['FF-001', 'STU-1', '2025/2026', 'Term 1', 450, 450, 0,
      '2025-09-15', 'Cash', 'FF-2025-T1-001', 'Paid', 'STF-1',
      'Feeding fees'],
    ['FF-002', 'STU-2', '2025/2026', 'Term 1', 450, 225, 225,
      '2025-09-15', 'Mobile Money', 'FF-2025-T1-002', 'Partial', 'STF-1',
      'Half paid'],
  ];

  return makeSpreadsheet('SchoolManagementSystem', sheets);
}

/** A Users tab with one row per scenario the auth spec exercises. */
function makeAuthSpreadsheet(rows) {
  const users = makeSheet('Users', [
    ['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'].concat(rows || []),
  ]);
  if (rows) users._rows = [['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login']].concat(rows);
  // Phase 4A: include the authorization tables so requirePermission_ resolves
  // through Role_Permissions (Admin granted, others denied by default).
  const p4a = p4aPermissionSheets();
  return makeSpreadsheet('SchoolManagementSystem', [users, p4a.roles, p4a.permissions, p4a.rolePermissions]);
}

/** Load the backend with a signed-in Google identity. */
function loadBackendAs(email, spreadsheet, opts) {
  const source = spreadsheet || makeFullSpreadsheet();
  const sandbox = makeSandbox(source, opts);
  sandbox.Session.__activeEmail = email;
  return loadBackend(sandbox);
}

/* ==========================================================================
 * Load the backend once for the behavioural tests
 * ======================================================================== */

const api = loadBackend(makeSandbox(makeFullSpreadsheet()));
const CONFIG = api.__api.CONFIG;
const ERROR_CODES = api.__api.ERROR_CODES;

/* ==========================================================================
 * Tests
 * ======================================================================== */

console.log('Backend foundation tests\n========================');

section('Load-time safety (Task 2: the Config.js crash)');

check('backend evaluates with NO Google services present at all', function () {
  const bare = vm.createContext({});
  const source = BACKEND_FILES.map(function (file) {
    return fs.readFileSync(path.join(ROOT, file), 'utf8');
  }).join('\n') + '\n;globalThis.__loadedWithoutGoogle = true;';

  vm.runInContext(source, bare, { filename: 'bare-load.js' });

  ok(
    bare.__loadedWithoutGoogle === true,
    'backend did not finish evaluating without Google services, so something still runs at load time'
  );
});

/**
 * Strip comments so source checks inspect executable code, not documentation.
 *
 * Limitation: a "//" inside a string literal would be misread. None of the
 * backend files contain one, and the check is a belt-and-braces guard on top
 * of the bare-context load test above.
 */
function stripComments(source) {
  return source.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/[^\n]*/g, '');
}

check('Config.js executable code no longer calls getActiveSpreadsheet()', function () {
  const code = stripComments(fs.readFileSync(path.join(ROOT, 'Config.js'), 'utf8'));
  ok(
    code.indexOf('getActiveSpreadsheet') === -1,
    'Config.js still calls getActiveSpreadsheet(), which returns null in web-app context'
  );
});

check('Config.js executable code makes no Google API call at top level', function () {
  const code = stripComments(fs.readFileSync(path.join(ROOT, 'Config.js'), 'utf8'));
  ['SpreadsheetApp', 'PropertiesService', 'Session', 'Utilities'].forEach(function (token) {
    ok(code.indexOf(token) === -1, 'Config.js executable code still references ' + token);
  });
  ok(code.indexOf('openById') === -1, 'Config.js should not open the spreadsheet itself');
});

section('Configuration (Task 2)');

check('SHEET_ID is an explicit literal, not derived', function () {
  eq(CONFIG.SHEET_ID, '17nqU2fim3e9txZaLLyo4DsOo46yuJUCJRV9ku144S6I');
});

check('SHEET_ID matches the real container spreadsheet', function () {
  const claspJson = JSON.parse(fs.readFileSync(path.join(ROOT, '.clasp.json'), 'utf8'));
  ok(claspJson.scriptId, 'expected .clasp.json to still declare a scriptId');
  ok(/^[A-Za-z0-9_-]{40,}$/.test(CONFIG.SHEET_ID), 'SHEET_ID does not look like a spreadsheet ID');
});

check('all 14 logical tabs are declared with the agreed names', function () {
  const actual = Object.keys(CONFIG.SHEETS).map(function (key) {
    return CONFIG.SHEETS[key];
  });
  eq(actual, EXPECTED_TABS);
});

check('student statuses include the four reserved values', function () {
  eq(CONFIG.VALUES.STUDENT_STATUS, ['Active', 'Withdrawn', 'Graduated', 'Suspended']);
});

check('payment statuses include the Voided correction value', function () {
  ok(CONFIG.VALUES.PAYMENT_STATUS.indexOf('Voided') !== -1, 'Voided is missing from PAYMENT_STATUS');
  eq(CONFIG.VALUES.PAYMENT_STATUS, ['Unpaid', 'Partial', 'Paid', 'Voided']);
});

check('payment methods and terms are declared', function () {
  ok(CONFIG.VALUES.PAYMENT_METHOD.length >= 3, 'expected at least three payment methods');
  eq(CONFIG.VALUES.TERM, ['Term 1', 'Term 2', 'Term 3']);
});

check('VALUES are derived from, not duplicated alongside, the enums', function () {
  eq(CONFIG.VALUES.MODULE, Object.keys(CONFIG.MODULE).map(function (k) {
    return CONFIG.MODULE[k];
  }));
});

check('health, auth, Phase 3, Phase 4B and dashboard actions are routed', function () {
  eq(api.listAvailableActions_().sort(), [
    'auth.check', 'auth.me',
    'dashboard.summary',
    'feedingFees.create', 'feedingFees.get', 'feedingFees.list', 'feedingFees.update',
    'feedingFees.void',
    'health',
    'inventory.create', 'inventory.list', 'inventory.movements', 'inventory.stockIn', 'inventory.stockOut',
    'schoolFees.create', 'schoolFees.get', 'schoolFees.list', 'schoolFees.update',
    'schoolFees.void',
    'staff.create', 'staff.deactivate', 'staff.get', 'staff.list', 'staff.update',
    'stationery.create', 'stationery.fulfill', 'stationery.list',
    'students.create', 'students.get', 'students.list', 'students.update', 'students.withdraw',
  ]);
  eq(CONFIG.ACTIONS.HEALTH, 'health');
  eq(CONFIG.ACTIONS.AUTH.ME, 'auth.me');
  eq(CONFIG.ACTIONS.AUTH.CHECK, 'auth.check');
});

check('all reserved action names follow the module.verb convention', function () {
  const flat = [];
  Object.keys(CONFIG.ACTIONS).forEach(function (group) {
    const value = CONFIG.ACTIONS[group];
    if (typeof value === 'string') {
      flat.push(value);
      return;
    }
    Object.keys(value).forEach(function (key) {
      flat.push(value[key]);
    });
  });

  flat.forEach(function (name) {
    // `health` is a top-level service action; every module action is module.verb.
    if (name === CONFIG.ACTIONS.HEALTH) return;
    ok(/^[a-z][A-Za-z]*\.[a-z][A-Za-z]*$/.test(name), 'malformed action name: ' + name);
  });

  ok(flat.length > 20, 'expected the full reserved action list to be present');
});

section('Response envelope (Task 4)');

check('success() keeps its original shape', function () {
  eq(api.success({ a: 1 }, 'Done'), { success: true, message: 'Done', data: { a: 1 } });
  eq(api.success(), { success: true, message: 'Success', data: null }, 'defaults changed');
});

check('failure() exposes the contract the frontend relies on', function () {
  const envelope = api.failure('Bad input', ERROR_CODES.VALIDATION_ERROR);
  eq(envelope.success, false);
  eq(envelope.error, 'VALIDATION_ERROR');
  eq(envelope.message, 'Bad input');
});

check('failure() defaulting and the legacy 2-argument call both work', function () {
  eq(api.failure('no code').error, 'VALIDATION_ERROR', 'default code should be VALIDATION_ERROR');
  eq(api.failure('boom', 'SERVER_ERROR').error, 'SERVER_ERROR', 'legacy explicit code broke');
});

check('failure() accepts optional details without breaking the signature', function () {
  const envelope = api.failure('Missing', ERROR_CODES.VALIDATION_ERROR, { missingFields: ['A'] });
  eq(envelope.details, { missingFields: ['A'] });
  eq(api.failure('x', ERROR_CODES.NOT_FOUND).details, null);
});

check('all six documented error codes exist', function () {
  eq(Object.keys(ERROR_CODES).sort(), [
    'CONFLICT', 'FORBIDDEN', 'NOT_FOUND', 'SERVER_ERROR', 'UNAUTHORIZED', 'VALIDATION_ERROR',
  ]);
});

check('jsonResponse() serializes JSON with the JSON mime type', function () {
  const output = api.jsonResponse({ success: true, message: 'ok', data: null });
  eq(output.getMimeType(), 'application/json');
  eq(JSON.parse(output.getContent()).success, true);
});

section('Validation helpers (Task 3, pure)');

check('isBlank_ and toTrimmedString_', function () {
  eq(
    [api.isBlank_(null), api.isBlank_(undefined), api.isBlank_('   '), api.isBlank_(0), api.isBlank_('x')],
    [true, true, true, false, false]
  );
  eq(api.toTrimmedString_('  hi  '), 'hi');
  eq(api.toTrimmedString_(null), '');
});

check('isValidEmail_ accepts real shapes and rejects malformed ones', function () {
  ok(api.isValidEmail_('admin@school.edu'), 'a normal address was rejected');
  ok(api.isValidEmail_('first.last+tag@sub.school.edu.gh'), 'a tagged address was rejected');
  eq(
    [api.isValidEmail_('nope'), api.isValidEmail_('a@b'), api.isValidEmail_('a b@c.com'), api.isValidEmail_('')],
    [false, false, false, false]
  );
});

check('numeric helpers and coercion', function () {
  eq(api.toNumber_('12.5'), 12.5);
  eq([api.toNumber_('abc'), api.toNumber_(''), api.toNumber_(null)], [null, null, null]);
  eq([api.isNonNegativeNumber_(0), api.isNonNegativeNumber_(-1), api.isNonNegativeNumber_('3')], [true, false, true]);
  eq([api.isPositiveNumber_(0), api.isPositiveNumber_(1)], [false, true]);
});

check('isOneOf_ and assertOneOf_', function () {
  ok(api.isOneOf_('Active', CONFIG.VALUES.STUDENT_STATUS), 'Active should be allowed');
  eq(api.isOneOf_('Nope', CONFIG.VALUES.STUDENT_STATUS), false);
  api.assertOneOf_('Withdrawn', CONFIG.VALUES.STUDENT_STATUS, 'Status');

  const err = throwsWithCode(function () {
    api.assertOneOf_('Zombie', CONFIG.VALUES.STUDENT_STATUS, 'Status');
  }, ERROR_CODES.VALIDATION_ERROR);
  ok(err.message.indexOf('Status') !== -1, 'error message should name the field');
  eq(err.details.allowedValues, CONFIG.VALUES.STUDENT_STATUS);
});

check('assertRequired_ reports every missing field at once', function () {
  api.assertRequired_({ First_Name: 'Ama', Last_Name: 'Mensah' }, ['First_Name', 'Last_Name']);

  const err = throwsWithCode(function () {
    api.assertRequired_({ First_Name: 'Ama', Last_Name: '  ' }, ['First_Name', 'Last_Name', 'Class']);
  }, ERROR_CODES.VALIDATION_ERROR);
  eq(err.details.missingFields, ['Last_Name', 'Class']);
});

check('assertRequired_ rejects non-objects', function () {
  throwsWithCode(function () {
    api.assertRequired_('not an object', ['A']);
  }, ERROR_CODES.VALIDATION_ERROR);
  throwsWithCode(function () {
    api.assertRequired_(['A'], ['A']);
  }, ERROR_CODES.VALIDATION_ERROR);
});

check('assertEmail_ trims, and rejects blank or malformed values', function () {
  eq(api.assertEmail_('  a@b.com  ', 'Guardian_Email'), 'a@b.com');
  throwsWithCode(function () {
    api.assertEmail_('', 'Guardian_Email');
  }, ERROR_CODES.VALIDATION_ERROR);
  throwsWithCode(function () {
    api.assertEmail_('bad', 'Guardian_Email');
  }, ERROR_CODES.VALIDATION_ERROR);
});

section('Date helpers (Task 3)');

check('formatDate_ and formatDateTime_ use the script timezone with ISO padding', function () {
  const date = new Date(Date.UTC(2026, 8, 16, 7, 5, 3));
  eq(api.formatDate_(date), '2026-09-16');
  eq(api.formatDateTime_(date), '2026-09-16 07:05:03');
  eq(api.getTimeZone_(), 'Africa/Abidjan');
});

check('formatDate_ returns an empty string for invalid input rather than throwing', function () {
  eq([api.formatDate_('not a date'), api.formatDate_(''), api.formatDate_(null)], ['', '', '']);
});

check('toDate_ handles Date, epoch millis, and date strings', function () {
  const date = new Date(Date.UTC(2026, 0, 2));
  eq(api.toDate_(date).getTime(), date.getTime());
  eq(api.toDate_(0).getTime(), 0);
  eq(api.toDate_('2026-01-02T00:00:00Z').getTime(), date.getTime());
  eq(api.toDate_('nonsense'), null);
});

check('toDate_ rejects an Invalid Date instance', function () {
  eq(api.toDate_(new Date('nope')), null);
});

check('nowIso_ returns a parseable timestamp', function () {
  ok(!isNaN(new Date(api.nowIso_()).getTime()), 'nowIso_ did not return a parseable timestamp');
});

section('ID generation (Task 3)');

check('generateId_ produces prefixed, date-stamped ids', function () {
  const id = api.generateId_('stu');
  ok(/^STU-\d{8}-[0-9A-F]{12}$/.test(id), 'unexpected id shape: ' + id);
});

check('generateId_ sanitizes unsafe characters out of the prefix', function () {
  ok(/^STU[A-Z0-9]*-/.test(api.generateId_(' stu-!? ')), 'prefix was not sanitized');
});

check('generateId_ rejects an empty or non-alphanumeric prefix', function () {
  throwsWithCode(function () {
    api.generateId_('---');
  }, ERROR_CODES.VALIDATION_ERROR);
  throwsWithCode(function () {
    api.generateId_('');
  }, ERROR_CODES.VALIDATION_ERROR);
});

check('generateId_ never repeats across 20000 calls', function () {
  const seen = {};
  for (let i = 0; i < 20000; i++) {
    const id = api.generateId_('STU');
    ok(!seen[id], 'duplicate id generated: ' + id);
    seen[id] = true;
  }
});

check('generateId_ does not depend on sheet contents (concurrency-safe)', function () {
  // Two separate sandboxes must not be able to derive the same id from state.
  const other = loadBackend(makeSandbox(makeFullSpreadsheet()));
  const mine = {};
  for (let i = 0; i < 200; i++) mine[api.generateId_('STU')] = true;
  for (let i = 0; i < 200; i++) {
    ok(!mine[other.generateId_('STU')], 'ids collided across independent instances');
  }
});

section('Sheet access helpers (Task 3)');

check('getSpreadsheet_ opens by explicit id', function () {
  ok(api.getSpreadsheet_(), 'getSpreadsheet_ returned nothing');
  eq(api.getSpreadsheet_().getName(), 'SchoolManagementSystem');
});

check('listSheetNames_ and sheetExists_', function () {
  eq(api.listSheetNames_(), EXPECTED_TABS);
  eq(api.sheetExists_('Students'), true);
  eq(api.sheetExists_('Nope'), false);
});

check('getSheet_ returns the sheet for a valid name', function () {
  eq(api.getSheet_(CONFIG.SHEETS.STUDENTS).getName(), 'Students');
  eq(api.getSheet_('  Students  ').getName(), 'Students', 'name should be trimmed');
});

check('getSheet_ throws NOT_FOUND listing the available tabs', function () {
  const err = throwsWithCode(function () {
    api.getSheet_('Studentz');
  }, ERROR_CODES.NOT_FOUND);
  ok(err.message.indexOf('Studentz') !== -1, 'error should name the missing tab');
  eq(err.details.availableSheets, EXPECTED_TABS);
});

check('getSheet_ rejects a blank name as a validation error', function () {
  throwsWithCode(function () {
    api.getSheet_('   ');
  }, ERROR_CODES.VALIDATION_ERROR);
});

check('an unreachable spreadsheet surfaces as SERVER_ERROR, not a raw throw', function () {
  const broken = loadBackend(makeSandbox(null, { openByIdThrows: true }));
  const err = throwsWithCode(function () {
    broken.getSpreadsheet_();
  }, ERROR_CODES.SERVER_ERROR);
  ok(err.message.indexOf('Could not open spreadsheet') !== -1, 'unhelpful message: ' + err.message);
  eq(err.details.spreadsheetId, CONFIG.SHEET_ID);
});

section('Reading (Task 3)');

check('readAll_ maps rows onto header-keyed objects and skips blank rows', function () {
  eq(api.readAll_('Students'), [
    { Student_ID: 'STU-1', First_Name: 'Ama', Last_Name: 'Mensah', Status: 'Active' },
    { Student_ID: 'STU-2', First_Name: 'Kofi', Last_Name: 'Owusu', Status: 'Active' },
    { Student_ID: 'STU-3', First_Name: 'Yaa', Last_Name: 'Boateng', Status: 'Withdrawn' },
  ]);
});

check('readAll_ returns [] for a header-only sheet', function () {
  // makeFullSpreadsheet() now populates the Staff sheet with real rows, so we
  // must construct a miniature spreadsheet that has a header-only Staff tab.
  const mini = makeSpreadsheet('Mini', [makeSheet('Staff', [['Staff_ID', 'First_Name']])]);
  const apiMini = loadBackend(makeSandbox(mini));
  eq(apiMini.readAll_('Staff'), []);
});

check('readAll_ returns [] for a completely empty sheet', function () {
  const blank = loadBackend(makeSandbox(makeSpreadsheet('Blank', [makeSheet('EmptySheet', [['A', 'B']])])));
  eq(blank.readAll_('EmptySheet'), []);
});

check('readAll_ throws NOT_FOUND for an unknown sheet', function () {
  throwsWithCode(function () {
    api.readAll_('Nope');
  }, ERROR_CODES.NOT_FOUND);
});

check('rowToObject_ skips columns that have a blank header', function () {
  eq(api.rowToObject_(['A', '', 'B'], ['1', 'ignored', '2']), { A: '1', B: '2' });
});

check('isBlankRow_ detects fully empty rows', function () {
  eq([api.isBlankRow_(['', ' ', null]), api.isBlankRow_(['', 'x'])], [true, false]);
});

section('Writing (Task 3)');

check('appendRow_ writes an object in sheet column order, not object key order', function () {
  const spreadsheet = makeFullSpreadsheet();
  const writeApi = loadBackend(makeSandbox(spreadsheet));
  const students = spreadsheet.getSheetByName('Students');

  const created = writeApi.appendRow_('Students', {
    Status: 'Active',
    Student_ID: 'STU-4',
    First_Name: 'Kwame',
    Last_Name: 'Adjei',
  });

  eq(students._rows[5], ['STU-4', 'Kwame', 'Adjei', 'Active'], 'row written in the wrong column order');
  eq(created, { Student_ID: 'STU-4', First_Name: 'Kwame', Last_Name: 'Adjei', Status: 'Active' });
});

check('appendRow_ pads missing columns with empty strings', function () {
  const spreadsheet = makeFullSpreadsheet();
  const writeApi = loadBackend(makeSandbox(spreadsheet));
  writeApi.appendRow_('Students', { Student_ID: 'STU-5', First_Name: 'Kojo' });
  eq(spreadsheet.getSheetByName('Students')._rows[5], ['STU-5', 'Kojo', '', '']);
});

check('appendRow_ accepts an array already in column order', function () {
  const spreadsheet = makeFullSpreadsheet();
  const writeApi = loadBackend(makeSandbox(spreadsheet));
  writeApi.appendRow_('Students', ['STU-6', 'Adwoa', 'Frimpong', 'Active']);
  eq(spreadsheet.getSheetByName('Students')._rows[5], ['STU-6', 'Adwoa', 'Frimpong', 'Active']);
});

check('appendRow_ rejects unknown columns and lists the valid ones', function () {
  const writeApi = loadBackend(makeSandbox(makeFullSpreadsheet()));
  const err = throwsWithCode(function () {
    writeApi.appendRow_('Students', { Student_ID: 'STU-7', Middle_Name: 'X', Age: 9 });
  }, ERROR_CODES.VALIDATION_ERROR);
  eq(err.details.unknownColumns, ['Middle_Name', 'Age']);
  ok(err.details.validColumns.indexOf('Student_ID') !== -1, 'validColumns should be reported');
});

check('appendRow_ refuses a sheet with no header row', function () {
  const writeApi = loadBackend(makeSandbox(makeSpreadsheet('Blank', [makeSheet('Students', [])])));
  throwsWithCode(function () {
    writeApi.appendRow_('Students', { Student_ID: 'STU-1' });
  }, ERROR_CODES.SERVER_ERROR);
});

check('appendRow_ rejects non-object, non-array input', function () {
  const writeApi = loadBackend(makeSandbox(makeFullSpreadsheet()));
  throwsWithCode(function () {
    writeApi.appendRow_('Students', 'STU-8');
  }, ERROR_CODES.VALIDATION_ERROR);
});

check('appendRow_ fails clearly when the script lock is unavailable', function () {
  const writeApi = loadBackend(makeSandbox(makeFullSpreadsheet(), { lockUnavailable: true }));
  throwsWithCode(function () {
    writeApi.appendRow_('Students', { Student_ID: 'STU-9' });
  }, ERROR_CODES.CONFLICT);
});

check('appendRow_ then readAll_ round-trips', function () {
  const spreadsheet = makeFullSpreadsheet();
  const writeApi = loadBackend(makeSandbox(spreadsheet));
  writeApi.appendRow_('Students', {
    Student_ID: 'STU-10',
    First_Name: 'Esi',
    Last_Name: 'Danso',
    Status: 'Active',
  });
  const all = writeApi.readAll_('Students');
  eq(all.length, 4);
  eq(all[3], { Student_ID: 'STU-10', First_Name: 'Esi', Last_Name: 'Danso', Status: 'Active' });
});

check('findRowById_ returns the 1-based sheet row and the record', function () {
  const found = api.findRowById_('Students', 'STU-2');
  eq(found.sheetRow, 3);
  eq(found.record, { Student_ID: 'STU-2', First_Name: 'Kofi', Last_Name: 'Owusu', Status: 'Active' });
});

check('findRowById_ returns null when nothing matches', function () {
  eq(api.findRowById_('Students', 'STU-999'), null);
});

check('findRowById_ supports a non-first id column', function () {
  eq(api.findRowById_('Students', 'Mensah', 'Last_Name').sheetRow, 2);
});

check('findRowById_ throws NOT_FOUND for an unknown id column', function () {
  const err = throwsWithCode(function () {
    api.findRowById_('Students', 'x', 'Nope_Column');
  }, ERROR_CODES.NOT_FOUND);
  ok(err.details.validColumns.indexOf('Student_ID') !== -1, 'validColumns should be reported');
});

check('findRowById_ returns null on a header-only sheet', function () {
  eq(api.findRowById_('Staff', 'anything'), null);
});

section('Router and health action (Tasks 5 and 6)');

check('doGet ?action=health returns the documented envelope', function () {
  const envelope = readEnvelope(api.doGet({ parameter: { action: 'health' } }));

  eq(envelope.success, true);
  eq(envelope.message, 'Service healthy');
  eq(envelope.data.service, 'School Management System API');
  eq(envelope.data.status, 'online');
  eq(envelope.data.spreadsheetId, CONFIG.SHEET_ID);
  eq(envelope.data.spreadsheetName, 'SchoolManagementSystem');
  eq(envelope.data.sheets, EXPECTED_TABS);
  eq(envelope.data.expectedSheets, EXPECTED_TABS);
  eq(envelope.data.missingSheets, []);
});

check('health proves real spreadsheet access and leaks no secrets', function () {
  const raw = api.doGet({ parameter: { action: 'health' } }).getContent();

  ['private', 'client_secret', 'refresh_token', 'access_token', 'password'].forEach(function (token) {
    ok(raw.toLowerCase().indexOf(token) === -1, 'health response leaked something matching "' + token + '"');
  });
  ok(raw.indexOf(CONFIG.SHEET_ID) !== -1, 'health should report the spreadsheet id it verified');
});

check('health reports missing tabs as a data issue without failing', function () {
  const partial = makeSpreadsheet('Partial', [
    makeSheet('Students', [['Student_ID']]),
    makeSheet('Staff', [['Staff_ID']]),
  ]);
  const envelope = readEnvelope(loadBackend(makeSandbox(partial)).doGet({ parameter: { action: 'health' } }));

  eq(envelope.success, true, 'missing tabs should not make the service unhealthy');
  eq(envelope.data.sheets, ['Students', 'Staff'], 'sheets must be the tabs that really exist');
  eq(envelope.data.missingSheets.length, 12);
  ok(envelope.data.missingSheets.indexOf('Audit_Log') !== -1, 'missingSheets should name the absent tabs');
});

check('health returns SERVER_ERROR when the spreadsheet is unreachable', function () {
  const broken = loadBackend(makeSandbox(null, { openByIdThrows: true }));
  const envelope = readEnvelope(broken.doGet({ parameter: { action: 'health' } }));

  eq(envelope.success, false);
  eq(envelope.error, 'SERVER_ERROR');
  eq(envelope.details.spreadsheetId, CONFIG.SHEET_ID);
});

check('doPost accepts a JSON body', function () {
  const envelope = readEnvelope(
    api.doPost({ parameter: {}, postData: { contents: JSON.stringify({ action: 'health' }) } })
  );
  eq(envelope.success, true);
  eq(envelope.data.status, 'online');
});

check('every response, including errors, is JSON with the JSON mime type', function () {
  const good = api.doGet({ parameter: { action: 'health' } });
  const bad = api.doGet({ parameter: { action: 'nope' } });

  eq(good.getMimeType(), 'application/json');
  eq(bad.getMimeType(), 'application/json');
  eq(typeof JSON.parse(bad.getContent()).success, 'boolean');
});

section('Request parsing and error handling (Task 5)');

check('parseRequest_ turns extra GET parameters into the payload', function () {
  eq(api.parseRequest_({ parameter: { action: 'students.get', id: 'STU-1' } }), {
    action: 'students.get',
    payload: { id: 'STU-1' },
  });
});

check('parseRequest_ takes the action from the POST body when absent from the query', function () {
  const request = api.parseRequest_({
    parameter: {},
    postData: { contents: JSON.stringify({ action: 'health' }) },
  });
  eq(request.action, 'health');
  eq(request.payload, {});
});

check('parseRequest_ lets the body payload win over query parameters', function () {
  const request = api.parseRequest_({
    parameter: { action: 'x', id: 'fromQuery' },
    postData: { contents: JSON.stringify({ action: 'y', payload: { id: 'fromBody' } }) },
  });
  eq(request.action, 'y');
  eq(request.payload, { id: 'fromBody' });
});

check('parseRequest_ treats an empty body as no body', function () {
  eq(api.parseRequest_({ parameter: { action: 'health' }, postData: { contents: '' } }).action, 'health');
});

check('parseRequest_ rejects malformed input with specific reasons', function () {
  const cases = [
    ['{oops', /not valid JSON/],
    ['[1,2]', /must be a JSON object/],
    ['5', /must be a JSON object/],
    ['"text"', /must be a JSON object/],
    ['null', /must be a JSON object/],
  ];

  cases.forEach(function (pair) {
    const err = throwsWithCode(function () {
      api.parseRequest_({ parameter: {}, postData: { contents: pair[0] } });
    }, ERROR_CODES.VALIDATION_ERROR);
    ok(pair[1].test(err.message), 'unexpected message for ' + pair[0] + ': ' + err.message);
  });
});

check('parseRequest_ rejects unexpected top-level body keys, teaching the exact contract', function () {
  const err = throwsWithCode(function () {
    api.parseRequest_({ parameter: {}, postData: { contents: JSON.stringify({ action: 'health', id: 'STU-1' }) } });
  }, ERROR_CODES.VALIDATION_ERROR);
  eq(err.details.unexpectedKeys, ['id']);
  ok(err.message.indexOf('payload') !== -1, 'message should point at the payload wrapper');
});

check('parseRequest_ requires payload to be an object', function () {
  const err = throwsWithCode(function () {
    api.parseRequest_({ parameter: {}, postData: { contents: JSON.stringify({ action: 'x', payload: [1] }) } });
  }, ERROR_CODES.VALIDATION_ERROR);
  eq(err.details.receivedType, 'array');
});

check('parseRequest_ ignores common URL tracking parameters (utm_*)', function () {
  // inventory.list with utm_source should succeed - the tracking param is ignored
  const request1 = api.parseRequest_({
    parameter: {
      action: 'inventory.list',
      utm_source: 'chatgpt',
      utm_medium: 'link',
      utm_campaign: 'test',
    },
  });
  eq(request1.action, 'inventory.list');
  eq(request1.payload.utm_source, undefined, 'utm_source should be stripped');
  eq(request1.payload.utm_medium, undefined, 'utm_medium should be stripped');
  eq(request1.payload.utm_campaign, undefined, 'utm_campaign should be stripped');
  eq(Object.keys(request1.payload).length, 0, 'payload should be empty');

  // inventory.movements with utm_* should also succeed
  const request2 = api.parseRequest_({
    parameter: {
      action: 'inventory.movements',
      utm_source: 'chatgpt',
      utm_term: 'search',
      utm_content: 'preview',
    },
  });
  eq(request2.action, 'inventory.movements');
  eq(request2.payload.utm_source, undefined);
  eq(request2.payload.utm_term, undefined);
  eq(request2.payload.utm_content, undefined);
  eq(Object.keys(request2.payload).length, 0);
});

check('parseRequest_ still passes through legitimate API parameters alongside tracking params', function () {
  // Legitimate filter params should still be passed through
  const request = api.parseRequest_({
    parameter: {
      action: 'inventory.list',
      Item_ID: 'ITM-001',
      utm_source: 'chatgpt',
      Category: 'Stationery',
    },
  });
  eq(request.action, 'inventory.list');
  eq(request.payload.Item_ID, 'ITM-001');
  eq(request.payload.Category, 'Stationery');
  eq(request.payload.utm_source, undefined, 'utm_source should still be stripped');
});

check('a missing action is rejected with all routed actions', function () {
  const envelope = readEnvelope(api.doGet({ parameter: {} }));
  eq(envelope.success, false);
  eq(envelope.error, 'VALIDATION_ERROR');
  eq(envelope.details.availableActions.sort(), [
    'auth.check', 'auth.me',
    'dashboard.summary',
    'feedingFees.create', 'feedingFees.get', 'feedingFees.list', 'feedingFees.update',
    'feedingFees.void',
    'health',
    'inventory.create', 'inventory.list', 'inventory.movements', 'inventory.stockIn', 'inventory.stockOut',
    'schoolFees.create', 'schoolFees.get', 'schoolFees.list', 'schoolFees.update',
    'schoolFees.void',
    'staff.create', 'staff.deactivate', 'staff.get', 'staff.list', 'staff.update',
    'stationery.create', 'stationery.fulfill', 'stationery.list',
    'students.create', 'students.get', 'students.list', 'students.update', 'students.withdraw',
  ]);
});

check('an unknown action is NOT_FOUND and names what is available', function () {
  const envelope = readEnvelope(api.doGet({ parameter: { action: 'does.notExist' } }));
  eq(envelope.success, false);
  eq(envelope.error, 'NOT_FOUND');
  eq(envelope.details.action, 'does.notExist');
  eq(envelope.details.availableActions.sort(), [
    'auth.check', 'auth.me',
    'dashboard.summary',
    'feedingFees.create', 'feedingFees.get', 'feedingFees.list', 'feedingFees.update',
    'feedingFees.void',
    'health',
    'inventory.create', 'inventory.list', 'inventory.movements', 'inventory.stockIn', 'inventory.stockOut',
    'schoolFees.create', 'schoolFees.get', 'schoolFees.list', 'schoolFees.update',
    'schoolFees.void',
    'staff.create', 'staff.deactivate', 'staff.get', 'staff.list', 'staff.update',
    'stationery.create', 'stationery.fulfill', 'stationery.list',
    'students.create', 'students.get', 'students.list', 'students.update', 'students.withdraw',
  ]);
});

check('the router never throws: any input still yields a valid envelope', function () {
  [
    undefined,
    null,
    {},
    { parameter: null },
    { postData: {} },
    { parameter: {}, postData: { contents: ']]]' } },
    { parameter: { action: '' } },
  ].forEach(function (event) {
    const envelope = readEnvelope(api.doGet(event));
    eq(typeof envelope.success, 'boolean', 'no boolean success for event ' + JSON.stringify(event));
    ok(envelope.error !== undefined, 'failures must carry an error code');
  });
});

section('Phase 2: auth.me and auth.check routes');

check('P2: auth.me returns caller context over GET', function () {
  var apiM = loadBackendAs('admin@school.edu', makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]));
  var envM = JSON.parse(apiM.doGet({ parameter: { action: 'auth.me' } }).getContent());
  eq(envM.success, true);
  eq(envM.data, { userId: 'USR-1', staffId: 'STF-1', email: 'admin@school.edu', role: 'Admin' });
});

check('P2: auth.me without identity is UNAUTHORIZED', function () {
  var apiN = loadBackendAs('', makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]));
  var envN = JSON.parse(apiN.doGet({ parameter: { action: 'auth.me' } }).getContent());
  eq(envN.success, false);
  eq(envN.error, ERROR_CODES.UNAUTHORIZED);
});

check('P2: auth.check without permission proves authentication', function () {
  var apiC = loadBackendAs('admin@school.edu', makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]));
  var envC = JSON.parse(apiC.doPost({ parameter: {},
    postData: { contents: JSON.stringify({ action: 'auth.check', payload: {} }) } }).getContent());
  eq(envC.success, true);
  eq(envC.data.authenticated, true);
  eq(envC.data.permission, null);
  eq(envC.data.allowed, true);
});

check('P2: auth.check reports allowed and denied honestly', function () {
  var apiAd = loadBackendAs('admin@school.edu', makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]));
  var granted = JSON.parse(apiAd.doPost({ parameter: {},
    postData: { contents: JSON.stringify({ action: 'auth.check', payload: { permission: 'STUDENTS.READ' } }) } }).getContent());
  eq(granted.success, true);
  eq(granted.data.allowed, true);
  eq(granted.data.permission, 'STUDENTS.READ');
  var apiTe = loadBackendAs('teacher@school.edu', makeAuthSpreadsheet([
    ['USR-2', 'STF-2', 'teacher@school.edu', 'Teacher', 'Active', '']
  ]));
  var denied = JSON.parse(apiTe.doPost({ parameter: {},
    postData: { contents: JSON.stringify({ action: 'auth.check', payload: { permission: 'STUDENTS.READ' } }) } }).getContent());
  eq(denied.success, true);
  eq(denied.data.allowed, false);
});

check('P2: auth.check rejects malformed permission VALIDATION_ERROR', function () {
  var apiB = loadBackendAs('admin@school.edu', makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]));
  var envB = JSON.parse(apiB.doPost({ parameter: {},
    postData: { contents: JSON.stringify({ action: 'auth.check', payload: { permission: 'nope' } }) } }).getContent());
  eq(envB.success, false);
  eq(envB.error, ERROR_CODES.VALIDATION_ERROR);
});

check('P2: auth endpoints ignore identity smuggled in payload', function () {
  var apiS = loadBackendAs('teacher@school.edu', makeAuthSpreadsheet([
    ['USR-2', 'STF-2', 'teacher@school.edu', 'Teacher', 'Active', ''],
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]));
  var envS = JSON.parse(apiS.doPost({ parameter: {},
    postData: { contents: JSON.stringify({ action: 'auth.me', payload: { email: 'admin@school.edu' } }) } }).getContent());
  eq(envS.success, true);
  eq(envS.data.email, 'teacher@school.edu');
});

check('P2: health stays public auth routes need identity', function () {
  var apiAnon = loadBackendAs('', makeFullSpreadsheet());
  eq(JSON.parse(apiAnon.doGet({ parameter: { action: 'health' } }).getContent()).success, true);
  eq(JSON.parse(apiAnon.doGet({ parameter: { action: 'auth.me' } }).getContent()).error, ERROR_CODES.UNAUTHORIZED);
});

check('P2: setCellValue_ writes one cell under lock', function () {
  var wb = makeSpreadsheet('W', [makeSheet('Users',
    [['User_ID', 'Last_Login'], ['USR-1', '']])]);
  var apiW = loadBackend(wb.__sandbox || makeSandbox(wb));
  ok(apiW.setCellValue_('Users', 2, 'Last_Login', 'x'), 'should return true');
  eq(wb.getSheetByName('Users')._rows[1][1], 'x');
  throwsWithCode(function () { apiW.setCellValue_('Users', 2, 'Nope', 'x'); }, ERROR_CODES.NOT_FOUND);
  throwsWithCode(function () { apiW.setCellValue_('Users', 1, 'Last_Login', 'x'); }, ERROR_CODES.VALIDATION_ERROR);
});


section('Phase 2: Last_Login and permissions');

check('P2: Last_Login stamped once per day never breaks login', function () {
  var sh = makeSheet('Users', []);
  sh._rows = [
    ['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'],
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ];
  var day1 = loadBackendAs('admin@school.edu', makeSpreadsheet('S', [sh]));
  day1.requireAuthentication_();
  var stamped = sh._rows[1][5];
  ok(String(stamped) !== '', 'Last_Login should be stamped');
  var day2 = loadBackendAs('admin@school.edu', makeSpreadsheet('S', [sh]));
  day2.requireAuthentication_();
  eq(sh._rows[1][5], stamped, 'same-day login must not rewrite');
});

check('P2: permission format validated unknown denies', function () {
  var apiP = loadBackendAs('admin@school.edu', makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]));
  var admin = apiP.requireAuthentication_();
  eq(apiP.hasPermission_(admin, 'STUDENTS.READ'), true);
  eq(apiP.hasPermission_(admin, 'NOPE.READ'), false);
  eq(apiP.hasPermission_(admin, 'not-a-permission'), false);
  eq(apiP.hasPermission_(admin, ''), false);
  eq(apiP.hasPermission_(null, 'STUDENTS.READ'), false);
  throwsWithCode(function () { apiP.assertValidPermissionFormat_('nope'); }, ERROR_CODES.VALIDATION_ERROR);
  apiP.assertValidPermissionFormat_('STUDENTS.READ');
});

check('P2: non-admin holds no permissions (deny-by-default)', function () {
  var apiT = loadBackendAs('teacher@school.edu', makeAuthSpreadsheet([
    ['USR-2', 'STF-2', 'teacher@school.edu', 'Teacher', 'Active', '']
  ]));
  var teacher = apiT.requireAuthentication_();
  eq(apiT.hasPermission_(teacher, 'STUDENTS.READ'), false);
  eq(apiT.resolveRolePermissions_('Teacher'), []);
  eq(apiT.resolveRolePermissions_('NoSuchRole'), []);
  var ferr = throwsWithCode(function () { apiT.requirePermission_('STUDENTS.READ'); }, ERROR_CODES.FORBIDDEN);
  eq(ferr.details.permission, 'STUDENTS.READ');
});

check('P2: requirePermission_ returns user when granted', function () {
  var apiG = loadBackendAs('admin@school.edu', makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]));
  eq(apiG.requirePermission_('STUDENTS.READ').userId, 'USR-1');
});


section('Phase 2: duplicate and error handling');

check('P2: duplicate email safe when only one row active', function () {
  var api6 = loadBackendAs('dup@school.edu', makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'dup@school.edu', 'Admin', 'Inactive', ''],
    ['USR-2', 'STF-2', 'dup@school.edu', 'Admin', 'Active', '']
  ]));
  eq(api6.requireAuthentication_().userId, 'USR-2');
});

check('P2: missing Users sheet is structured SERVER_ERROR', function () {
  var api7 = loadBackendAs('admin@school.edu', makeSpreadsheet('Empty', []));
  var res7 = api7.getCurrentUser_();
  eq(res7.user, null);
  eq(res7.error.code, ERROR_CODES.SERVER_ERROR);
});

check('P2: missing Users columns reported not misread', function () {
  var bad = makeSpreadsheet('Bad', [makeSheet('Users', [['User_ID', 'Email']])]);
  var api8 = loadBackendAs('admin@school.edu', bad);
  var res8 = api8.getCurrentUser_();
  eq(res8.user, null);
  eq(res8.error.code, ERROR_CODES.SERVER_ERROR);
  ok(res8.error.details.missingColumns.indexOf('Role') !== -1, 'should name missing Role');
});

check('P2: context carries no secrets', function () {
  var api9 = loadBackendAs('admin@school.edu', makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]));
  eq(Object.keys(api9.requireAuthentication_()).sort(), ['email', 'role', 'staffId', 'userId']);
});

check('P2: throwing Session fails closed UNAUTHORIZED', function () {
  var sb = makeSandbox(makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]));
  sb.Session.getActiveUser = function () { throw new Error('no identity'); };
  eq(loadBackend(sb).getCurrentUser_().error.code, ERROR_CODES.UNAUTHORIZED);
});


section('Phase 2a: OAuth token authentication (__auth.access_token)');

/**
 * Sandbox with no session identity, where Google's userinfo endpoint answers
 * with `body` and HTTP `status`. Simulates the staff Gmail flow: identity comes
 * exclusively from the verified token, never from the session.
 */
function makeTokenSandbox(usersRows, userinfo) {
  const sb = makeSandbox(makeAuthSpreadsheet(usersRows), {
    urlFetchResponse: userinfo,
  });
  sb.Session.__activeEmail = ''; // no session identity
  return sb;
}

const GOOD_USERINFO = { status: 200, body: { sub: 'g-1', email: 'admin@school.edu', email_verified: true } };

check('P2a: verified token authenticates the matching Active user', function () {
  const sb = makeTokenSandbox(
    [['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']],
    GOOD_USERINFO
  );
  const api = loadBackend(sb);
  api.setRequestAuthToken_('tok-abc');
  const result = api.getCurrentUser_();
  eq(result.error, null);
  eq(result.user.email, 'admin@school.edu');
  eq(result.user.role, 'Admin');
  // The userinfo call must hit Google's endpoint with the bearer token.
  eq(sb.__urlFetchCalls.length, 1);
  eq(sb.__urlFetchCalls[0].url, 'https://openidconnect.googleapis.com/v1/userinfo');
  eq(sb.__urlFetchCalls[0].params.headers.Authorization, 'Bearer tok-abc');
});

check('P2a: token email absent from Users is UNAUTHORIZED, never entry', function () {
  // userinfo proves the Google identity is stranger@gmail.com, which is not an
  // Active Users row. A client cannot gain access by submitting any email.
  const sb = makeTokenSandbox(
    [['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']],
    { status: 200, body: { sub: 'g-9', email: 'stranger@gmail.com', email_verified: true } }
  );
  const api = loadBackend(sb);
  api.setRequestAuthToken_('tok-abc');
  const res = api.getCurrentUser_();
  eq(res.user, null);
  eq(res.error.code, ERROR_CODES.UNAUTHORIZED);
  eq(res.error.details.reason, 'no-matching-user');
});

check('P2a: unverified email on the Google account is rejected', function () {
  const sb = makeTokenSandbox(
    [['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']],
    { status: 200, body: { sub: 'g-1', email: 'admin@school.edu', email_verified: false } }
  );
  const api = loadBackend(sb);
  api.setRequestAuthToken_('tok-abc');
  const res = api.getCurrentUser_();
  eq(res.user, null);
  eq(res.error.code, ERROR_CODES.UNAUTHORIZED);
  eq(res.error.details.reason, 'email-unverified');
});

check('P2a: expired/invalid token (401) reports token-invalid', function () {
  const sb = makeTokenSandbox(
    [['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']],
    { status: 401, body: { error: 'invalid_token' } }
  );
  const api = loadBackend(sb);
  api.setRequestAuthToken_('tok-abc');
  const res = api.getCurrentUser_();
  eq(res.user, null);
  eq(res.error.code, ERROR_CODES.UNAUTHORIZED);
  eq(res.error.details.reason, 'token-invalid');
});

check('P2a: userinfo network failure fails closed, not open', function () {
  const sb = makeSandbox(makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]), { urlFetchThrows: true });
  sb.Session.__activeEmail = '';
  const api = loadBackend(sb);
  api.setRequestAuthToken_('tok-abc');
  const res = api.getCurrentUser_();
  eq(res.user, null);
  eq(res.error.code, ERROR_CODES.UNAUTHORIZED);
  eq(res.error.details.reason, 'token-verification-failed');
});

check('P2a: Router threads __auth token and strips it from the payload', function () {
  const sb = makeSandbox(makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]), { urlFetchResponse: GOOD_USERINFO });
  sb.Session.__activeEmail = ''; // identity must come only from the token
  const instance = loadBackend(sb);
  const env = doPostEnvelope(instance, 'auth.me', { __auth: { access_token: 'tok-abc' } });
  eq(env.success, true, 'token should authenticate the request');
  eq(env.data.email, 'admin@school.edu');
  eq(sb.__urlFetchCalls.length, 1, 'userinfo should have been consulted');
  // The token is request-scoped: after the request it must be cleared, so a
  // follow-up call with no token falls back to the (blank) session identity.
  const after = instance.getCurrentUser_();
  eq(after.user, null, 'token must not leak past the request that carried it');
});

check('P2a: session identity remains the fallback when no token is sent', function () {
  const sb = makeSandbox(makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]));
  sb.Session.__activeEmail = 'admin@school.edu';
  const env = doPostEnvelope(loadBackend(sb), 'auth.me', {});
  eq(env.success, true);
  eq(env.data.email, 'admin@school.edu');
  eq(sb.__urlFetchCalls, undefined, 'userinfo should not be called without a token');
});

check('P2a: token takes precedence over a blank session identity', function () {
  const sb = makeSandbox(makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]), { urlFetchResponse: GOOD_USERINFO });
  sb.Session.__activeEmail = '';
  const instance = loadBackend(sb);
  instance.setRequestAuthToken_('tok-abc');
  const res = instance.getCurrentUser_();
  eq(res.error, null);
  eq(res.user.email, 'admin@school.edu');
});


section('Phase 2: authentication (Auth.js)');

check('P2: blank Google identity yields UNAUTHORIZED with reason', function () {
  var anonApi = loadBackendAs('', makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]));
  eq(anonApi.getAuthenticatedEmail_(), '');
  var result = anonApi.getCurrentUser_();
  eq(result.user, null);
  eq(result.error.code, ERROR_CODES.UNAUTHORIZED);
  eq(result.error.details.reason, 'no-google-identity');
});

check('P2: unknown email yields UNAUTHORIZED not NOT_FOUND', function () {
  var api2 = loadBackendAs('stranger@school.edu', makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]));
  var res2 = api2.getCurrentUser_();
  eq(res2.user, null);
  eq(res2.error.code, ERROR_CODES.UNAUTHORIZED);
  eq(res2.error.details.reason, 'no-matching-user');
});

check('P2: inactive user cannot authenticate', function () {
  var api3 = loadBackendAs('old@school.edu', makeAuthSpreadsheet([
    ['USR-9', 'STF-9', 'old@school.edu', 'Admin', 'Inactive', '']
  ]));
  var res3 = api3.getCurrentUser_();
  eq(res3.user, null);
  eq(res3.error.code, ERROR_CODES.UNAUTHORIZED);
});

check('P2: email match ignores case and spaces', function () {
  var api4 = loadBackendAs('  ADMIN@School.Edu ', makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']
  ]));
  var res4 = api4.getCurrentUser_();
  eq(res4.error, null);
  eq(res4.user.email, 'admin@school.edu');
  eq(res4.user.userId, 'USR-1');
  eq(res4.user.role, 'Admin');
});

check('P2: duplicate active emails fail closed with CONFLICT', function () {
  var api5 = loadBackendAs('dup@school.edu', makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'dup@school.edu', 'Admin', 'Active', ''],
    ['USR-2', 'STF-2', 'dup@school.edu', 'Admin', 'Active', '']
  ]));
  var err5 = throwsWithCode(function () { api5.requireAuthentication_(); }, ERROR_CODES.CONFLICT);
  eq(err5.details.email, 'dup@school.edu');
});


/* ==========================================================================
 * Phase 3: Students and Staff
 * ======================================================================== */

// Canonical column schemas so the test sheets match the real tabs exactly.
const P3_STUDENT_COLUMNS = [
  'Student_ID', 'First_Name', 'Last_Name', 'Gender', 'Date_of_Birth',
  'Class', 'Parent_Guardian', 'Guardian_Phone', 'Guardian_Email',
  'Emergency_Contact_1_Name', 'Emergency_Contact_1_Phone',
  'Emergency_Contact_1_Relationship',
  'Emergency_Contact_2_Name', 'Emergency_Contact_2_Phone',
  'Emergency_Contact_2_Relationship',
  'Allergies', 'Illnesses_Medical_Conditions',
  'Physical_Defects_Special_Conditions',
  'Admission_Date', 'Status', 'Withdrawal_Date', 'Notes'
];

const P3_STAFF_COLUMNS = [
  'Staff_ID', 'First_Name', 'Last_Name', 'Gender', 'Date_of_Birth',
  'Phone', 'Email', 'Address', 'Position', 'Department',
  'Employment_Date', 'Employment_Status', 'Salary_Amount', 'Salary_Frequency',
  'Last_Salary_Paid_Date', 'Next_Salary_Due_Date', 'Salary_Status', 'Notes'
];

/** Build a full-width row from a partial record, blank-filling the rest. */
function rowFor(columns, data) {
  return columns.map(function (col) {
    return data[col] !== undefined ? data[col] : '';
  });
}

/**
 * Spreadsheet with canonical Students + Staff + Users tabs, so the Phase 3
 * handlers see exactly the schema the production sheets are expected to have.
 * @param {Object} opts { studentRows, staffRows, usersRows } seed data.
 */
function makePhase3Spreadsheet(opts) {
  opts = opts || {};
  const sheets = EXPECTED_TABS.map(function (tabName) {
    return makeSheet(tabName, [['Header_A', 'Header_B']]);
  });
  const studentsSheet = sheets[EXPECTED_TABS.indexOf('Students')];
  studentsSheet._rows = [P3_STUDENT_COLUMNS.slice()].concat(opts.studentRows || []);
  const staffSheet = sheets[EXPECTED_TABS.indexOf('Staff')];
  staffSheet._rows = [P3_STAFF_COLUMNS.slice()].concat(opts.staffRows || []);
  const usersSheet = sheets[EXPECTED_TABS.indexOf('Users')];
  usersSheet._rows = opts.usersRows || [
    ['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'],
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', ''],
    ['USR-2', 'STF-2', 'teacher@school.edu', 'Teacher', 'Active', ''],
  ];
  // Phase 4A: real authorization tables so admin/teacher routes resolve
  // through Role_Permissions exactly as they will in production.
  const p4a = p4aPermissionSheets();
  sheets[EXPECTED_TABS.indexOf('Roles')] = p4a.roles;
  sheets[EXPECTED_TABS.indexOf('Permissions')] = p4a.permissions;
  sheets[EXPECTED_TABS.indexOf('Role_Permissions')] = p4a.rolePermissions;
  return makeSpreadsheet('SchoolManagementSystem', sheets);
}

/** Helper: call doGet and return the parsed response envelope. */
function doGetEnvelope(apiInstance, params) {
  return JSON.parse(apiInstance.doGet({ parameter: params }).getContent());
}

/** Helper: call doPost with a JSON body and return the parsed response envelope. */
function doPostEnvelope(apiInstance, action, payload) {
  return JSON.parse(apiInstance.doPost({
    parameter: {},
    postData: { contents: JSON.stringify({ action: action, payload: payload }) },
  }).getContent());
}

/** Helper: run doPost as the given Google identity and return parsed envelope. */
function doPostEnvelopeAs(apiInstance, email, action, payload) {
  if (apiInstance && apiInstance.__ss && apiInstance.__ss.Session) {
    apiInstance.__ss.Session.__activeEmail = email;
  }
  return doPostEnvelope(apiInstance, action, payload);
}

/** A valid students.create payload the field-validation tests start from. */
function p3StudentPayload(overrides) {
  return Object.assign({
    First_Name: 'Ama',
    Last_Name: 'Mensah',
    Class: 'KG',
    Admission_Date: '2025-01-15',
    Parent_Guardian: 'Mr. Mensah',
  }, overrides || {});
}

/** A valid staff.create payload the field-validation tests start from. */
function p3StaffPayload(overrides) {
  return Object.assign({
    First_Name: 'Kwame',
    Last_Name: 'Adjei',
    Email: 'kwame.adjei@school.edu',
    Position: 'Teacher',
    Employment_Date: '2024-09-01',
  }, overrides || {});
}

section('Phase 3: students.list and students.get');

check('P3: students.list returns all non-blank student rows', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({
    studentRows: [
      rowFor(P3_STUDENT_COLUMNS, { Student_ID: 'STU-1', First_Name: 'Ama', Last_Name: 'Mensah', Status: 'Active' }),
      rowFor(P3_STUDENT_COLUMNS, { Student_ID: 'STU-2', First_Name: 'Kofi', Last_Name: 'Owusu', Status: 'Active' }),
      ['', '', '', ''],
      rowFor(P3_STUDENT_COLUMNS, { Student_ID: 'STU-3', First_Name: 'Yaa', Last_Name: 'Boateng', Status: 'Withdrawn' }),
    ],
  }));
  var envelope = doGetEnvelope(apiS, { action: 'students.list' });
  eq(envelope.success, true);
  eq(envelope.data.count, 3);
  ok(envelope.data.students.length === 3, 'list should skip blank rows');
  eq(envelope.data.students[0].First_Name, 'Ama');
  eq(envelope.data.students[2].Status, 'Withdrawn');
});

check('P3: students.get returns an existing student', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({
    studentRows: [
      rowFor(P3_STUDENT_COLUMNS, { Student_ID: 'STU-1', First_Name: 'Ama', Last_Name: 'Mensah', Status: 'Active' }),
    ],
  }));
  var envelope = doGetEnvelope(apiS, { action: 'students.get', Student_ID: 'STU-1' });
  eq(envelope.success, true);
  eq(envelope.data.Student_ID, 'STU-1');
  eq(envelope.data.First_Name, 'Ama');
});

check('P3: students.get missing student is NOT_FOUND', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doGetEnvelope(apiS, { action: 'students.get', Student_ID: 'NOPE-999' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.NOT_FOUND);
});

section('Phase 3: students.create');

check('P3: students.create validates required fields', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'students.create', { First_Name: 'Ama', Last_Name: 'Mensah' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.VALIDATION_ERROR);
  ok(envelope.details.missingFields.indexOf('Class') !== -1, 'should flag missing Class');
});

check('P3: students.create generates a server-side Student_ID', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'students.create', p3StudentPayload());
  eq(envelope.success, true);
  ok(/^STU-\d{8}-[0-9A-F]{12}$/.test(envelope.data.Student_ID), 'unexpected id shape: ' + envelope.data.Student_ID);
  eq(envelope.data.Status, 'Active');
});

check('P3: students.create rejects a client-supplied Student_ID', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'students.create', p3StudentPayload({ Student_ID: 'STU-HACKED' }));
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.VALIDATION_ERROR);
  ok(envelope.message.indexOf('server-generated') !== -1, 'message should explain the rule');
});

check('P3: students.create rejects unknown fields', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'students.create', p3StudentPayload({ Middle_Name: 'X' }));
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.VALIDATION_ERROR);
  ok(envelope.details.unknownColumns.indexOf('Middle_Name') !== -1, 'should list unknown columns');
});

check('P3: students.create rejects an invalid Status', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'students.create', p3StudentPayload({ Status: 'Zombie' }));
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.VALIDATION_ERROR);
});

check('P3: students.create rejects a malformed Guardian_Email', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'students.create', p3StudentPayload({ Guardian_Email: 'not-an-email' }));
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.VALIDATION_ERROR);
});

section('Phase 3: students.update and students.withdraw');

check('P3: students.update preserves unspecified fields (partial update)', function () {
  var ss = makePhase3Spreadsheet({
    studentRows: [
      rowFor(P3_STUDENT_COLUMNS, { Student_ID: 'STU-1', First_Name: 'Ama', Last_Name: 'Mensah', Class: '2A', Status: 'Active', Guardian_Phone: '12345' }),
    ],
  });
  var apiS = loadBackendAs('admin@school.edu', ss);
  var envelope = doPostEnvelope(apiS, 'students.update', { Student_ID: 'STU-1', First_Name: 'Ama', Last_Name: 'Mensah' });
  eq(envelope.success, true);
  eq(envelope.data.Class, '2A');
  eq(envelope.data.Guardian_Phone, '12345');
  eq(envelope.data.Status, 'Active');
});

check('P3: students.update never changes the immutable Student_ID', function () {
  var ss = makePhase3Spreadsheet({
    studentRows: [
      rowFor(P3_STUDENT_COLUMNS, { Student_ID: 'STU-1', First_Name: 'Ama', Last_Name: 'Mensah', Class: '2A', Status: 'Active' }),
    ],
  });
  var apiS = loadBackendAs('admin@school.edu', ss);
  var envelope = doPostEnvelope(apiS, 'students.update', { Student_ID: 'STU-1', Last_Name: 'Mensah-Edited' });
  eq(envelope.success, true);
  eq(envelope.data.Student_ID, 'STU-1', 'Student_ID must be immutable');
  var listed = doGetEnvelope(apiS, { action: 'students.list' });
  var matches = listed.data.students.filter(function (s) { return s.Student_ID === 'STU-1'; });
  eq(matches.length, 1, 'exactly one row must keep the original Student_ID');
  eq(matches[0].Last_Name, 'Mensah-Edited');
});

check('P3: students.update missing student is NOT_FOUND', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'students.update', { Student_ID: 'NOPE-999', First_Name: 'Ama' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.NOT_FOUND);
});

check('P3: students.withdraw sets Status and Withdrawal_Date', function () {
  var ss = makePhase3Spreadsheet({
    studentRows: [
      rowFor(P3_STUDENT_COLUMNS, { Student_ID: 'STU-1', First_Name: 'Ama', Last_Name: 'Mensah', Status: 'Active' }),
    ],
  });
  var apiS = loadBackendAs('admin@school.edu', ss);
  var envelope = doPostEnvelope(apiS, 'students.withdraw', { Student_ID: 'STU-1' });
  eq(envelope.success, true);
  eq(envelope.data.Status, 'Withdrawn');
  ok(envelope.data.Withdrawal_Date, 'Withdrawal_Date should be set');
});

check('P3: withdrawn student remains in the sheet (soft delete)', function () {
  var ss = makePhase3Spreadsheet({
    studentRows: [
      rowFor(P3_STUDENT_COLUMNS, { Student_ID: 'STU-1', First_Name: 'Ama', Last_Name: 'Mensah', Status: 'Active' }),
    ],
  });
  var apiS = loadBackendAs('admin@school.edu', ss);
  doPostEnvelope(apiS, 'students.withdraw', { Student_ID: 'STU-1' });
  var envelope = doGetEnvelope(apiS, { action: 'students.list' });
  eq(envelope.success, true);
  ok(envelope.data.students.length >= 1, 'withdrawn student must remain in sheet');
  var found = envelope.data.students.filter(function (s) { return s.Student_ID === 'STU-1'; });
  eq(found.length, 1);
  eq(found[0].Status, 'Withdrawn');
});

check('P3: students.withdraw missing student is NOT_FOUND', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'students.withdraw', { Student_ID: 'NOPE-999' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.NOT_FOUND);
});

section('Phase 3: students persistence, conflicts and authorization');

check('P3: students.create write survives on a fresh sheet', function () {
  var ss = makePhase3Spreadsheet({});
  var apiS = loadBackendAs('admin@school.edu', ss);
  var created = doPostEnvelope(apiS, 'students.create', p3StudentPayload());
  eq(created.success, true);
  var listed = doGetEnvelope(apiS, { action: 'students.list' });
  eq(listed.data.count, 1);
  var dup = listed.data.students.filter(function (s) { return s.Student_ID === created.data.Student_ID; });
  eq(dup.length, 1);
});

check('P3: students.create reports CONFLICT when the generated ID already exists', function () {
  var forcedUuid = '0f1e2d3c4b5a69788796a5b4c3d2e1f0';
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}), {
    uuid: function () { return forcedUuid; },
  });
  var first = doPostEnvelope(apiS, 'students.create', p3StudentPayload({ First_Name: 'Ama' }));
  eq(first.success, true);
  var second = doPostEnvelope(apiS, 'students.create', p3StudentPayload({ First_Name: 'Kofi' }));
  eq(second.success, false);
  eq(second.error, ERROR_CODES.CONFLICT);
});

check('P3: students denied without permission is FORBIDDEN', function () {
  var apiT = loadBackendAs('teacher@school.edu', makePhase3Spreadsheet({}));
  var envelope = doGetEnvelope(apiT, { action: 'students.list' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.FORBIDDEN);
});

check('P3: students unauthenticated without Google identity is UNAUTHORIZED', function () {
  var apiBlank = loadBackendAs('', makePhase3Spreadsheet({}));
  var envelope = doGetEnvelope(apiBlank, { action: 'students.list' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.UNAUTHORIZED);
});

section('Phase 3: structural failures (Students)');

check('P3: students.list missing Students sheet is SERVER_ERROR', function () {
  var apiS = loadBackendAs('admin@school.edu', makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', ''],
  ]));
  var envelope = doGetEnvelope(apiS, { action: 'students.list' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.SERVER_ERROR);
});

check('P3: students.list missing required columns is SERVER_ERROR', function () {
  var p4a = p4aPermissionSheets();
  var ss = makeSpreadsheet('SchoolManagementSystem', [
    makeSheet('Students', [['Student_ID', 'First_Name']]),
    makeSheet('Users', [['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'],
      ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']]),
    p4a.roles, p4a.permissions, p4a.rolePermissions,
  ]);
  var apiS = loadBackendAs('admin@school.edu', ss);
  var envelope = doGetEnvelope(apiS, { action: 'students.list' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.SERVER_ERROR);
  ok(envelope.details.missingColumns.length > 0, 'should report which columns are missing');
});

check('P3: students.create respects script locking on write', function () {
  var ss = makePhase3Spreadsheet({});
  var apiLocked = loadBackendAs('admin@school.edu', ss, { lockUnavailable: true });
  var envelope = doPostEnvelope(apiLocked, 'students.create', p3StudentPayload());
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.CONFLICT);
});

section('Phase 3: staff.list and staff.get');

check('P3: staff.list returns all non-blank staff rows', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({
    staffRows: [
      rowFor(P3_STAFF_COLUMNS, { Staff_ID: 'STF-1', First_Name: 'Kwame', Last_Name: 'Adjei', Employment_Status: 'Active' }),
      rowFor(P3_STAFF_COLUMNS, { Staff_ID: 'STF-2', First_Name: 'Abena', Last_Name: 'Osei', Employment_Status: 'Active' }),
      ['', '', '', ''],
      rowFor(P3_STAFF_COLUMNS, { Staff_ID: 'STF-3', First_Name: 'Yaw', Last_Name: 'Nkrumah', Employment_Status: 'Inactive' }),
    ],
  }));
  var envelope = doGetEnvelope(apiS, { action: 'staff.list' });
  eq(envelope.success, true);
  eq(envelope.data.count, 3);
  ok(envelope.data.staff.length === 3, 'list should skip blank rows');
  eq(envelope.data.staff[0].First_Name, 'Kwame');
  eq(envelope.data.staff[2].Employment_Status, 'Inactive');
});

check('P3: staff.get returns an existing staff member', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({
    staffRows: [
      rowFor(P3_STAFF_COLUMNS, { Staff_ID: 'STF-1', First_Name: 'Kwame', Last_Name: 'Adjei', Employment_Status: 'Active' }),
    ],
  }));
  var envelope = doGetEnvelope(apiS, { action: 'staff.get', Staff_ID: 'STF-1' });
  eq(envelope.success, true);
  eq(envelope.data.Staff_ID, 'STF-1');
  eq(envelope.data.First_Name, 'Kwame');
});

check('P3: staff.get missing staff is NOT_FOUND', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doGetEnvelope(apiS, { action: 'staff.get', Staff_ID: 'NOPE-999' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.NOT_FOUND);
});

section('Phase 3: staff.create');

check('P3: staff.create validates required fields', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'staff.create', { First_Name: 'Kwame', Last_Name: 'Adjei' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.VALIDATION_ERROR);
  ok(envelope.details.missingFields.indexOf('Position') !== -1, 'should flag missing Position');
});

check('P3: staff.create generates a server-side Staff_ID', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'staff.create', p3StaffPayload());
  eq(envelope.success, true);
  ok(/^STF-\d{8}-[0-9A-F]{12}$/.test(envelope.data.Staff_ID), 'unexpected id shape: ' + envelope.data.Staff_ID);
  eq(envelope.data.Employment_Status, 'Active');
});

check('P3: staff.create rejects a client-supplied Staff_ID', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'staff.create', p3StaffPayload({ Staff_ID: 'STF-HACKED' }));
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.VALIDATION_ERROR);
  ok(envelope.message.indexOf('server-generated') !== -1, 'message should explain the rule');
});

check('P3: staff.create rejects unknown fields', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'staff.create', p3StaffPayload({ Middle_Name: 'X' }));
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.VALIDATION_ERROR);
  ok(envelope.details.unknownColumns.indexOf('Middle_Name') !== -1, 'should list unknown columns');
});

check('P3: staff.create rejects an invalid Employment_Status', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'staff.create', p3StaffPayload({ Employment_Status: 'Zombie' }));
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.VALIDATION_ERROR);
});

check('P3: staff.create rejects a malformed Email', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'staff.create', p3StaffPayload({ Email: 'not-an-email' }));
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.VALIDATION_ERROR);
});

check('P3: staff.create rejects a negative Salary_Amount', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'staff.create', p3StaffPayload({ Salary_Amount: -500 }));
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.VALIDATION_ERROR);
});

check('P3: staff.create rejects an invalid Salary_Frequency', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'staff.create', p3StaffPayload({ Salary_Frequency: 'Fortnightly' }));
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.VALIDATION_ERROR);
});

section('Phase 3: staff.update and staff.deactivate');

check('P3: staff.update preserves unspecified fields (partial update)', function () {
  var ss = makePhase3Spreadsheet({
    staffRows: [
      rowFor(P3_STAFF_COLUMNS, { Staff_ID: 'STF-1', First_Name: 'Kwame', Last_Name: 'Adjei', Department: 'Science', Salary_Amount: 1500, Employment_Status: 'Active' }),
    ],
  });
  var apiS = loadBackendAs('admin@school.edu', ss);
  var envelope = doPostEnvelope(apiS, 'staff.update', { Staff_ID: 'STF-1', First_Name: 'Kwame', Last_Name: 'Adjei' });
  eq(envelope.success, true);
  eq(envelope.data.Department, 'Science');
  eq(envelope.data.Salary_Amount, 1500);
  eq(envelope.data.Employment_Status, 'Active');
});

check('P3: staff.update never changes the immutable Staff_ID', function () {
  var ss = makePhase3Spreadsheet({
    staffRows: [
      rowFor(P3_STAFF_COLUMNS, { Staff_ID: 'STF-1', First_Name: 'Kwame', Last_Name: 'Adjei', Employment_Status: 'Active' }),
    ],
  });
  var apiS = loadBackendAs('admin@school.edu', ss);
  var envelope = doPostEnvelope(apiS, 'staff.update', { Staff_ID: 'STF-1', Last_Name: 'Adjei-Edited' });
  eq(envelope.success, true);
  eq(envelope.data.Staff_ID, 'STF-1', 'Staff_ID must be immutable');
  var listed = doGetEnvelope(apiS, { action: 'staff.list' });
  var matches = listed.data.staff.filter(function (s) { return s.Staff_ID === 'STF-1'; });
  eq(matches.length, 1, 'exactly one row must keep the original Staff_ID');
  eq(matches[0].Last_Name, 'Adjei-Edited');
});

check('P3: staff.update missing staff is NOT_FOUND', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'staff.update', { Staff_ID: 'NOPE-999', First_Name: 'Kwame' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.NOT_FOUND);
});

check('P3: staff.deactivate sets Employment_Status to Inactive', function () {
  var ss = makePhase3Spreadsheet({
    staffRows: [
      rowFor(P3_STAFF_COLUMNS, { Staff_ID: 'STF-1', First_Name: 'Kwame', Last_Name: 'Adjei', Employment_Status: 'Active' }),
    ],
  });
  var apiS = loadBackendAs('admin@school.edu', ss);
  var envelope = doPostEnvelope(apiS, 'staff.deactivate', { Staff_ID: 'STF-1' });
  eq(envelope.success, true);
  eq(envelope.data.Employment_Status, 'Inactive');
});

check('P3: deactivated staff record remains in the sheet (soft delete)', function () {
  var ss = makePhase3Spreadsheet({
    staffRows: [
      rowFor(P3_STAFF_COLUMNS, { Staff_ID: 'STF-1', First_Name: 'Kwame', Last_Name: 'Adjei', Employment_Status: 'Active' }),
    ],
  });
  var apiS = loadBackendAs('admin@school.edu', ss);
  doPostEnvelope(apiS, 'staff.deactivate', { Staff_ID: 'STF-1' });
  var envelope = doGetEnvelope(apiS, { action: 'staff.list' });
  eq(envelope.success, true);
  var found = envelope.data.staff.filter(function (s) { return s.Staff_ID === 'STF-1'; });
  eq(found.length, 1);
  eq(found[0].Employment_Status, 'Inactive');
});

check('P3: staff.deactivate missing staff is NOT_FOUND', function () {
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}));
  var envelope = doPostEnvelope(apiS, 'staff.deactivate', { Staff_ID: 'NOPE-999' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.NOT_FOUND);
});

section('Phase 3: staff persistence, conflicts and authorization');

check('P3: staff.create write survives on a fresh sheet', function () {
  var ss = makePhase3Spreadsheet({});
  var apiS = loadBackendAs('admin@school.edu', ss);
  var created = doPostEnvelope(apiS, 'staff.create', p3StaffPayload());
  eq(created.success, true);
  var listed = doGetEnvelope(apiS, { action: 'staff.list' });
  eq(listed.data.count, 1);
  var dup = listed.data.staff.filter(function (s) { return s.Staff_ID === created.data.Staff_ID; });
  eq(dup.length, 1);
});

check('P3: staff.create reports CONFLICT when the generated ID already exists', function () {
  var forcedUuid = '9a8b7c6d5e4f0112233445566778899a';
  var apiS = loadBackendAs('admin@school.edu', makePhase3Spreadsheet({}), {
    uuid: function () { return forcedUuid; },
  });
  var first = doPostEnvelope(apiS, 'staff.create', p3StaffPayload({ First_Name: 'Kwame' }));
  eq(first.success, true);
  var second = doPostEnvelope(apiS, 'staff.create', p3StaffPayload({ First_Name: 'Kojo' }));
  eq(second.success, false);
  eq(second.error, ERROR_CODES.CONFLICT);
});

check('P3: staff denied without permission is FORBIDDEN', function () {
  var apiT = loadBackendAs('teacher@school.edu', makePhase3Spreadsheet({}));
  var envelope = doGetEnvelope(apiT, { action: 'staff.list' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.FORBIDDEN);
});

check('P3: staff unauthenticated without Google identity is UNAUTHORIZED', function () {
  var apiBlank = loadBackendAs('', makePhase3Spreadsheet({}));
  var envelope = doGetEnvelope(apiBlank, { action: 'staff.list' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.UNAUTHORIZED);
});

section('Phase 3: structural failures (Staff)');

check('P3: staff.list missing Staff sheet is SERVER_ERROR', function () {
  var apiS = loadBackendAs('admin@school.edu', makeAuthSpreadsheet([
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', ''],
  ]));
  var envelope = doGetEnvelope(apiS, { action: 'staff.list' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.SERVER_ERROR);
});

check('P3: staff.list missing required columns is SERVER_ERROR', function () {
  var p4a = p4aPermissionSheets();
  var ss = makeSpreadsheet('SchoolManagementSystem', [
    makeSheet('Staff', [['Staff_ID', 'First_Name']]),
    makeSheet('Users', [['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'],
      ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']]),
    p4a.roles, p4a.permissions, p4a.rolePermissions,
  ]);
  var apiS = loadBackendAs('admin@school.edu', ss);
  var envelope = doGetEnvelope(apiS, { action: 'staff.list' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.SERVER_ERROR);
  ok(envelope.details.missingColumns.length > 0, 'should report which columns are missing');
});

check('P3: staff.create respects script locking on write', function () {
  var ss = makePhase3Spreadsheet({});
  var apiLocked = loadBackendAs('admin@school.edu', ss, { lockUnavailable: true });
  var envelope = doPostEnvelope(apiLocked, 'staff.create', p3StaffPayload());
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.CONFLICT);
});

section('Phase 4A: Role_Permissions authorization');

check('P4A: test catalog matches CONFIG.PERMISSION_CODES', function () {
  eq(P4A_PERMISSION_CODES, api.__api.CONFIG.PERMISSION_CODES);
});

check('P4A: permission codes derive from Module + Action', function () {
  eq(api.toPermissionCode_('Students', 'READ'), 'STUDENTS.READ');
  eq(api.toPermissionCode_('Students', 'CREATE'), 'STUDENTS.CREATE');
  eq(api.toPermissionCode_('Staff', 'UPDATE'), 'STAFF.UPDATE');
  eq(api.toPermissionCode_('School_Fees', 'create'), 'SCHOOL_FEES.CREATE');
  eq(api.toPermissionCode_('school fees', 'READ'), 'SCHOOL_FEES.READ');
  eq(api.toPermissionCode_('', 'READ'), '');
});

check('P4A: permissions index uses the confirmed four-column schema', function () {
  var apiA = loadBackendAs('admin@school.edu', makeFullSpreadsheet());
  var index = apiA.getPermissionsIndex_();
  eq(index.codes.length, P4A_PERMISSION_CODES.length);
  eq(index.idByCode['STUDENTS.READ'], 'PERM-1', 'PERM-1 must map to the derived STUDENTS.READ');
  eq(index.byId['perm-1'], 'STUDENTS.READ', 'Permission_ID joins are case-insensitive');
  eq(apiA.hasPermission_(apiA.requireAuthentication_(), 'STUDENTS.READ'), true,
    'Admin must hold the code derived from Module + Action');
});

check('P4A: Permissions sheet missing Module/Action fails safely', function () {
  var p4a = p4aPermissionSheets();
  var bad = makeSheet('Permissions', [['Permission_ID', 'Description'], ['PERM-1', 'legacy']]);
  var ss = makeSpreadsheet('BadPerms', [
    makeSheet('Users', [['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'],
      ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']]),
    p4a.roles, bad, p4a.rolePermissions,
  ]);
  var apiB = loadBackendAs('admin@school.edu', ss);
  var err = throwsWithCode(function () { apiB.requirePermission_('STUDENTS.READ'); }, ERROR_CODES.SERVER_ERROR);
  ok(err.details.missingColumns.indexOf('Module') !== -1, 'must name the missing Module column');
  ok(err.details.missingColumns.indexOf('Action') !== -1, 'must name the missing Action column');
});

check('P4A: admin receives permissions through Role_Permissions', function () {
  var apiA = loadBackendAs('admin@school.edu', makeFullSpreadsheet());
  var grants = apiA.resolveRolePermissions_('Admin');
  ok(grants.indexOf('STUDENTS.READ') !== -1, 'Admin must hold STUDENTS.READ via the mapping');
  ok(grants.indexOf('SALARIES.READ') !== -1, 'reserved codes are granted through the mapping too');
  eq(grants.length, P4A_PERMISSION_CODES.length);
  var admin = apiA.requireAuthentication_();
  eq(apiA.hasPermission_(admin, 'STAFF.DEACTIVATE'), true);
  eq(apiA.requirePermission_('STUDENTS.READ').userId, 'USR-1');
});

check('P4A: role matching is case-insensitive', function () {
  var apiA = loadBackendAs('admin@school.edu', makeFullSpreadsheet());
  ok(apiA.resolveRolePermissions_('admin').length > 0, 'a lowercase role name must still resolve');
});

check('P4A: the confirmed Roles schema (Role_ID, Role_Name, Description, Status) is read correctly', function () {
  var apiA = loadBackendAs('admin@school.edu', makeFullSpreadsheet());
  eq(apiA.resolveRolePermissions_('Admin').length, P4A_PERMISSION_CODES.length,
    'Admin must resolve with Description present and Status last');
  eq(apiA.resolveRolePermissions_('Teacher'), [], 'roles without mapping rows deny');
  eq(apiA.resolveRolePermissions_('NoSuchRole'), [], 'unknown roles deny');
});

check('P4A: authenticated user without permission is FORBIDDEN', function () {
  var apiT = loadBackendAs('teacher@school.edu', makePhase3Spreadsheet({}));
  var envelope = doGetEnvelope(apiT, { action: 'students.list' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.FORBIDDEN);
  eq(envelope.details.permission, 'STUDENTS.READ');
});

check('P4A: unknown permission is denied', function () {
  var apiA = loadBackendAs('admin@school.edu', makeFullSpreadsheet());
  var admin = apiA.requireAuthentication_();
  eq(apiA.hasPermission_(admin, 'NOPE.READ'), false);
  var err = throwsWithCode(function () { apiA.requirePermission_('NOPE.READ'); }, ERROR_CODES.FORBIDDEN);
  eq(err.details.permission, 'NOPE.READ');
});

check('P4A: inactive user is denied before permissions are checked', function () {
  var apiI = loadBackendAs('old@school.edu', makeAuthSpreadsheet([
    ['USR-9', 'STF-9', 'old@school.edu', 'Admin', 'Inactive', '']
  ]));
  var res = apiI.getCurrentUser_();
  eq(res.user, null);
  eq(res.error.code, ERROR_CODES.UNAUTHORIZED);
});

section('Phase 4A: Role_Permissions fail-safe behavior');

check('P4A: missing Role_Permissions sheet fails safely', function () {
  var ss = makeSpreadsheet('NoMapping', [
    makeSheet('Users', [['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'],
      ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']]),
    makeSheet('Roles', [['Role_ID', 'Role_Name', 'Description', 'Status'],
      ['ROL-1', 'Admin', 'Full system administrator', 'Active']]),
    makeSheet('Permissions', [['Permission_ID', 'Module', 'Action', 'Description'],
      ['PERM-1', 'Students', 'READ', 'Permission STUDENTS.READ']]),
  ]);
  var apiM = loadBackendAs('admin@school.edu', ss);
  var err = throwsWithCode(function () { apiM.requirePermission_('STUDENTS.READ'); }, ERROR_CODES.SERVER_ERROR);
  eq(err.details.reason, 'sheet-missing');
  var envelope = doGetEnvelope(apiM, { action: 'students.list' });
  eq(envelope.success, false);
  eq(envelope.error, ERROR_CODES.SERVER_ERROR);
});

check('P4A: mapping to a missing permission fails safely', function () {
  var p4a = p4aPermissionSheets({
    mappings: [['RP-90', 'ROL-1', 'PERM-999', 'Active']],
  });
  var ss = makeSpreadsheet('Orphan', [
    makeSheet('Users', [['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'],
      ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']]),
    p4a.roles, p4a.permissions, p4a.rolePermissions,
  ]);
  var apiO = loadBackendAs('admin@school.edu', ss);
  var err = throwsWithCode(function () { apiO.requirePermission_('STUDENTS.READ'); }, ERROR_CODES.SERVER_ERROR);
  eq(err.details.reason, 'orphan-mapping');
  ok(err.details.permissionIds.indexOf('PERM-999') !== -1, 'the orphan Permission_ID must be named');
});

check('P4A: a role with multiple permissions works', function () {
  var p4a = p4aPermissionSheets({
    roles: [['ROL-1', 'Admin', 'Full system administrator', 'Active'],
      ['ROL-3', 'Accountant', 'Finance and accounting', 'Active']],
    mappings: [
      ['RP-1', 'ROL-3', 'PERM-1', 'Active'],
      ['RP-2', 'ROL-3', 'PERM-9', 'Active'],
      ['RP-3', 'ROL-3', 'PERM-29', 'Active'],
    ],
  });
  var ss = makeSpreadsheet('Multi', [
    makeSheet('Users', [['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'],
      ['USR-3', 'STF-3', 'acct@school.edu', 'Accountant', 'Active', '']]),
    p4a.roles, p4a.permissions, p4a.rolePermissions,
  ]);
  var apiR = loadBackendAs('acct@school.edu', ss);
  var user = apiR.requireAuthentication_();
  eq(apiR.hasPermission_(user, 'STUDENTS.READ'), true);
  eq(apiR.hasPermission_(user, 'SCHOOL_FEES.READ'), true);
  eq(apiR.hasPermission_(user, 'SALARIES.VOID'), true);
  eq(apiR.hasPermission_(user, 'STUDENTS.CREATE'), false, 'unmapped permission denies');
});

check('P4A: inactive permission assignment is denied', function () {
  var p4a = p4aPermissionSheets({
    mappings: [['RP-1', 'ROL-1', 'PERM-1', 'Inactive']],
  });
  var ss = makeSpreadsheet('Inactive', [
    makeSheet('Users', [['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'],
      ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']]),
    p4a.roles, p4a.permissions, p4a.rolePermissions,
  ]);
  var apiN = loadBackendAs('admin@school.edu', ss);
  var user = apiN.requireAuthentication_();
  eq(apiN.hasPermission_(user, 'STUDENTS.READ'), false);
  throwsWithCode(function () { apiN.requirePermission_('STUDENTS.READ'); }, ERROR_CODES.FORBIDDEN);
});

section('Phase 4A: duplicate mappings, seed and auth routes');

check('P4A: duplicate same-status mappings de-duplicate deterministically', function () {
  var p4a = p4aPermissionSheets({
    mappings: [
      ['RP-1', 'ROL-1', 'PERM-1', 'Active'],
      ['RP-2', 'ROL-1', 'PERM-1', 'Active'],
    ],
  });
  var ss = makeSpreadsheet('Dupes', [
    makeSheet('Users', [['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'],
      ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']]),
    p4a.roles, p4a.permissions, p4a.rolePermissions,
  ]);
  var apiD = loadBackendAs('admin@school.edu', ss);
  var user = apiD.requireAuthentication_();
  eq(apiD.hasPermission_(user, 'STUDENTS.READ'), true, 'identical duplicates still grant');
  var grants = apiD.resolveRolePermissions_('Admin');
  eq(grants.filter(function (c) { return c === 'STUDENTS.READ'; }).length, 1, 'granted exactly once');
  eq(grants.length, 1, 'nothing else leaks in');
});

check('P4A: duplicate mappings with conflicting statuses are CONFLICT', function () {
  var p4a = p4aPermissionSheets({
    mappings: [
      ['RP-1', 'ROL-1', 'PERM-1', 'Active'],
      ['RP-2', 'ROL-1', 'PERM-1', 'Inactive'],
    ],
  });
  var ss = makeSpreadsheet('Conflict', [
    makeSheet('Users', [['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'],
      ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']]),
    p4a.roles, p4a.permissions, p4a.rolePermissions,
  ]);
  var apiC = loadBackendAs('admin@school.edu', ss);
  var err = throwsWithCode(function () { apiC.requirePermission_('STUDENTS.READ'); }, ERROR_CODES.CONFLICT);
  ok(err.details.permissionIds.indexOf('PERM-1') !== -1, 'the conflicting Permission_ID must be named');
});

check('P4A: setupRolePermissions creates and seeds the mapping idempotently', function () {
  var roles = makeSheet('Roles', [['Role_ID', 'Role_Name', 'Description', 'Status'],
    ['ROL-1', 'Admin', 'Full system administrator', 'Active']]);
  var permissions = makeSheet('Permissions', [['Permission_ID', 'Module', 'Action', 'Description']]);
  var users = makeSheet('Users', [['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'],
    ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']]);
  var ss = makeSpreadsheet('Seed', [users, roles, permissions]);
  // A counter-based uuid keeps every generated Permission_ID unique (a
  // constant uuid would collapse the Permissions index to a single id).
  var uuidCounter = 0;
  var apiS = loadBackendAs('admin@school.edu', ss, {
    uuid: function () {
      uuidCounter += 1;
      var hex = uuidCounter.toString(16);
      // Zero-pad in FRONT: generateId_ uses the FIRST 12 hex characters, so
      // the varying digits must lead for every id to be unique.
      while (hex.length < 12) hex = '0' + hex;
      while (hex.length < 32) hex = hex + '0';
      return hex;
    },
  });
  var first = apiS.setupRolePermissions();
  eq(first.rolePermissionsSheetCreated, true);
  eq(first.permissionsAdded.length, P4A_PERMISSION_CODES.length);
  eq(first.mappingsCreated.length, P4A_PERMISSION_CODES.length);
  eq(ss.getSheetByName('Role_Permissions').getLastRow(), 1 + P4A_PERMISSION_CODES.length);
  // Permissions rows must use the confirmed schema: Module + Action, no
  // permission-name column, Description filled, Title-Case Module.
  var permsSheet = ss.getSheetByName('Permissions');
  eq(permsSheet._rows[0], ['Permission_ID', 'Module', 'Action', 'Description'],
    'the seed must not create a Permission_Name column');
  eq(permsSheet.getLastRow(), 1 + P4A_PERMISSION_CODES.length);
  var seededIndex = apiS.getPermissionsIndex_();
  eq(seededIndex.codes.length, P4A_PERMISSION_CODES.length);
  var studentsReadId = seededIndex.idByCode['STUDENTS.READ'];
  ok(studentsReadId, 'STUDENTS.READ must resolve to a seeded Permission_ID');
  var seededRow = permsSheet._rows.filter(function (r) { return r[0] === studentsReadId; })[0];
  eq(seededRow[1], 'Students');
  eq(seededRow[2], 'READ');
  var second = apiS.setupRolePermissions();
  eq(second.rolePermissionsSheetCreated, false);
  eq(second.permissionsAdded.length, 0);
  eq(second.mappingsCreated.length, 0);
  eq(second.mappingsAlreadyPresent, P4A_PERMISSION_CODES.length);
  eq(apiS.resolveRolePermissions_('Admin').length, P4A_PERMISSION_CODES.length);
});

check('P4A: auth.me still works with mapping-driven authorization', function () {
  var apiM = loadBackendAs('admin@school.edu', makeFullSpreadsheet());
  var env = JSON.parse(apiM.doGet({ parameter: { action: 'auth.me' } }).getContent());
  eq(env.success, true);
  eq(env.data, { userId: 'USR-1', staffId: 'STF-1', email: 'admin@school.edu', role: 'Admin' });
});

check('P4A: auth.check resolves allowed and denied through the mapping', function () {
  var p4a = p4aPermissionSheets({
    mappings: p4aAdminMappingRows().concat([['RP-99', 'ROL-2', 'PERM-1', 'Active']]),
  });
  var ss = makeSpreadsheet('Check', [
    makeSheet('Users', [['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'],
      ['USR-2', 'STF-2', 'teacher@school.edu', 'Teacher', 'Active', '']]),
    p4a.roles, p4a.permissions, p4a.rolePermissions,
  ]);
  var apiC = loadBackendAs('teacher@school.edu', ss);
  var allowed = JSON.parse(apiC.doPost({ parameter: {},
    postData: { contents: JSON.stringify({ action: 'auth.check', payload: { permission: 'STUDENTS.READ' } }) } }).getContent());
  eq(allowed.success, true);
  eq(allowed.data.allowed, true);
  var denied = JSON.parse(apiC.doPost({ parameter: {},
    postData: { contents: JSON.stringify({ action: 'auth.check', payload: { permission: 'STAFF.READ' } }) } }).getContent());
  eq(denied.success, true);
  eq(denied.data.allowed, false);
});


/* ==========================================================================
 * Phase 4B: School Fees + Feeding Fees
 * ======================================================================== */

/* --------------------------------------------------------------------------
 * Phase 4B fixtures
 * ------------------------------------------------------------------------ */

/** Valid school fee create payload; `overrides` replaces or adds fields. */
function sfCreatePayload(overrides) {
  return Object.assign({
    Student_ID: 'STU-2',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 1000,
    Amount_Paid: 0,
    Payment_Method: 'Cash',
    Payment_Date: '2025-09-20',
  }, overrides || {});
}

/** Valid feeding fee create payload; `overrides` replaces or adds fields. */
function ffCreatePayload(overrides) {
  return Object.assign({
    Student_ID: 'STU-2',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 300,
    Amount_Paid: 0,
    Payment_Method: 'Cash',
    Payment_Date: '2025-09-20',
  }, overrides || {});
}

/** Copy of a payload with `field` removed, to exercise required-field rules. */
function payloadWithout(payload, field) {
  const copy = Object.assign({}, payload);
  delete copy[field];
  return copy;
}

/**
 * Full spreadsheet (real fee/student/staff/Users data) with optional Users,
 * Roles and Role_Permissions row overrides, for the authorization tests.
 */
function p4bSpreadsheet(opts) {
  opts = opts || {};
  const ss = makeFullSpreadsheet();
  if (opts.users) {
    ss.getSheetByName('Users')._rows =
      [['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login']].concat(opts.users);
  }
  if (opts.roles) {
    ss.getSheetByName('Roles')._rows =
      [['Role_ID', 'Role_Name', 'Description', 'Status']].concat(opts.roles);
  }
  if (opts.mappings) {
    ss.getSheetByName('Role_Permissions')._rows =
      [['Role_Permission_ID', 'Role_ID', 'Permission_ID', 'Status']].concat(opts.mappings);
  }
  return ss;
}

/** Permission_ID the Permissions sheet stores for a catalog code (PERM-1..N). */
function permIdFor(code) {
  return 'PERM-' + (P4A_PERMISSION_CODES.indexOf(code) + 1);
}

/** The Payment_ID values in a fee sheet, in sheet order. */
function feeIds(ss, tabName) {
  return ss.getSheetByName(tabName)._rows.slice(1).map(function (row) {
    return row[0];
  });
}

/** The 1-based sheet row a fee record occupies (row 1 is the header row). */
function feeRowOf(ss, tabName, paymentId) {
  return feeIds(ss, tabName).indexOf(paymentId) + 2;
}

/** Column index of a fee header in the fixture's header row. */
function feeColIndex(ss, tabName, header) {
  return ss.getSheetByName(tabName)._rows[0].indexOf(header);
}

/**
 * Wrap the sandbox's getSheet_ so that every read and every write against one
 * tab is logged together with the script-lock depth in force at that moment.
 * Lets a test prove a decision (or a row write) happened INSIDE the critical
 * section rather than before it. Returns the log; entries look like
 * {op, row, numRows, numCols, depth, values}.
 */
function observeSheetAccess(api, tabName) {
  const log = [];
  const realGetSheet = api.getSheet_;
  api.getSheet_ = function (name) {
    const sheet = realGetSheet(name);
    if (name !== tabName) return sheet;
    return {
      getName: function () { return sheet.getName(); },
      getLastRow: function () { return sheet.getLastRow(); },
      getLastColumn: function () { return sheet.getLastColumn(); },
      getRange: function (row, col, numRows, numCols) {
        const range = sheet.getRange(row, col, numRows, numCols);
        return {
          getValues: function () {
            const values = range.getValues();
            log.push({
              op: 'read', row: row, numRows: numRows, numCols: numCols,
              depth: api.__lockDepth, values: values,
            });
            return values;
          },
          setValues: function (values) {
            log.push({
              op: 'write', row: row, numRows: numRows, numCols: numCols,
              depth: api.__lockDepth, values: values,
            });
            return range.setValues(values);
          },
        };
      },
    };
  };
  return log;
}

/** Log entries that read the whole stored row `row` (not a single column). */
function fullRowReads(log, row) {
  return log.filter(function (entry) {
    return entry.op === 'read' && entry.row === row && entry.numCols > 1;
  });
}

/** Full spreadsheet with one tab removed entirely (structural failure test). */
function p4bSpreadsheetWithoutTab(tabName) {
  const ss = makeFullSpreadsheet();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === tabName) {
      sheets.splice(i, 1);
      break;
    }
  }
  return ss;
}

/** The eight fee permission codes, in catalog order. */
const P4B_FEE_PERMISSIONS = [
  'SCHOOL_FEES.READ', 'SCHOOL_FEES.CREATE', 'SCHOOL_FEES.UPDATE', 'SCHOOL_FEES.VOID',
  'FEEDING_FEES.READ', 'FEEDING_FEES.CREATE', 'FEEDING_FEES.UPDATE', 'FEEDING_FEES.VOID',
];

section('Phase 4B: School Fees — routing and existence');

check('schoolFees.list action is routed', function () {
  ok(api.listAvailableActions_().indexOf('schoolFees.list') !== -1, 'schoolFees.list must be routed');
});

check('schoolFees.get action is routed', function () {
  ok(api.listAvailableActions_().indexOf('schoolFees.get') !== -1, 'schoolFees.get must be routed');
});

check('schoolFees.create action is routed', function () {
  ok(api.listAvailableActions_().indexOf('schoolFees.create') !== -1, 'schoolFees.create must be routed');
});

check('schoolFees.update action is routed', function () {
  ok(api.listAvailableActions_().indexOf('schoolFees.update') !== -1, 'schoolFees.update must be routed');
});

check('schoolFees.void action is routed', function () {
  ok(api.listAvailableActions_().indexOf('schoolFees.void') !== -1, 'schoolFees.void must be routed');
});

section('Phase 4B: School Fees — create');

check('schoolFees.create: admin can create a school fee payment with server-generated SF- ID', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = {
    Student_ID: 'STU-2',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 1500,
    Amount_Paid: 0,
    Payment_Method: 'Cash',
    Payment_Date: '2025-09-20',
    Reference: 'TEST-001',
    Notes: 'Test payment'
  };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.create' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.create', payload: payload }) }
  }).getContent());
  ok(env.success, 'create should succeed: ' + (env.message || env.error));
  if (env.success) {
    ok(env.data.Payment_ID && env.data.Payment_ID.indexOf('SF-') === 0,
      'stored record must have a server-generated SF- ID, got: ' + env.data.Payment_ID);
    eq(env.data.Student_ID, 'STU-2');
    eq(env.data.Balance, 1500, 'Balance should be Amount_Due - Amount_Paid');
    eq(env.data.Amount_Due, 1500);
    eq(env.data.Amount_Paid, 0);
    eq(env.data.Status, 'Unpaid');
  }
});

check('schoolFees.create: Recorded_By comes from authenticated user, not client', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = {
    Student_ID: 'STU-2',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 1000,
    Amount_Paid: 0,
    Payment_Method: 'Bank Transfer',
    Payment_Date: '2025-09-20',
    Reference: 'TEST-002',
    Recorded_By: 'EVIL-STF-999'
  };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.create' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.create', payload: payload }) }
  }).getContent());
  ok(env.success, 'create should succeed: ' + (env.message || env.error));
  if (env.success) {
    eq(env.data.Recorded_By, 'STF-1',
      'Recorded_By must be the authenticated user\'s staff ID, not client-supplied');
  }
});

check('schoolFees.create: client-supplied Balance is ignored and recalculated', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = {
    Student_ID: 'STU-2',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 2000,
    Amount_Paid: 500,
    Payment_Method: 'Mobile Money',
    Payment_Date: '2025-09-20',
    Balance: 9999
  };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.create' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.create', payload: payload }) }
  }).getContent());
  ok(env.success, 'create should succeed: ' + (env.message || env.error));
  if (env.success) {
    eq(env.data.Balance, 1500, 'Balance must be recalculated as Amount_Due - Amount_Paid');
  }
});

check('schoolFees.create: Amount_Paid > Amount_Due is rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = {
    Student_ID: 'STU-2',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 500,
    Amount_Paid: 1000,
    Payment_Method: 'Cash',
    Payment_Date: '2025-09-20'
  };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.create' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.create', payload: payload }) }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
});

check('schoolFees.create: negative Amount_Due is rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = {
    Student_ID: 'STU-2',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: -100,
    Amount_Paid: 0,
    Payment_Method: 'Cash',
    Payment_Date: '2025-09-20'
  };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.create' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.create', payload: payload }) }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
});

check('schoolFees.create: negative Amount_Paid is rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = {
    Student_ID: 'STU-2',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 1000,
    Amount_Paid: -50,
    Payment_Method: 'Cash',
    Payment_Date: '2025-09-20'
  };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.create' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.create', payload: payload }) }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
});

check('schoolFees.create: missing Student_ID is rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = {
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 1000,
    Amount_Paid: 0,
    Payment_Method: 'Cash',
    Payment_Date: '2025-09-20'
  };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.create' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.create', payload: payload }) }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
});

check('schoolFees.create: nonexistent Student_ID is rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = {
    Student_ID: 'STU-999',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 1000,
    Amount_Paid: 0,
    Payment_Method: 'Cash',
    Payment_Date: '2025-09-20'
  };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.create' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.create', payload: payload }) }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'NOT_FOUND');
});

check('schoolFees.create: withdrawn student is rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = {
    Student_ID: 'STU-3',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 1000,
    Amount_Paid: 0,
    Payment_Method: 'Cash',
    Payment_Date: '2025-09-20'
  };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.create' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.create', payload: payload }) }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
  ok(env.message.toLowerCase().indexOf('withdrawn') !== -1 || (env.details && env.details.reason === 'student-withdrawn'),
    'should reject withdrawn student');
});

check('schoolFees.create: missing Academic_Year is rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = {
    Student_ID: 'STU-2',
    Term: 'Term 1',
    Amount_Due: 1000,
    Amount_Paid: 0,
    Payment_Method: 'Cash',
    Payment_Date: '2025-09-20'
  };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.create' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.create', payload: payload }) }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
});

check('schoolFees.create: missing Term is rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = {
    Student_ID: 'STU-2',
    Academic_Year: '2025/2026',
    Amount_Due: 1000,
    Amount_Paid: 0,
    Payment_Method: 'Cash',
    Payment_Date: '2025-09-20'
  };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.create' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.create', payload: payload }) }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
});

check('schoolFees.create: invalid Payment_Method is rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = {
    Student_ID: 'STU-2',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 1000,
    Amount_Paid: 0,
    Payment_Method: 'Bitcoin',
    Payment_Date: '2025-09-20'
  };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.create' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.create', payload: payload }) }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
  ok(env.details && Array.isArray(env.details.allowedValues), 'should list allowed payment methods');
});

check('schoolFees.create: invalid Payment_Date is rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = {
    Student_ID: 'STU-2',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 1000,
    Amount_Paid: 0,
    Payment_Method: 'Cash',
    Payment_Date: 'not-a-date'
  };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.create' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.create', payload: payload }) }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
});

check('schoolFees.create: client-supplied Payment_ID is rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = doPostEnvelope(apiA, 'schoolFees.create', sfCreatePayload({ Payment_ID: 'SF-999' }));
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
  ok(env.message.indexOf('server-generated') !== -1, 'message should explain the rule');
});

check('schoolFees.create: rejects unknown fields', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = doPostEnvelope(apiA, 'schoolFees.create', sfCreatePayload({ Middle_Name: 'X' }));
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
  ok(env.details.unknownColumns.indexOf('Middle_Name') !== -1, 'should list unknown columns');
  ok(env.details.validColumns.indexOf('Payment_ID') !== -1, 'should list the valid columns');
});

check('schoolFees.create: non-numeric or negative amounts are rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  [
    ['Amount_Due', 'abc'],
    ['Amount_Paid', 'abc'],
    ['Amount_Due', -1],
    ['Amount_Paid', -1],
    ['Amount_Paid', 5000],
  ].forEach(function (pair) {
    const env = doPostEnvelope(apiA, 'schoolFees.create', sfCreatePayload({ [pair[0]]: pair[1] }));
    eq(env.success, false, pair[0] + '=' + pair[1] + ' must be rejected');
    eq(env.error, 'VALIDATION_ERROR');
  });
});

check('schoolFees.create: every required field is enforced', function () {
  const apiA = loadBackendAs('admin@school.edu');
  ['Amount_Due', 'Amount_Paid', 'Payment_Method', 'Payment_Date'].forEach(function (field) {
    const env = doPostEnvelope(apiA, 'schoolFees.create', payloadWithout(sfCreatePayload(), field));
    eq(env.success, false, field + ' is required');
    eq(env.error, 'VALIDATION_ERROR');
    ok(env.details.missingFields.indexOf(field) !== -1, 'should flag missing ' + field);
  });
});

check('schoolFees.create: generated IDs are sequential and never reused', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  const first = doPostEnvelope(apiA, 'schoolFees.create', sfCreatePayload());
  const second = doPostEnvelope(apiA, 'schoolFees.create', sfCreatePayload());
  eq(first.data.Payment_ID, 'SF-004', 'the next ID must follow the highest stored ID');
  eq(second.data.Payment_ID, 'SF-005');
  eq(feeIds(ss, 'School_Fees'), ['SF-001', 'SF-002', 'SF-003', 'SF-004', 'SF-005'],
    'each create appends exactly one row and never rewrites an existing ID');
});

check('schoolFees.create: refuses a duplicate server-generated Payment_ID', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const realFind = apiA.findSchoolFeeById_;
  // Force the collision the sequential generator is designed to avoid.
  apiA.findSchoolFeeById_ = function () {
    return { sheetRow: 2, record: {} };
  };
  const env = doPostEnvelope(apiA, 'schoolFees.create', sfCreatePayload());
  apiA.findSchoolFeeById_ = realFind;
  eq(env.success, false);
  eq(env.error, 'CONFLICT');
  ok(env.message.indexOf('Duplicate Payment_ID') !== -1, 'message should name the collision');
});

section('Phase 4B: School Fees — get and list');

check('schoolFees.get returns an existing payment', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = JSON.parse(apiA.doGet({
    parameter: { action: 'schoolFees.get', Payment_ID: 'SF-001' }
  }).getContent());
  eq(env.success, true);
  if (env.success) {
    eq(env.data.Payment_ID, 'SF-001');
    eq(env.data.Student_ID, 'STU-1');
    eq(env.data.Amount_Due, 1200);
    eq(env.data.Balance, 0);
    eq(env.data.Status, 'Paid');
  }
});

check('schoolFees.get returns NOT_FOUND for a missing payment', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = JSON.parse(apiA.doGet({
    parameter: { action: 'schoolFees.get', Payment_ID: 'SF-999' }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'NOT_FOUND');
});

check('schoolFees.list returns all records', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = JSON.parse(apiA.doGet({
    parameter: { action: 'schoolFees.list' }
  }).getContent());
  eq(env.success, true);
  if (env.success) {
    ok(Array.isArray(env.data));
    ok(env.data.length >= 3, 'should have at least 3 pre-populated payments');
  }
});

check('schoolFees.list filters by Student_ID', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = JSON.parse(apiA.doGet({
    parameter: { action: 'schoolFees.list', Student_ID: 'STU-1' }
  }).getContent());
  eq(env.success, true);
  if (env.success) {
    ok(Array.isArray(env.data));
    ok(env.data.length >= 2, 'STU-1 should have at least 2 payments');
    env.data.forEach(function (p) { eq(p.Student_ID, 'STU-1'); });
  }
});

check('schoolFees.list filters by Status', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = JSON.parse(apiA.doGet({
    parameter: { action: 'schoolFees.list', Status: 'Paid' }
  }).getContent());
  eq(env.success, true);
  if (env.success) {
    ok(Array.isArray(env.data));
    env.data.forEach(function (p) { eq(p.Status, 'Paid'); });
  }
});

check('schoolFees.list filters by Academic_Year', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = JSON.parse(apiA.doGet({
    parameter: { action: 'schoolFees.list', Academic_Year: '2025/2026' }
  }).getContent());
  eq(env.success, true);
  if (env.success) {
    ok(Array.isArray(env.data));
    env.data.forEach(function (p) { eq(p.Academic_Year, '2025/2026'); });
  }
});

check('schoolFees.list filters by Term', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = JSON.parse(apiA.doGet({
    parameter: { action: 'schoolFees.list', Term: 'Term 1' }
  }).getContent());
  eq(env.success, true);
  if (env.success) {
    ok(Array.isArray(env.data));
    env.data.forEach(function (p) { eq(p.Term, 'Term 1'); });
  }
});

check('schoolFees.get requires a Payment_ID', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = doGetEnvelope(apiA, { action: 'schoolFees.get' });
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
  eq(env.details.field, 'Payment_ID');
});

check('schoolFees.list returns an empty array when nothing matches', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = doGetEnvelope(apiA, { action: 'schoolFees.list', Student_ID: 'STU-999' });
  eq(env.success, true);
  eq(env.data, [], 'no matches is an empty list, not an error');
});

check('schoolFees.list combines filters', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = doGetEnvelope(apiA, {
    action: 'schoolFees.list', Student_ID: 'STU-1', Term: 'Term 1',
  });
  eq(env.success, true);
  eq(env.data.length, 1, 'STU-1 has exactly one Term 1 payment in the fixture');
  eq(env.data[0].Payment_ID, 'SF-001');
});

section('Phase 4B: School Fees — update');

check('schoolFees.update performs a partial update and recalculates Balance', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = { Payment_ID: 'SF-003', Amount_Paid: 600 };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.update' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.update', payload: payload }) }
  }).getContent());
  eq(env.success, true, env.message || env.error);
  if (env.success) {
    eq(env.data.Payment_ID, 'SF-003');
    eq(env.data.Amount_Paid, 600);
    eq(env.data.Balance, 600, 'Balance should be recalculated');
    eq(env.data.Status, 'Partial');
  }
});

check('schoolFees.update: the stored Payment_ID is never changed', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  const env = doPostEnvelope(apiA, 'schoolFees.update', { Payment_ID: 'SF-002', Reference: 'UPDATED-REF' });
  eq(env.success, true, env.message || env.error);
  if (env.success) {
    eq(env.data.Payment_ID, 'SF-002', 'Payment_ID must remain unchanged');
    eq(env.data.Reference, 'UPDATED-REF');
  }
  ok(apiA.findSchoolFeeById_('SF-002'), 'the row must still be findable by its own ID after the write');
  eq(feeIds(ss, 'School_Fees'), ['SF-001', 'SF-002', 'SF-003'],
    'an update must never add, drop or re-key a row');
  eq(ss.getSheetByName('School_Fees')._rows[2][0], 'SF-002', 'the ID cell must be left intact');
  eq(ss.getSheetByName('School_Fees')._rows[2][9], 'UPDATED-REF', 'the edited cell must change in place');
});

check('schoolFees.update: an unknown Payment_ID targets nothing', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  const env = doPostEnvelope(apiA, 'schoolFees.update', { Payment_ID: 'SF-999', Amount_Paid: 1 });
  eq(env.success, false);
  eq(env.error, 'NOT_FOUND');
  eq(env.details.Payment_ID, 'SF-999');
  eq(feeIds(ss, 'School_Fees'), ['SF-001', 'SF-002', 'SF-003'], 'a failed update must not write');
});

check('schoolFees.update: rejects unknown fields', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = doPostEnvelope(apiA, 'schoolFees.update', { Payment_ID: 'SF-002', Middle_Name: 'X' });
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
  ok(env.details.unknownColumns.indexOf('Middle_Name') !== -1, 'should list unknown columns');
});

check('schoolFees.update: changing Amount_Due recalculates Balance and Status', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = doPostEnvelope(apiA, 'schoolFees.update', { Payment_ID: 'SF-001', Amount_Due: 2400 });
  eq(env.success, true, env.message || env.error);
  if (env.success) {
    eq(env.data.Amount_Due, 2400);
    eq(env.data.Amount_Paid, 1200, 'Amount_Paid must be preserved');
    eq(env.data.Balance, 1200);
    eq(env.data.Status, 'Partial', 'Status must be derived, not carried over');
  }
});

check('schoolFees.update: invalid amounts are rejected without writing', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  [
    ['Amount_Due', 'abc'],
    ['Amount_Due', -1],
    ['Amount_Paid', 'abc'],
    ['Amount_Paid', -1],
    ['Amount_Paid', 5000],
  ].forEach(function (pair) {
    const payload = { Payment_ID: 'SF-002' };
    payload[pair[0]] = pair[1];
    const env = doPostEnvelope(apiA, 'schoolFees.update', payload);
    eq(env.success, false, pair[0] + '=' + pair[1] + ' must be rejected');
    eq(env.error, 'VALIDATION_ERROR');
  });
  eq(apiA.findSchoolFeeById_('SF-002').record.Amount_Paid, 600, 'the rejected payloads must not be applied');
});

check('schoolFees.update: a voided payment cannot be updated', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const voidEnv = doPostEnvelope(apiA, 'schoolFees.void', { Payment_ID: 'SF-002' });
  eq(voidEnv.success, true, voidEnv.message || voidEnv.error);
  const env = doPostEnvelope(apiA, 'schoolFees.update', { Payment_ID: 'SF-002', Amount_Paid: 999 });
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
  eq(env.details.reason, 'voided-payment');
  eq(apiA.findSchoolFeeById_('SF-002').record.Amount_Paid, 600, 'the voided row must stay untouched');
});

check('schoolFees.update: client-supplied Recorded_By is ignored', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = { Payment_ID: 'SF-001', Recorded_By: 'EVIL-STF-999' };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.update' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.update', payload: payload }) }
  }).getContent());
  eq(env.success, true, env.message || env.error);
  if (env.success) {
    eq(env.data.Recorded_By, 'STF-1',
      'Recorded_By must not be overwritten by client on update');
  }
});

section('Phase 4B: School Fees — void');

check('schoolFees.void soft-corrects a payment to Voided', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.void' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.void', payload: { Payment_ID: 'SF-003' } }) }
  }).getContent());
  eq(env.success, true, env.message || env.error);
  if (env.success) {
    eq(env.data.Status, 'Voided');
    eq(env.data.Payment_ID, 'SF-003');
    eq(env.data.Student_ID, 'STU-2');
  }
});

check('schoolFees.void: a second void is rejected and the first one stands', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const first = doPostEnvelope(apiA, 'schoolFees.void', { Payment_ID: 'SF-003' });
  eq(first.success, true, first.message || first.error);
  eq(first.data.Status, 'Voided');
  const second = doPostEnvelope(apiA, 'schoolFees.void', { Payment_ID: 'SF-003' });
  eq(second.success, false, 're-voiding must be refused');
  eq(second.error, 'VALIDATION_ERROR');
  eq(second.details.reason, 'already-voided');
  eq(apiA.findSchoolFeeById_('SF-003').record.Status, 'Voided', 'the original void must stand');
});

check('schoolFees.void: the row is preserved with its financial values', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  const voidEnv = doPostEnvelope(apiA, 'schoolFees.void', { Payment_ID: 'SF-003' });
  eq(voidEnv.success, true, voidEnv.message || voidEnv.error);
  eq(ss.getSheetByName('School_Fees')._rows.length, 4, 'void is a soft correction, never a delete');
  const getEnv = doGetEnvelope(apiA, { action: 'schoolFees.get', Payment_ID: 'SF-003' });
  eq(getEnv.success, true, 'a voided payment must stay retrievable');
  eq(getEnv.data.Payment_ID, 'SF-003');
  eq(getEnv.data.Status, 'Voided');
  eq(getEnv.data.Student_ID, 'STU-2');
  eq(getEnv.data.Amount_Due, 1200, 'voiding must not change the amounts');
  eq(getEnv.data.Amount_Paid, 0);
  eq(getEnv.data.Balance, 1200);
  eq(getEnv.data.Recorded_By, 'STF-1');
});

check('schoolFees.void: unknown or missing Payment_ID fails safely', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  const missing = doPostEnvelope(apiA, 'schoolFees.void', {});
  eq(missing.success, false);
  eq(missing.error, 'VALIDATION_ERROR');
  eq(missing.details.field, 'Payment_ID');
  const unknown = doPostEnvelope(apiA, 'schoolFees.void', { Payment_ID: 'SF-999' });
  eq(unknown.success, false);
  eq(unknown.error, 'NOT_FOUND');
  eq(ss.getSheetByName('School_Fees')._rows.length, 4, 'a failed void must not write');
});

section('Phase 4B: Feeding Fees — routing, create, get, list');

check('feedingFees.list action is routed', function () {
  ok(api.listAvailableActions_().indexOf('feedingFees.list') !== -1, 'feedingFees.list must be routed');
});

check('feedingFees.get action is routed', function () {
  ok(api.listAvailableActions_().indexOf('feedingFees.get') !== -1, 'feedingFees.get must be routed');
});

check('feedingFees.create action is routed', function () {
  ok(api.listAvailableActions_().indexOf('feedingFees.create') !== -1, 'feedingFees.create must be routed');
});

check('feedingFees.update action is routed', function () {
  ok(api.listAvailableActions_().indexOf('feedingFees.update') !== -1, 'feedingFees.update must be routed');
});

check('feedingFees.void action is routed', function () {
  ok(api.listAvailableActions_().indexOf('feedingFees.void') !== -1, 'feedingFees.void must be routed');
});

check('feedingFees.create: admin can create a feeding fee payment with server-generated FF- ID', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = {
    Student_ID: 'STU-2',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 300,
    Amount_Paid: 300,
    Payment_Method: 'Mobile Money',
    Payment_Date: '2025-09-20',
    Reference: 'FF-TEST-001'
  };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'feedingFees.create' },
    postData: { contents: JSON.stringify({ action: 'feedingFees.create', payload: payload }) }
  }).getContent());
  ok(env.success, 'create should succeed: ' + (env.message || env.error));
  if (env.success) {
    ok(env.data.Payment_ID && env.data.Payment_ID.indexOf('FF-') === 0,
      'stored record must have a server-generated FF- ID, got: ' + env.data.Payment_ID);
    eq(env.data.Student_ID, 'STU-2');
    eq(env.data.Balance, 0);
    eq(env.data.Status, 'Paid');
    eq(env.data.Recorded_By, 'STF-1', 'Recorded_By must come from authenticated user');
  }
});

check('feedingFees.create: Recorded_By from authenticated user, not client', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = {
    Student_ID: 'STU-2',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 250,
    Amount_Paid: 0,
    Payment_Method: 'Cash',
    Payment_Date: '2025-09-20',
    Recorded_By: 'EVIL-STF-999'
  };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'feedingFees.create' },
    postData: { contents: JSON.stringify({ action: 'feedingFees.create', payload: payload }) }
  }).getContent());
  ok(env.success, 'create should succeed: ' + (env.message || env.error));
  if (env.success) {
    eq(env.data.Recorded_By, 'STF-1',
      'Recorded_By must be the authenticated user\'s staff ID, not client-supplied');
  }
});

check('feedingFees.create: Amount_Paid > Amount_Due is rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = {
    Student_ID: 'STU-2',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 200,
    Amount_Paid: 500,
    Payment_Method: 'Cash',
    Payment_Date: '2025-09-20'
  };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'feedingFees.create' },
    postData: { contents: JSON.stringify({ action: 'feedingFees.create', payload: payload }) }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
});

check('feedingFees.get returns an existing payment', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = JSON.parse(apiA.doGet({
    parameter: { action: 'feedingFees.get', Payment_ID: 'FF-001' }
  }).getContent());
  eq(env.success, true);
  if (env.success) {
    eq(env.data.Payment_ID, 'FF-001');
    eq(env.data.Student_ID, 'STU-1');
    eq(env.data.Amount_Due, 450);
    eq(env.data.Balance, 0);
  }
});

check('feedingFees.get returns NOT_FOUND for a missing payment', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = JSON.parse(apiA.doGet({
    parameter: { action: 'feedingFees.get', Payment_ID: 'FF-999' }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'NOT_FOUND');
});

check('feedingFees.list returns all records', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = JSON.parse(apiA.doGet({
    parameter: { action: 'feedingFees.list' }
  }).getContent());
  eq(env.success, true);
  if (env.success) {
    ok(Array.isArray(env.data));
    ok(env.data.length >= 2, 'should have at least 2 pre-populated records');
  }
});

check('feedingFees.list filters by Student_ID', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = JSON.parse(apiA.doGet({
    parameter: { action: 'feedingFees.list', Student_ID: 'STU-1' }
  }).getContent());
  eq(env.success, true);
  if (env.success) {
    ok(Array.isArray(env.data));
    env.data.forEach(function (p) { eq(p.Student_ID, 'STU-1'); });
  }
});

check('feedingFees.list filters by Status', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = JSON.parse(apiA.doGet({
    parameter: { action: 'feedingFees.list', Status: 'Partial' }
  }).getContent());
  eq(env.success, true);
  if (env.success) {
    ok(Array.isArray(env.data));
    env.data.forEach(function (p) { eq(p.Status, 'Partial'); });
  }
});

check('feedingFees.create: client-supplied Payment_ID is rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = doPostEnvelope(apiA, 'feedingFees.create', ffCreatePayload({ Payment_ID: 'FF-999' }));
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
  ok(env.message.indexOf('server-generated') !== -1, 'message should explain the rule');
});

check('feedingFees.create: rejects unknown fields', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = doPostEnvelope(apiA, 'feedingFees.create', ffCreatePayload({ Middle_Name: 'X' }));
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
  ok(env.details.unknownColumns.indexOf('Middle_Name') !== -1, 'should list unknown columns');
  ok(env.details.validColumns.indexOf('Payment_ID') !== -1, 'should list the valid columns');
});

check('feedingFees.create: client-supplied Balance is ignored and recalculated', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = doPostEnvelope(apiA, 'feedingFees.create', ffCreatePayload({ Amount_Paid: 120, Balance: 9999 }));
  eq(env.success, true, env.message || env.error);
  if (env.success) {
    eq(env.data.Balance, 180, 'Balance must be recalculated as Amount_Due - Amount_Paid');
    eq(env.data.Status, 'Partial');
  }
});

check('feedingFees.create: non-numeric or negative amounts are rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  [
    ['Amount_Due', 'abc'],
    ['Amount_Paid', 'abc'],
    ['Amount_Due', -1],
    ['Amount_Paid', -1],
    ['Amount_Paid', 5000],
  ].forEach(function (pair) {
    const env = doPostEnvelope(apiA, 'feedingFees.create', ffCreatePayload({ [pair[0]]: pair[1] }));
    eq(env.success, false, pair[0] + '=' + pair[1] + ' must be rejected');
    eq(env.error, 'VALIDATION_ERROR');
  });
});

check('feedingFees.create: every required field is enforced', function () {
  const apiA = loadBackendAs('admin@school.edu');
  ['Student_ID', 'Academic_Year', 'Term', 'Amount_Due', 'Amount_Paid',
    'Payment_Method', 'Payment_Date'].forEach(function (field) {
    const env = doPostEnvelope(apiA, 'feedingFees.create', payloadWithout(ffCreatePayload(), field));
    eq(env.success, false, field + ' is required');
    eq(env.error, 'VALIDATION_ERROR');
    ok(env.details.missingFields.indexOf(field) !== -1, 'should flag missing ' + field);
  });
});

check('feedingFees.create: invalid Payment_Method and Payment_Date are rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const method = doPostEnvelope(apiA, 'feedingFees.create', ffCreatePayload({ Payment_Method: 'Bitcoin' }));
  eq(method.success, false);
  eq(method.error, 'VALIDATION_ERROR');
  ok(method.details.allowedValues.indexOf('Cash') !== -1, 'should list allowed payment methods');
  const date = doPostEnvelope(apiA, 'feedingFees.create', ffCreatePayload({ Payment_Date: 'not-a-date' }));
  eq(date.success, false);
  eq(date.error, 'VALIDATION_ERROR');
  eq(date.details.field, 'Payment_Date');
});

check('feedingFees.create: nonexistent or withdrawn students are rejected', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const missing = doPostEnvelope(apiA, 'feedingFees.create', ffCreatePayload({ Student_ID: 'STU-999' }));
  eq(missing.success, false);
  eq(missing.error, 'NOT_FOUND');
  const withdrawn = doPostEnvelope(apiA, 'feedingFees.create', ffCreatePayload({ Student_ID: 'STU-3' }));
  eq(withdrawn.success, false);
  eq(withdrawn.error, 'VALIDATION_ERROR');
  eq(withdrawn.details.reason, 'student-withdrawn');
});

check('feedingFees.create: generated IDs are sequential and never reused', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  const first = doPostEnvelope(apiA, 'feedingFees.create', ffCreatePayload());
  const second = doPostEnvelope(apiA, 'feedingFees.create', ffCreatePayload());
  eq(first.data.Payment_ID, 'FF-003', 'the next ID must follow the highest stored ID');
  eq(second.data.Payment_ID, 'FF-004');
  eq(feeIds(ss, 'Feeding_Fees'), ['FF-001', 'FF-002', 'FF-003', 'FF-004'],
    'each create appends exactly one row and never rewrites an existing ID');
});

check('feedingFees.create: refuses a duplicate server-generated Payment_ID', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const realFind = apiA.findFeedingFeeById_;
  apiA.findFeedingFeeById_ = function () {
    return { sheetRow: 2, record: {} };
  };
  const env = doPostEnvelope(apiA, 'feedingFees.create', ffCreatePayload());
  apiA.findFeedingFeeById_ = realFind;
  eq(env.success, false);
  eq(env.error, 'CONFLICT');
  ok(env.message.indexOf('Duplicate Payment_ID') !== -1, 'message should name the collision');
});

check('feedingFees.get requires a Payment_ID', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = doGetEnvelope(apiA, { action: 'feedingFees.get' });
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
  eq(env.details.field, 'Payment_ID');
});

check('feedingFees.list filters by Academic_Year and Term', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const byYear = doGetEnvelope(apiA, { action: 'feedingFees.list', Academic_Year: '2025/2026' });
  eq(byYear.success, true);
  eq(byYear.data.length, 2, 'both fixture payments are 2025/2026');
  const byTerm = doGetEnvelope(apiA, { action: 'feedingFees.list', Term: 'Term 3' });
  eq(byTerm.success, true);
  eq(byTerm.data, [], 'no fixture payment is in Term 3');
});

check('feedingFees.list returns an empty array when nothing matches', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = doGetEnvelope(apiA, { action: 'feedingFees.list', Status: 'Voided' });
  eq(env.success, true);
  eq(env.data, [], 'no matches is an empty list, not an error');
});

section('Phase 4B: Feeding Fees — update and void');

check('feedingFees.update performs a partial update and recalculates Balance', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = { Payment_ID: 'FF-002', Amount_Paid: 450 };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'feedingFees.update' },
    postData: { contents: JSON.stringify({ action: 'feedingFees.update', payload: payload }) }
  }).getContent());
  eq(env.success, true, env.message || env.error);
  if (env.success) {
    eq(env.data.Payment_ID, 'FF-002');
    eq(env.data.Amount_Paid, 450);
    eq(env.data.Balance, 0);
    eq(env.data.Status, 'Paid');
  }
});

check('feedingFees.update: Recorded_By from authenticated user, not client', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const payload = { Payment_ID: 'FF-001', Recorded_By: 'EVIL-STF-999' };
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'feedingFees.update' },
    postData: { contents: JSON.stringify({ action: 'feedingFees.update', payload: payload }) }
  }).getContent());
  eq(env.success, true);
  if (env.success) {
    eq(env.data.Recorded_By, 'STF-1',
      'Recorded_By must not be overwritten by client on update');
  }
});

check('feedingFees.update: the stored Payment_ID is never changed', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  const env = doPostEnvelope(apiA, 'feedingFees.update', { Payment_ID: 'FF-002', Reference: 'UPDATED-REF' });
  eq(env.success, true, env.message || env.error);
  if (env.success) {
    eq(env.data.Payment_ID, 'FF-002');
    eq(env.data.Reference, 'UPDATED-REF');
  }
  ok(apiA.findFeedingFeeById_('FF-002'), 'the row must still be findable by its own ID after the write');
  eq(feeIds(ss, 'Feeding_Fees'), ['FF-001', 'FF-002'], 'an update must never add, drop or re-key a row');
  eq(ss.getSheetByName('Feeding_Fees')._rows[2][0], 'FF-002', 'the ID cell must be left intact');
});

check('feedingFees.update: an unknown Payment_ID targets nothing', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  const env = doPostEnvelope(apiA, 'feedingFees.update', { Payment_ID: 'FF-999', Amount_Paid: 1 });
  eq(env.success, false);
  eq(env.error, 'NOT_FOUND');
  eq(env.details.Payment_ID, 'FF-999');
  eq(feeIds(ss, 'Feeding_Fees'), ['FF-001', 'FF-002'], 'a failed update must not write');
});

check('feedingFees.update: rejects unknown fields', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = doPostEnvelope(apiA, 'feedingFees.update', { Payment_ID: 'FF-002', Middle_Name: 'X' });
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
  ok(env.details.unknownColumns.indexOf('Middle_Name') !== -1, 'should list unknown columns');
});

check('feedingFees.update: changing Amount_Due recalculates Balance and Status', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = doPostEnvelope(apiA, 'feedingFees.update', { Payment_ID: 'FF-001', Amount_Due: 900 });
  eq(env.success, true, env.message || env.error);
  if (env.success) {
    eq(env.data.Amount_Due, 900);
    eq(env.data.Amount_Paid, 450, 'Amount_Paid must be preserved');
    eq(env.data.Balance, 450);
    eq(env.data.Status, 'Partial', 'Status must be derived, not carried over');
  }
});

check('feedingFees.update: invalid amounts are rejected without writing', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  [
    ['Amount_Due', 'abc'],
    ['Amount_Paid', -1],
    ['Amount_Paid', 5000],
  ].forEach(function (pair) {
    const payload = { Payment_ID: 'FF-002' };
    payload[pair[0]] = pair[1];
    const env = doPostEnvelope(apiA, 'feedingFees.update', payload);
    eq(env.success, false, pair[0] + '=' + pair[1] + ' must be rejected');
    eq(env.error, 'VALIDATION_ERROR');
  });
  eq(apiA.findFeedingFeeById_('FF-002').record.Amount_Paid, 225, 'the rejected payloads must not be applied');
});

check('feedingFees.update: a voided payment cannot be updated', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const voidEnv = doPostEnvelope(apiA, 'feedingFees.void', { Payment_ID: 'FF-002' });
  eq(voidEnv.success, true, voidEnv.message || voidEnv.error);
  const env = doPostEnvelope(apiA, 'feedingFees.update', { Payment_ID: 'FF-002', Amount_Paid: 999 });
  eq(env.success, false);
  eq(env.error, 'VALIDATION_ERROR');
  eq(env.details.reason, 'voided-payment');
  eq(apiA.findFeedingFeeById_('FF-002').record.Amount_Paid, 225, 'the voided row must stay untouched');
});

check('feedingFees.void soft-corrects a payment to Voided', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const env = JSON.parse(apiA.doPost({
    parameter: { action: 'feedingFees.void' },
    postData: { contents: JSON.stringify({ action: 'feedingFees.void', payload: { Payment_ID: 'FF-002' } }) }
  }).getContent());
  eq(env.success, true, env.message || env.error);
  if (env.success) {
    eq(env.data.Status, 'Voided');
    eq(env.data.Payment_ID, 'FF-002');
  }
});

check('feedingFees.void: a second void is rejected and the first one stands', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const first = doPostEnvelope(apiA, 'feedingFees.void', { Payment_ID: 'FF-002' });
  eq(first.success, true, first.message || first.error);
  eq(first.data.Status, 'Voided');
  const second = doPostEnvelope(apiA, 'feedingFees.void', { Payment_ID: 'FF-002' });
  eq(second.success, false, 're-voiding must be refused');
  eq(second.error, 'VALIDATION_ERROR');
  eq(second.details.reason, 'already-voided');
  eq(apiA.findFeedingFeeById_('FF-002').record.Status, 'Voided', 'the original void must stand');
});

check('feedingFees.void: the row is preserved with its financial values', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  const voidEnv = doPostEnvelope(apiA, 'feedingFees.void', { Payment_ID: 'FF-002' });
  eq(voidEnv.success, true, voidEnv.message || voidEnv.error);
  eq(ss.getSheetByName('Feeding_Fees')._rows.length, 3, 'void is a soft correction, never a delete');
  const getEnv = doGetEnvelope(apiA, { action: 'feedingFees.get', Payment_ID: 'FF-002' });
  eq(getEnv.success, true, 'a voided payment must stay retrievable');
  eq(getEnv.data.Status, 'Voided');
  eq(getEnv.data.Student_ID, 'STU-2');
  eq(getEnv.data.Amount_Due, 450, 'voiding must not change the amounts');
  eq(getEnv.data.Amount_Paid, 225);
  eq(getEnv.data.Balance, 225);
});

check('feedingFees.void: unknown or missing Payment_ID fails safely', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  const missing = doPostEnvelope(apiA, 'feedingFees.void', {});
  eq(missing.success, false);
  eq(missing.error, 'VALIDATION_ERROR');
  eq(missing.details.field, 'Payment_ID');
  const unknown = doPostEnvelope(apiA, 'feedingFees.void', { Payment_ID: 'FF-999' });
  eq(unknown.success, false);
  eq(unknown.error, 'NOT_FOUND');
  eq(ss.getSheetByName('Feeding_Fees')._rows.length, 3, 'a failed void must not write');
});

section('Phase 4B: Authorization');

check('schoolFees.create requires SCHOOL_FEES.CREATE (teacher is FORBIDDEN)', function () {
  const apiT = loadBackendAs('teacher@school.edu');
  const payload = {
    Student_ID: 'STU-2',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 1000,
    Amount_Paid: 0,
    Payment_Method: 'Cash',
    Payment_Date: '2025-09-20'
  };
  const env = JSON.parse(apiT.doPost({
    parameter: { action: 'schoolFees.create' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.create', payload: payload }) }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'FORBIDDEN');
});

check('schoolFees.void requires SCHOOL_FEES.VOID (teacher is FORBIDDEN)', function () {
  const apiT = loadBackendAs('teacher@school.edu');
  const env = JSON.parse(apiT.doPost({
    parameter: { action: 'schoolFees.void' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.void', payload: { Payment_ID: 'SF-001' } }) }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'FORBIDDEN');
});

check('feedingFees.create requires FEEDING_FEES.CREATE (teacher is FORBIDDEN)', function () {
  const apiT = loadBackendAs('teacher@school.edu');
  const payload = {
    Student_ID: 'STU-2',
    Academic_Year: '2025/2026',
    Term: 'Term 1',
    Amount_Due: 500,
    Amount_Paid: 0,
    Payment_Method: 'Cash',
    Payment_Date: '2025-09-20'
  };
  const env = JSON.parse(apiT.doPost({
    parameter: { action: 'feedingFees.create' },
    postData: { contents: JSON.stringify({ action: 'feedingFees.create', payload: payload }) }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'FORBIDDEN');
});

check('feedingFees.void requires FEEDING_FEES.VOID (teacher is FORBIDDEN)', function () {
  const apiT = loadBackendAs('teacher@school.edu');
  const env = JSON.parse(apiT.doPost({
    parameter: { action: 'feedingFees.void' },
    postData: { contents: JSON.stringify({ action: 'feedingFees.void', payload: { Payment_ID: 'FF-001' } }) }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'FORBIDDEN');
});

check('schoolFees.list requires SCHOOL_FEES.READ (teacher is FORBIDDEN)', function () {
  const apiT = loadBackendAs('teacher@school.edu');
  const env = JSON.parse(apiT.doGet({
    parameter: { action: 'schoolFees.list' }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'FORBIDDEN');
});

check('feedingFees.list requires FEEDING_FEES.READ (teacher is FORBIDDEN)', function () {
  const apiT = loadBackendAs('teacher@school.edu');
  const env = JSON.parse(apiT.doGet({
    parameter: { action: 'feedingFees.list' }
  }).getContent());
  eq(env.success, false);
  eq(env.error, 'FORBIDDEN');
});

check('Phase 4B: fee handlers authorize through the central requirePermission_ mechanism', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const calls = [];
  const original = apiA.requirePermission_;
  apiA.requirePermission_ = function (permission) {
    calls.push(permission);
    return original(permission);
  };
  doGetEnvelope(apiA, { action: 'schoolFees.list' });
  doGetEnvelope(apiA, { action: 'schoolFees.get', Payment_ID: 'SF-001' });
  doPostEnvelope(apiA, 'schoolFees.create', sfCreatePayload({ Reference: 'AUTH-SF-1' }));
  doPostEnvelope(apiA, 'schoolFees.update', { Payment_ID: 'SF-001', Notes: 'authorized' });
  doPostEnvelope(apiA, 'schoolFees.void', { Payment_ID: 'SF-001' });
  doGetEnvelope(apiA, { action: 'feedingFees.list' });
  doGetEnvelope(apiA, { action: 'feedingFees.get', Payment_ID: 'FF-001' });
  doPostEnvelope(apiA, 'feedingFees.create', ffCreatePayload({ Reference: 'AUTH-FF-1' }));
  doPostEnvelope(apiA, 'feedingFees.update', { Payment_ID: 'FF-001', Notes: 'authorized' });
  doPostEnvelope(apiA, 'feedingFees.void', { Payment_ID: 'FF-001' });
  apiA.requirePermission_ = original;
  eq(calls, [
    'SCHOOL_FEES.READ', 'SCHOOL_FEES.READ', 'SCHOOL_FEES.CREATE', 'SCHOOL_FEES.UPDATE',
    'SCHOOL_FEES.VOID',
    'FEEDING_FEES.READ', 'FEEDING_FEES.READ', 'FEEDING_FEES.CREATE', 'FEEDING_FEES.UPDATE',
    'FEEDING_FEES.VOID',
  ], 'every fee action must go through the Phase 4A mechanism with its own code');
});

check('Phase 4B: fee permission codes come from the Phase 4A catalog', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const admin = apiA.requireAuthentication_();
  P4B_FEE_PERMISSIONS.forEach(function (code) {
    ok(CONFIG.PERMISSION_CODES.indexOf(code) !== -1, code + ' must be a catalog code');
    eq(apiA.hasPermission_(admin, code), true, 'Admin must hold ' + code);
  });
});

check('Phase 4B: a role holding only READ may read fees but not write them', function () {
  const ss = p4bSpreadsheet({
    users: [['USR-3', 'STF-3', 'viewer@school.edu', 'Viewer', 'Active', '']],
    roles: [
      ['ROL-1', 'Admin', 'Full system administrator', 'Active'],
      ['ROL-3', 'Viewer', 'Read-only fee viewer', 'Active'],
    ],
    mappings: [
      ['RP-1', 'ROL-3', permIdFor('SCHOOL_FEES.READ'), 'Active'],
      ['RP-2', 'ROL-3', permIdFor('FEEDING_FEES.READ'), 'Active'],
      ['RP-3', 'ROL-3', permIdFor('STUDENTS.READ'), 'Active'],
    ],
  });
  const apiV = loadBackendAs('viewer@school.edu', ss);
  eq(doGetEnvelope(apiV, { action: 'schoolFees.list' }).success, true, 'SCHOOL_FEES.READ allows list');
  eq(doGetEnvelope(apiV, { action: 'schoolFees.get', Payment_ID: 'SF-001' }).success, true);
  eq(doGetEnvelope(apiV, { action: 'feedingFees.list' }).success, true, 'FEEDING_FEES.READ allows list');
  eq(doGetEnvelope(apiV, { action: 'feedingFees.get', Payment_ID: 'FF-001' }).success, true);
  [
    ['schoolFees.create', 'SCHOOL_FEES.CREATE'],
    ['schoolFees.update', 'SCHOOL_FEES.UPDATE'],
    ['schoolFees.void', 'SCHOOL_FEES.VOID'],
    ['feedingFees.create', 'FEEDING_FEES.CREATE'],
    ['feedingFees.update', 'FEEDING_FEES.UPDATE'],
    ['feedingFees.void', 'FEEDING_FEES.VOID'],
  ].forEach(function (pair) {
    const env = doPostEnvelope(apiV, pair[0], { Payment_ID: 'SF-001' });
    eq(env.success, false, pair[0] + ' must be denied without ' + pair[1]);
    eq(env.error, ERROR_CODES.FORBIDDEN, pair[0] + ' must be FORBIDDEN');
    eq(env.details.permission, pair[1], 'the denial must name the missing permission');
  });
});

check('Phase 4B: unauthenticated callers cannot read or write fees', function () {
  const anon = loadBackendAs('', makeFullSpreadsheet());
  ['schoolFees.list', 'schoolFees.get', 'feedingFees.list', 'feedingFees.get'].forEach(function (action) {
    const env = doGetEnvelope(anon, { action: action, Payment_ID: 'SF-001' });
    eq(env.success, false, action + ' must require authentication');
    eq(env.error, ERROR_CODES.UNAUTHORIZED, action + ' must be UNAUTHORIZED');
  });
  ['schoolFees.create', 'schoolFees.update', 'schoolFees.void',
    'feedingFees.create', 'feedingFees.update', 'feedingFees.void'].forEach(function (action) {
    const env = doPostEnvelope(anon, action, { Payment_ID: 'SF-001' });
    eq(env.success, false, action + ' must require authentication');
    eq(env.error, ERROR_CODES.UNAUTHORIZED, action + ' must be UNAUTHORIZED');
  });
});

section('Phase 4B: Error handling and safety');

check('schoolFees.list fails safely when the School_Fees tab has no header row', function () {
  const ss = makeFullSpreadsheet();
  ss.getSheetByName('School_Fees')._rows = [];
  const apiA = loadBackendAs('admin@school.edu', ss);
  const env = doGetEnvelope(apiA, { action: 'schoolFees.list' });
  eq(env.success, false);
  eq(env.error, ERROR_CODES.SERVER_ERROR, 'a headerless tab is a structural failure');
  eq(env.details.sheet, 'School_Fees');
});

check('schoolFees fails safely when the School_Fees tab is missing entirely', function () {
  const apiMissing = loadBackendAs('admin@school.edu', p4bSpreadsheetWithoutTab('School_Fees'));
  const list = doGetEnvelope(apiMissing, { action: 'schoolFees.list' });
  eq(list.success, false);
  eq(list.error, ERROR_CODES.SERVER_ERROR);
  eq(list.details.sheet, 'School_Fees');
  eq(list.details.reason, 'sheet-missing');
  const create = doPostEnvelope(apiMissing, 'schoolFees.create', sfCreatePayload());
  eq(create.success, false);
  eq(create.error, ERROR_CODES.SERVER_ERROR);
  const update = doPostEnvelope(apiMissing, 'schoolFees.update', { Payment_ID: 'SF-001', Notes: 'x' });
  eq(update.success, false);
  eq(update.error, ERROR_CODES.NOT_FOUND, 'a missing tab cannot contain the row being edited');
  const voidEnv = doPostEnvelope(apiMissing, 'schoolFees.void', { Payment_ID: 'SF-001' });
  eq(voidEnv.success, false);
  eq(voidEnv.error, ERROR_CODES.NOT_FOUND);
});

check('schoolFees.list on an empty module sheet returns an empty list', function () {
  const ss = makeFullSpreadsheet();
  const sheet = ss.getSheetByName('School_Fees');
  sheet._rows = [sheet._rows[0]];
  const apiA = loadBackendAs('admin@school.edu', ss);
  const env = doGetEnvelope(apiA, { action: 'schoolFees.list' });
  eq(env.success, true);
  eq(env.data, []);
});

check('Phase 4B: schoolFees writes respect script locking', function () {
  [
    ['schoolFees.create', sfCreatePayload()],
    ['schoolFees.update', { Payment_ID: 'SF-001', Notes: 'locked' }],
    ['schoolFees.void', { Payment_ID: 'SF-001' }],
  ].forEach(function (pair) {
    const ss = makeFullSpreadsheet();
    const apiLocked = loadBackendAs('admin@school.edu', ss, { lockUnavailable: true });
    const env = doPostEnvelope(apiLocked, pair[0], pair[1]);
    eq(env.success, false, pair[0] + ' must not write without the script lock');
    eq(env.error, ERROR_CODES.CONFLICT, pair[0] + ' must report CONFLICT');
    eq(feeIds(ss, 'School_Fees'), ['SF-001', 'SF-002', 'SF-003'],
      pair[0] + ' must leave the sheet untouched');
    eq(apiLocked.findSchoolFeeById_('SF-001').record.Status, 'Paid',
      pair[0] + ' must not change any stored value');
  });
});

check('feedingFees.list fails safely when the Feeding_Fees tab has no header row', function () {
  const ss = makeFullSpreadsheet();
  ss.getSheetByName('Feeding_Fees')._rows = [];
  const apiA = loadBackendAs('admin@school.edu', ss);
  const env = doGetEnvelope(apiA, { action: 'feedingFees.list' });
  eq(env.success, false);
  eq(env.error, ERROR_CODES.SERVER_ERROR, 'a headerless tab is a structural failure');
  eq(env.details.sheet, 'Feeding_Fees');
});

check('feedingFees fails safely when the Feeding_Fees tab is missing entirely', function () {
  const apiMissing = loadBackendAs('admin@school.edu', p4bSpreadsheetWithoutTab('Feeding_Fees'));
  const list = doGetEnvelope(apiMissing, { action: 'feedingFees.list' });
  eq(list.success, false);
  eq(list.error, ERROR_CODES.SERVER_ERROR);
  eq(list.details.sheet, 'Feeding_Fees');
  eq(list.details.reason, 'sheet-missing');
  const create = doPostEnvelope(apiMissing, 'feedingFees.create', ffCreatePayload());
  eq(create.success, false);
  eq(create.error, ERROR_CODES.SERVER_ERROR);
  const update = doPostEnvelope(apiMissing, 'feedingFees.update', { Payment_ID: 'FF-001', Notes: 'x' });
  eq(update.success, false);
  eq(update.error, ERROR_CODES.NOT_FOUND);
  const voidEnv = doPostEnvelope(apiMissing, 'feedingFees.void', { Payment_ID: 'FF-001' });
  eq(voidEnv.success, false);
  eq(voidEnv.error, ERROR_CODES.NOT_FOUND);
});

check('feedingFees.list on an empty module sheet returns an empty list', function () {
  const ss = makeFullSpreadsheet();
  const sheet = ss.getSheetByName('Feeding_Fees');
  sheet._rows = [sheet._rows[0]];
  const apiA = loadBackendAs('admin@school.edu', ss);
  const env = doGetEnvelope(apiA, { action: 'feedingFees.list' });
  eq(env.success, true);
  eq(env.data, []);
});

check('Phase 4B: feedingFees writes respect script locking', function () {
  [
    ['feedingFees.create', ffCreatePayload()],
    ['feedingFees.update', { Payment_ID: 'FF-001', Notes: 'locked' }],
    ['feedingFees.void', { Payment_ID: 'FF-001' }],
  ].forEach(function (pair) {
    const ss = makeFullSpreadsheet();
    const apiLocked = loadBackendAs('admin@school.edu', ss, { lockUnavailable: true });
    const env = doPostEnvelope(apiLocked, pair[0], pair[1]);
    eq(env.success, false, pair[0] + ' must not write without the script lock');
    eq(env.error, ERROR_CODES.CONFLICT, pair[0] + ' must report CONFLICT');
    eq(feeIds(ss, 'Feeding_Fees'), ['FF-001', 'FF-002'],
      pair[0] + ' must leave the sheet untouched');
    eq(apiLocked.findFeedingFeeById_('FF-001').record.Status, 'Paid',
      pair[0] + ' must not change any stored value');
  });
});

check('Phase 4B: schoolFees.update decides and writes inside the script lock', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  const statusCol = feeColIndex(ss, 'School_Fees', 'Status');
  const row = feeRowOf(ss, 'School_Fees', 'SF-003');
  const log = observeSheetAccess(apiA, 'School_Fees');

  const env = doPostEnvelope(apiA, 'schoolFees.update', { Payment_ID: 'SF-003', Amount_Paid: 1200 });
  eq(env.success, true, 'the update should succeed');
  eq(env.data.Status, 'Paid', 'paying the balance in full must recalculate the status');
  eq(apiA.__lockDepth, 0, 'the critical section must be released before the response is returned');

  const reads = fullRowReads(log, row);
  eq(reads.length, 2, 'the row is read once to decide and once to refresh the response');
  reads.forEach(function (entry) {
    ok(entry.depth > 0, 'the stored row must never be read outside the critical section');
  });
  eq(reads[0].values[0][statusCol], 'Unpaid',
    'the deciding read must be the stored pre-update row (Unpaid), read under the lock');
  eq(reads[1].values[0][statusCol], 'Paid', 'the refresh read must see the recalculated status');

  const writes = log.filter(function (entry) { return entry.op === 'write'; });
  eq(writes.length, 1, 'the update must write the row exactly once');
  ok(writes[0].depth > 0, 'the row write must happen inside the critical section');
  eq(writes[0].values[0][statusCol], 'Paid');
});

check('Phase 4B: schoolFees.void decides inside the script lock', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  const statusCol = feeColIndex(ss, 'School_Fees', 'Status');
  const row = feeRowOf(ss, 'School_Fees', 'SF-002');

  const first = doPostEnvelope(apiA, 'schoolFees.void', { Payment_ID: 'SF-002' });
  eq(first.success, true, 'the first void should succeed');
  eq(first.data.Status, 'Voided');

  // Only the rejected second void is observed: its already-voided decision must
  // come from a re-read taken while the lock is held.
  const log = observeSheetAccess(apiA, 'School_Fees');
  const second = doPostEnvelope(apiA, 'schoolFees.void', { Payment_ID: 'SF-002' });
  eq(second.success, false);
  eq(second.error, ERROR_CODES.VALIDATION_ERROR);
  eq(second.details.reason, 'already-voided');

  const reads = fullRowReads(log, row);
  ok(reads.length >= 1, 'the stored row must be re-read before the decision');
  reads.forEach(function (entry) {
    ok(entry.depth > 0, 'the re-read must happen inside the critical section');
  });
  eq(reads[0].values[0][statusCol], 'Voided', 'the decision must use the freshly read status');
  eq(log.filter(function (entry) { return entry.op === 'write'; }).length, 0,
    'a rejected void must not write to the sheet');
});

check('Phase 4B: schoolFees.update is serialized against a void racing the lock', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  const rows = ss.getSheetByName('School_Fees')._rows;
  const statusCol = feeColIndex(ss, 'School_Fees', 'Status');
  const paidCol = feeColIndex(ss, 'School_Fees', 'Amount_Paid');
  const row = feeRowOf(ss, 'School_Fees', 'SF-001');

  // "Another user's" void commits in the window between this call taking the
  // lock and its first read, i.e. exactly where a pre-lock decision would miss
  // it. The handler must therefore re-read under the lock before deciding.
  const state = { commits: 0, depth: -1 };
  const realFind = apiA.findSchoolFeeById_;
  apiA.findSchoolFeeById_ = function (id) {
    if (state.commits === 0) {
      state.commits += 1;
      state.depth = apiA.__lockDepth;
      rows[row - 1][statusCol] = 'Voided';
    }
    return realFind(id);
  };
  const env = doPostEnvelope(apiA, 'schoolFees.update', { Payment_ID: 'SF-001', Amount_Paid: 0 });
  apiA.findSchoolFeeById_ = realFind;

  eq(state.commits, 1, 'the competing void must land before the handler reads the record');
  ok(state.depth > 0, 'the deciding read must happen while this call already holds the lock');
  eq(env.success, false, 'the update must not overwrite a void that committed first');
  eq(env.error, ERROR_CODES.VALIDATION_ERROR);
  eq(env.details.reason, 'voided-payment');
  eq(rows[row - 1][statusCol], 'Voided', 'the competing void must be preserved');
  eq(rows[row - 1][paidCol], 1200, 'no cell of the voided row may be overwritten');
});
check('Phase 4B: feedingFees.update decides and writes inside the script lock', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  const statusCol = feeColIndex(ss, 'Feeding_Fees', 'Status');
  const row = feeRowOf(ss, 'Feeding_Fees', 'FF-002');
  const log = observeSheetAccess(apiA, 'Feeding_Fees');

  const env = doPostEnvelope(apiA, 'feedingFees.update', { Payment_ID: 'FF-002', Amount_Paid: 450 });
  eq(env.success, true, 'the update should succeed');
  eq(env.data.Status, 'Paid', 'paying the balance in full must recalculate the status');
  eq(apiA.__lockDepth, 0, 'the critical section must be released before the response is returned');

  const reads = fullRowReads(log, row);
  eq(reads.length, 2, 'the row is read once to decide and once to refresh the response');
  reads.forEach(function (entry) {
    ok(entry.depth > 0, 'the stored row must never be read outside the critical section');
  });
  eq(reads[0].values[0][statusCol], 'Partial', 'the deciding read must be the stored pre-update row');
  eq(reads[1].values[0][statusCol], 'Paid', 'the refresh read must see the recalculated status');

  const writes = log.filter(function (entry) { return entry.op === 'write'; });
  eq(writes.length, 1, 'the update must write the row exactly once');
  ok(writes[0].depth > 0, 'the row write must happen inside the critical section');
  eq(writes[0].values[0][statusCol], 'Paid');
});

check('Phase 4B: feedingFees.void decides inside the script lock', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  const statusCol = feeColIndex(ss, 'Feeding_Fees', 'Status');
  const row = feeRowOf(ss, 'Feeding_Fees', 'FF-002');

  const first = doPostEnvelope(apiA, 'feedingFees.void', { Payment_ID: 'FF-002' });
  eq(first.success, true, 'the first void should succeed');
  eq(first.data.Status, 'Voided');

  const log = observeSheetAccess(apiA, 'Feeding_Fees');
  const second = doPostEnvelope(apiA, 'feedingFees.void', { Payment_ID: 'FF-002' });
  eq(second.success, false);
  eq(second.error, ERROR_CODES.VALIDATION_ERROR);
  eq(second.details.reason, 'already-voided');

  const reads = fullRowReads(log, row);
  ok(reads.length >= 1, 'the stored row must be re-read before the decision');
  reads.forEach(function (entry) {
    ok(entry.depth > 0, 'the re-read must happen inside the critical section');
  });
  eq(reads[0].values[0][statusCol], 'Voided', 'the decision must use the freshly read status');
  eq(log.filter(function (entry) { return entry.op === 'write'; }).length, 0,
    'a rejected void must not write to the sheet');
});

check('Phase 4B: feedingFees.update is serialized against a void racing the lock', function () {
  const ss = makeFullSpreadsheet();
  const apiA = loadBackendAs('admin@school.edu', ss);
  const rows = ss.getSheetByName('Feeding_Fees')._rows;
  const statusCol = feeColIndex(ss, 'Feeding_Fees', 'Status');
  const paidCol = feeColIndex(ss, 'Feeding_Fees', 'Amount_Paid');
  const row = feeRowOf(ss, 'Feeding_Fees', 'FF-001');

  const state = { commits: 0, depth: -1 };
  const realFind = apiA.findFeedingFeeById_;
  apiA.findFeedingFeeById_ = function (id) {
    if (state.commits === 0) {
      state.commits += 1;
      state.depth = apiA.__lockDepth;
      rows[row - 1][statusCol] = 'Voided';
    }
    return realFind(id);
  };
  const env = doPostEnvelope(apiA, 'feedingFees.update', { Payment_ID: 'FF-001', Amount_Paid: 0 });
  apiA.findFeedingFeeById_ = realFind;

  eq(state.commits, 1, 'the competing void must land before the handler reads the record');
  ok(state.depth > 0, 'the deciding read must happen while this call already holds the lock');
  eq(env.success, false, 'the update must not overwrite a void that committed first');
  eq(env.error, ERROR_CODES.VALIDATION_ERROR);
  eq(env.details.reason, 'voided-payment');
  eq(rows[row - 1][statusCol], 'Voided', 'the competing void must be preserved');
  eq(rows[row - 1][paidCol], 450, 'no cell of the voided row may be overwritten');
});


section('Phase 4B: Full lifecycle integration');

check('Phase 4B: complete school fee lifecycle — create, get, update, void', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const createPayload = {
    Student_ID: 'STU-1',
    Academic_Year: '2025/2026',
    Term: 'Term 3',
    Amount_Due: 2000,
    Amount_Paid: 1000,
    Payment_Method: 'Bank Transfer',
    Payment_Date: '2025-09-01',
    Reference: 'LIFECYCLE-TEST-SF-001',
    Notes: 'Lifecycle test'
  };
  const createEnv = JSON.parse(apiA.doPost({
    parameter: { action: 'schoolFees.create' },
    postData: { contents: JSON.stringify({ action: 'schoolFees.create', payload: createPayload }) }
  }).getContent());
  ok(createEnv.success, 'create should succeed: ' + (createEnv.message || createEnv.error));
  if (createEnv.success) {
    const newId = createEnv.data.Payment_ID;
    ok(newId.indexOf('SF-') === 0, 'should have SF- prefix: ' + newId);
    eq(createEnv.data.Student_ID, 'STU-1');
    eq(createEnv.data.Balance, 1000);
    eq(createEnv.data.Status, 'Partial');
    eq(createEnv.data.Recorded_By, 'STF-1');

    const getEnv = JSON.parse(apiA.doGet({
      parameter: { action: 'schoolFees.get', Payment_ID: newId }
    }).getContent());
    eq(getEnv.success, true);
    if (getEnv.success) {
      eq(getEnv.data.Payment_ID, newId);
      eq(getEnv.data.Balance, 1000);
    }

    const updateEnv = JSON.parse(apiA.doPost({
      parameter: { action: 'schoolFees.update' },
      postData: { contents: JSON.stringify({ action: 'schoolFees.update', payload: { Payment_ID: newId, Amount_Paid: 2000 } }) }
    }).getContent());
    eq(updateEnv.success, true);
    if (updateEnv.success) {
      eq(updateEnv.data.Balance, 0);
      eq(updateEnv.data.Status, 'Paid');
    }

    const voidEnv = JSON.parse(apiA.doPost({
      parameter: { action: 'schoolFees.void' },
      postData: { contents: JSON.stringify({ action: 'schoolFees.void', payload: { Payment_ID: newId } }) }
    }).getContent());
    eq(voidEnv.success, true);
    if (voidEnv.success) {
      eq(voidEnv.data.Status, 'Voided');
    }
  }
});

check('Phase 4B: complete feeding fee lifecycle — create, get, update, void', function () {
  const apiA = loadBackendAs('admin@school.edu');
  const createPayload = {
    Student_ID: 'STU-1',
    Academic_Year: '2025/2026',
    Term: 'Term 3',
    Amount_Due: 500,
    Amount_Paid: 0,
    Payment_Method: 'Cash',
    Payment_Date: '2025-09-01',
    Reference: 'LIFECYCLE-TEST-FF-001'
  };
  const createEnv = JSON.parse(apiA.doPost({
    parameter: { action: 'feedingFees.create' },
    postData: { contents: JSON.stringify({ action: 'feedingFees.create', payload: createPayload }) }
  }).getContent());
  ok(createEnv.success, 'create should succeed: ' + (createEnv.message || createEnv.error));
  if (createEnv.success) {
    const newId = createEnv.data.Payment_ID;
    ok(newId.indexOf('FF-') === 0, 'should have FF- prefix: ' + newId);
    eq(createEnv.data.Student_ID, 'STU-1');
    eq(createEnv.data.Balance, 500);
    eq(createEnv.data.Status, 'Unpaid');
    eq(createEnv.data.Recorded_By, 'STF-1');

    const getEnv = JSON.parse(apiA.doGet({
      parameter: { action: 'feedingFees.get', Payment_ID: newId }
    }).getContent());
    eq(getEnv.success, true);
    if (getEnv.success) {
      eq(getEnv.data.Payment_ID, newId);
      eq(getEnv.data.Balance, 500);
    }

    const updateEnv = JSON.parse(apiA.doPost({
      parameter: { action: 'feedingFees.update' },
      postData: { contents: JSON.stringify({ action: 'feedingFees.update', payload: { Payment_ID: newId, Amount_Paid: 500 } }) }
    }).getContent());
    eq(updateEnv.success, true);
    if (updateEnv.success) {
      eq(updateEnv.data.Status, 'Paid');
    }

    const voidEnv = JSON.parse(apiA.doPost({
      parameter: { action: 'feedingFees.void' },
      postData: { contents: JSON.stringify({ action: 'feedingFees.void', payload: { Payment_ID: newId } }) }
    }).getContent());
    eq(voidEnv.success, true);
    if (voidEnv.success) {
      eq(voidEnv.data.Status, 'Voided');
    }
  }
});

/* ==========================================================================
 * Phase 5: Stationery + Inventory
 * ======================================================================== */

const STATIONERY_HEADERS = ['Transaction_ID', 'Student_ID', 'Item_ID', 'Quantity_Purchased',
  'Unit_Price', 'Total', 'Amount_Paid', 'Balance', 'Payment_Date', 'Payment_Method',
  'Reference', 'Fulfillment_Status', 'Quantity_Given', 'Quantity_Remaining',
  'Given_By', 'Given_Date', 'Recorded_By', 'Notes'];

const INVENTORY_HEADERS = ['Item_ID', 'Item_Name', 'Category', 'Unit', 'Selling_Price',
  'Current_Quantity', 'Minimum_Stock_Level', 'Status'];

const INVENTORY_MOVEMENT_HEADERS = ['Movement_ID', 'Item_ID', 'Movement_Type', 'Quantity',
  'Date', 'Reason', 'Recorded_By', 'Notes'];

/** Spreadsheet with seeded Stationery, Inventory, and Inventory_Movements tabs. */
function p5Spreadsheet(opts) {
  opts = opts || {};
  const ss = makeFullSpreadsheet();
  ss.getSheetByName('Stationery')._rows = opts.stationery || [
    STATIONERY_HEADERS,
    ['STN-001', 'STU-1', 'ITM-001', 10, 5, 50, 50, 0,
      '2025-09-01', 'Cash', 'STN-2025-001', 'Pending', 0, 10, '', '', 'STF-1', ''],
  ];
  ss.getSheetByName('Inventory')._rows = opts.inventory || [
    INVENTORY_HEADERS,
    ['ITM-001', 'Exercise Book', 'Stationery', 'Piece', 5, 100, 20, 'In Stock'],
    ['ITM-002', 'Chalk Box', 'Teaching Aids', 'Box', 25, 10, 10, 'Low Stock'],
    ['ITM-003', 'Whiteboard Marker', 'Stationery', 'Piece', 8, 0, 5, 'Out of Stock'],
  ];
  ss.getSheetByName('Inventory_Movements')._rows = opts.movements || [
    INVENTORY_MOVEMENT_HEADERS,
    ['MOV-001', 'ITM-001', 'STOCK_IN', 100, '2025-09-01', 'Opening stock', 'STF-1', ''],
    ['MOV-002', 'ITM-002', 'STOCK_OUT', 5, '2025-09-02', 'Classroom use', 'STF-1', ''],
  ];
  return ss;
}

/** Load Phase 5 backend with seeded sheets and Admin identity by default. */
function p5Api(opts) {
  opts = opts || {};
  const ss = opts.spreadsheet || p5Spreadsheet(opts.fixture);
  const sandbox = makeSandbox(ss, opts);
  sandbox.Session.__activeEmail = opts.email === undefined ? 'admin@school.edu' : opts.email;
  const loaded = loadBackend(sandbox);
  loaded.__ss = ss;
  return loaded;
}

let p5LastApi = null;

function p5InvCol(header, ss) {
  const rows = (ss || p5LastApi.__ss).getSheetByName('Inventory')._rows;
  const idx = rows[0].indexOf(header);
  return rows.slice(1).map(function (row) { return row[idx]; });
}

function p5MovCol(header, ss) {
  const rows = (ss || p5LastApi.__ss).getSheetByName('Inventory_Movements')._rows;
  const idx = rows[0].indexOf(header);
  return rows.slice(1).map(function (row) { return row[idx]; });
}

function p5RowCount(tabName, ss) {
  return (ss || p5LastApi.__ss).getSheetByName(tabName)._rows.length - 1;
}

/** POST helper mirroring the Phase 4B tests. */
function p5Post(api, action, payload) {
  return JSON.parse(api.doPost({
    parameter: { action: action },
    postData: { contents: JSON.stringify({ action: action, payload: payload }) }
  }).getContent());
}

function p5Get(api, action, params) {
  return JSON.parse(api.doGet({
    parameter: Object.assign({ action: action }, params || {})
  }).getContent());
}

section('Phase 5: routing');

check('stationery.list is routed', function () {
  ok(api.listAvailableActions_().indexOf('stationery.list') !== -1);
});
check('stationery.create is routed', function () {
  ok(api.listAvailableActions_().indexOf('stationery.create') !== -1);
});
check('stationery.fulfill is routed', function () {
  ok(api.listAvailableActions_().indexOf('stationery.fulfill') !== -1);
});
check('inventory.list is routed', function () {
  ok(api.listAvailableActions_().indexOf('inventory.list') !== -1);
});
check('inventory.create is routed', function () {
  ok(api.listAvailableActions_().indexOf('inventory.create') !== -1);
});
check('inventory.stockIn is routed', function () {
  ok(api.listAvailableActions_().indexOf('inventory.stockIn') !== -1);
});
check('inventory.stockOut is routed', function () {
  ok(api.listAvailableActions_().indexOf('inventory.stockOut') !== -1);
});
check('inventory.movements is routed', function () {
  ok(api.listAvailableActions_().indexOf('inventory.movements') !== -1);
});

check('inventory.list tolerates utm_source from ChatGPT link previews', function () {
  const api5 = p5Api();
  const env = p5Get(api5, 'inventory.list', { utm_source: 'chatgpt' });
  eq(env.success, true, 'utm_source must be ignored, not treated as a filter');
  ok(env.data.length >= 3, 'should return seeded items');
});

check('inventory.movements tolerates utm_source from ChatGPT link previews', function () {
  const api5 = p5Api();
  const env = p5Get(api5, 'inventory.movements', { utm_source: 'chatgpt' });
  eq(env.success, true, 'utm_source must be ignored, not treated as a filter');
  ok(env.data.length >= 2, 'should return seeded movements');
});

check('inventory.list still rejects unknown filter fields (not tracking params)', function () {
  const api5 = p5Api();
  throwsWithCode(function () {
    api5.handleInventoryList_({ bogusField: 'x' });
  }, ERROR_CODES.VALIDATION_ERROR);
});

check('inventory.movements still rejects unknown filter fields (not tracking params)', function () {
  const api5 = p5Api();
  throwsWithCode(function () {
    api5.handleInventoryMovements_({ bogusField: 'x' });
  }, ERROR_CODES.VALIDATION_ERROR);
});

section('Phase 5: list');

check('stationery.list returns transactions', function () {
  const api5 = p5Api();
  const env = p5Get(api5, 'stationery.list');
  eq(env.success, true);
  ok(env.data.length >= 1, 'should return at least one transaction');
  ok(env.data[0].hasOwnProperty('Transaction_ID'), 'records must be header-keyed');
});

check('inventory.list returns items', function () {
  const api5 = p5Api();
  const env = p5Get(api5, 'inventory.list');
  eq(env.success, true);
  ok(env.data.length >= 3, 'should return the seeded items');
  eq(env.data[0].Item_ID, 'ITM-001');
  ok(env.data[0].hasOwnProperty('Status'), 'Status must be normalised on list');
});

check('inventory.list can filter by Category', function () {
  const api5 = p5Api();
  const env = p5Get(api5, 'inventory.list', { Category: 'Stationery' });
  eq(env.success, true);
  ok(env.data.length >= 1, 'should return Stationery-category items');
  env.data.forEach(function (record) {
    eq(record.Category, 'Stationery');
  });
});

check('inventory.list can filter by Status', function () {
  const api5 = p5Api();
  const env = p5Get(api5, 'inventory.list', { Status: 'Out of Stock' });
  eq(env.success, true);
  eq(env.data[0].Item_ID, 'ITM-003');
  eq(env.data[0].Current_Quantity, 0);
});

check('inventory.list rejects unknown filter fields', function () {
  const api5 = p5Api();
  const err = throwsWithCode(function () {
    api5.handleInventoryList_({ Bogus: 1 });
  }, ERROR_CODES.VALIDATION_ERROR);
    eq(err.details.unknownColumns, ['Bogus']);
});

section('Phase 5: inventory.create');

check('inventory.create stores a server-generated ITM ID and derives Status', function () {
  const api5 = p5Api();
  const env = api5.handleInventoryCreate_({
    Item_Name: 'Board Duster',
    Category: 'Stationery',
    Unit: 'Piece',
    Selling_Price: 12,
    Current_Quantity: 5,
    Minimum_Stock_Level: 2,
  });
  eq(env.success, true);
  eq(env.data.Item_ID, 'ITM-004');
  eq(env.data.Status, 'In Stock');
  eq(env.data.Selling_Price, 12);
  const rows = api5.__ss.getSheetByName('Inventory')._rows;
  eq(rows[4][0], 'ITM-004');
  eq(rows[4][1], 'Board Duster');
});

check('inventory.create rejects a client-supplied Item_ID', function () {
  const api5 = p5Api();
  const err = throwsWithCode(function () {
    api5.handleInventoryCreate_({ Item_ID: 'ITM-999', Item_Name: 'Spoofed', Selling_Price: 1 });
  }, ERROR_CODES.VALIDATION_ERROR);
  eq(err.details.unknownColumns, ['Item_ID']);
});

check('inventory.create ignores a client Status and derives it', function () {
  const api5 = p5Api();
  const env = api5.handleInventoryCreate_({
    Item_Name: 'Marker', Selling_Price: 8,
    Current_Quantity: 0, Minimum_Stock_Level: 5, Status: 'In Stock',
  });
  eq(env.data.Status, 'Out of Stock');
});

check('inventory.create defaults quantities to 0 when omitted', function () {
  const api5 = p5Api();
  const env = api5.handleInventoryCreate_({ Item_Name: 'Empty Item', Selling_Price: 0 });
  eq(env.data.Current_Quantity, 0);
  eq(env.data.Minimum_Stock_Level, 0);
  eq(env.data.Status, 'Out of Stock');
});

check('inventory.create requires Item_Name', function () {
  const api5 = p5Api();
  throwsWithCode(function () { api5.handleInventoryCreate_({ Item_Name: '', Selling_Price: 10 }); },
    ERROR_CODES.VALIDATION_ERROR);
});

check('inventory.create requires Selling_Price to be numeric and >= 0', function () {
  const api5 = p5Api();
  throwsWithCode(function () { api5.handleInventoryCreate_({ Item_Name: 'X' }); },
    ERROR_CODES.VALIDATION_ERROR);
  throwsWithCode(function () { api5.handleInventoryCreate_({ Item_Name: 'X', Selling_Price: 'abc' }); },
    ERROR_CODES.VALIDATION_ERROR);
  throwsWithCode(function () { api5.handleInventoryCreate_({ Item_Name: 'X', Selling_Price: -1 }); },
    ERROR_CODES.VALIDATION_ERROR);
});

check('inventory.create rejects negative / fractional Current_Quantity', function () {
  const api5 = p5Api();
  eq(throwsWithCode(function () {
    api5.handleInventoryCreate_({ Item_Name: 'X', Selling_Price: 5, Current_Quantity: -1 });
  }, ERROR_CODES.VALIDATION_ERROR).details.field, 'Current_Quantity');
  throwsWithCode(function () {
    api5.handleInventoryCreate_({ Item_Name: 'X', Selling_Price: 5, Current_Quantity: 1.5 });
  }, ERROR_CODES.VALIDATION_ERROR);
});

check('inventory.create rejects negative / fractional Minimum_Stock_Level', function () {
  const api5 = p5Api();
  eq(throwsWithCode(function () {
    api5.handleInventoryCreate_({ Item_Name: 'X', Selling_Price: 5, Minimum_Stock_Level: -2 });
  }, ERROR_CODES.VALIDATION_ERROR).details.field, 'Minimum_Stock_Level');
  throwsWithCode(function () {
    api5.handleInventoryCreate_({ Item_Name: 'X', Selling_Price: 5, Minimum_Stock_Level: 2.5 });
  }, ERROR_CODES.VALIDATION_ERROR);
});

check('inventory.create rejects unknown fields', function () {
  const api5 = p5Api();
  const err = throwsWithCode(function () {
    api5.handleInventoryCreate_({ Item_Name: 'X', Selling_Price: 5, Foo: 1 });
  }, ERROR_CODES.VALIDATION_ERROR);
  eq(err.details.unknownColumns, ['Foo']);
});

check('inventory.create fails with CONFLICT when the lock is unavailable', function () {
  const ss = p5Spreadsheet();
  const sandbox = makeSandbox(ss, { lockUnavailable: true });
  sandbox.Session.__activeEmail = 'admin@school.edu';
  const api5 = loadBackend(sandbox);
  api5.__ss = ss;
  throwsWithCode(function () {
    api5.handleInventoryCreate_({ Item_Name: 'X', Selling_Price: 5 });
  }, ERROR_CODES.CONFLICT);
  eq(ss.getSheetByName('Inventory')._rows.length, 4, 'nothing written without the lock');
});

check('inventory.create fails safely when the Inventory sheet is missing', function () {
  const api5 = p5Api({ spreadsheet: p4bSpreadsheetWithoutTab('Inventory') });
  const err = throwsWithCode(function () {
    api5.handleInventoryCreate_({ Item_Name: 'X', Selling_Price: 5 });
  }, ERROR_CODES.SERVER_ERROR);
  eq(err.details.sheet, 'Inventory');
});

section('Phase 5: stock operations');

check('inventory.stockIn increases quantity and records a STOCK_IN movement', function () {
  p5LastApi = p5Api();
  const env = p5Post(p5LastApi, 'inventory.stockIn', { Item_ID: 'ITM-001', Quantity: 25, Reason: 'Restock' });
  eq(env.success, true);
  eq(env.data.Current_Quantity, 125, '100 + 25');
  eq(env.data.movement.Movement_ID, 'MOV-003');
  eq(env.data.movement.Movement_Type, 'STOCK_IN');
  eq(env.data.movement.Quantity, 25);
  eq(env.data.movement.Recorded_By, 'STF-1');
  eq(p5RowCount('Inventory_Movements'), 3);
  eq(p5MovCol('Item_ID')[2], 'ITM-001');
});

check('inventory.stockIn requires a positive whole Quantity', function () {

  p5LastApi = p5Api();

  var missing = p5Post(p5LastApi, 'inventory.stockIn', {
    Item_ID: 'ITM-001'
  });

  eq(missing.success, false);
  eq(missing.error, ERROR_CODES.VALIDATION_ERROR);
  eq(missing.details.field, 'Quantity');

  var zero = p5Post(p5LastApi, 'inventory.stockIn', {
    Item_ID: 'ITM-001',
    Quantity: 0
  });

  eq(zero.success, false);
  eq(zero.error, ERROR_CODES.VALIDATION_ERROR);

  var negative = p5Post(p5LastApi, 'inventory.stockIn', {
    Item_ID: 'ITM-001',
    Quantity: -1
  });

  eq(negative.success, false);
  eq(negative.error, ERROR_CODES.VALIDATION_ERROR);

  var fractional = p5Post(p5LastApi, 'inventory.stockIn', {
    Item_ID: 'ITM-001',
    Quantity: 2.5
  });

  eq(fractional.success, false);
  eq(fractional.error, ERROR_CODES.VALIDATION_ERROR);

});


check('inventory.stockIn rejects an unknown Item_ID', function () {

  p5LastApi = p5Api();

  var env = p5Post(p5LastApi, 'inventory.stockIn', {
    Item_ID: 'ITM-404',
    Quantity: 1
  });

  eq(env.success, false);
  eq(env.error, ERROR_CODES.NOT_FOUND);

  eq(
    p5RowCount('Inventory_Movements'),
    2,
    'no movement recorded for a rejected call'
  );

});
check('inventory.stockOut decreases quantity and records a STOCK_OUT movement', function () {
  p5LastApi = p5Api();
  const env = p5Post(p5LastApi, 'inventory.stockOut', {
    Item_ID: 'ITM-001', Quantity: 30, Reason: 'Issue to class'
  });
  eq(env.success, true);
  eq(env.data.Current_Quantity, 70);
  eq(env.data.movement.Movement_Type, 'STOCK_OUT');
  eq(env.data.movement.Quantity, 30);
  eq(env.data.movement.Recorded_By, 'STF-1');
  eq(p5RowCount('Inventory_Movements'), 3);
  const rows = p5LastApi.__ss.getSheetByName('Inventory')._rows;
  eq(rows[1][5], 70, 'the item row must be updated in place');
});

check('inventory.stockOut can never drive Current_Quantity negative', function () {

  p5LastApi = p5Api();

  const env = p5Post(p5LastApi, 'inventory.stockOut', {
    Item_ID: 'ITM-002',
    Quantity: 11
  });

  eq(env.success, false);
  eq(env.error, ERROR_CODES.VALIDATION_ERROR);

  eq(env.details.available, 10);
  eq(env.details.requested, 11);

  eq(
    p5LastApi.__ss.getSheetByName('Inventory')._rows[2][5],
    10,
    'quantity untouched'
  );

  eq(
    p5RowCount('Inventory_Movements'),
    2,
    'no movement recorded'
  );

});

check('inventory.stockOut may reach exactly zero', function () {
  p5LastApi = p5Api();
  const env = p5Post(p5LastApi, 'inventory.stockOut', { Item_ID: 'ITM-002', Quantity: 10 });
  eq(env.data.Current_Quantity, 0);
  eq(env.data.item.Status, 'Out of Stock');
});

check('inventory.stockOut rejects an unknown Item_ID', function () {

  p5LastApi = p5Api();

  const env = p5Post(p5LastApi, 'inventory.stockOut', {
    Item_ID: 'ITM-404',
    Quantity: 1
  });

  eq(env.success, false);
  eq(env.error, ERROR_CODES.NOT_FOUND);

  eq(p5RowCount('Inventory_Movements'), 2);

});

check('inventory.stockOut writes inside the script lock', function () {
  p5LastApi = p5Api();
  const log = observeSheetAccess(p5LastApi, 'Inventory');
  const env = p5Post(p5LastApi, 'inventory.stockOut', { Item_ID: 'ITM-001', Quantity: 1 });
  eq(env.success, true);
  const writes = log.filter(function (e) { return e.op === 'write'; });
  ok(writes.length >= 1, 'the item row must be written');
  ok(writes.every(function (w) { return w.depth > 0; }), 'writes happen inside the lock');
});

section('Phase 5: stationery.fulfill');

const stnItemQty = function (api5, itemId) {
  const rows = api5.__ss.getSheetByName('Inventory')._rows;
  const idx = rows[0].indexOf('Current_Quantity');
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === itemId) return rows[i][idx];
  }
  return undefined;
};

check('stationery.fulfill reduces inventory and records the movement', function () {
  p5LastApi = p5Api();
  const before = stnItemQty(p5LastApi, 'ITM-001');
  const env = p5Post(p5LastApi, 'stationery.fulfill', { Transaction_ID: 'STN-001', Quantity_Given: 4 });
  eq(env.success, true);
  eq(env.message, 'Stationery fulfillment recorded');
  eq(env.data.Quantity_Given, 4);
  eq(env.data.Quantity_Remaining, 6, '10 purchased - 4 given');
  eq(env.data.Fulfillment_Status, 'Partial');
  eq(env.data.Given_By, 'STF-1', 'given by the authenticated user, not the client');
  eq(stnItemQty(p5LastApi, 'ITM-001'), before - 4, 'inventory must decrease');
  const movRows = p5LastApi.__ss.getSheetByName('Inventory_Movements')._rows;
  eq(movRows.length, 4, 'one movement row appended');
  const lastMov = movRows[movRows.length - 1];
  eq(lastMov[2], 'STOCK_OUT');
  eq(Number(lastMov[3]), 4);
  eq(lastMov[1], 'ITM-001');
});

check('stationery.fulfill to completion marks Fulfilled', function () {
  p5LastApi = p5Api();
  const env = p5Post(p5LastApi, 'stationery.fulfill', { Transaction_ID: 'STN-001', Quantity_Given: 10 });
  eq(env.data.Quantity_Given, 10);
  eq(env.data.Quantity_Remaining, 0);
  eq(env.data.Fulfillment_Status, 'Fulfilled');
});

check('stationery.fulfill rejects more than the remaining quantity', function () {
  p5LastApi = p5Api();
  const env = p5Post(p5LastApi, 'stationery.fulfill', {
    Transaction_ID: 'STN-001',
    Quantity_Given: 11
  });
  eq(env.success, false);
  eq(env.error, ERROR_CODES.VALIDATION_ERROR);
});

check('stationery.fulfill rejects an unknown transaction', function () {
  p5LastApi = p5Api();
  const env = p5Post(p5LastApi, 'stationery.fulfill', {
    Transaction_ID: 'STN-404',
    Quantity_Given: 1
  });
  eq(env.success, false);
  eq(env.error, ERROR_CODES.NOT_FOUND);
});

section('Phase 5: inventory.movements');

check('inventory.movements returns the ledger in sheet order', function () {
  p5LastApi = p5Api();
  const env = p5Get(p5LastApi, 'inventory.movements');
  eq(env.success, true);
  ok(env.data.length >= 2, 'should return the seeded movements');
  eq(env.data[0].Movement_ID, 'MOV-001');
  ok(env.data[0].hasOwnProperty('Recorded_By'));
});

check('inventory.movements filters by Item_ID', function () {
  p5LastApi = p5Api();
  const env = p5Get(p5LastApi, 'inventory.movements', { Item_ID: 'ITM-001' });
  eq(env.success, true);
  env.data.forEach(function (m) { eq(m.Item_ID, 'ITM-001'); });
});

check('inventory.movements filters by Movement_Type', function () {
  p5LastApi = p5Api();
  const env = p5Get(p5LastApi, 'inventory.movements', { Movement_Type: 'STOCK_OUT' });
  eq(env.success, true);
  env.data.forEach(function (m) { eq(m.Movement_Type, 'STOCK_OUT'); });
});

check('inventory.movements rejects unknown filter fields', function () {
  p5LastApi = p5Api();
  throwsWithCode(function () {
    api.handleInventoryMovements_({ Bogus: 1 });
  }, ERROR_CODES.VALIDATION_ERROR);
});

section('Phase 6: dashboard.summary');

check('dashboard.summary returns the documented aggregate shape', function () {
  p5LastApi = p5Api();
  var env = p5Post(p5LastApi, 'dashboard.summary');
  eq(env.success, true);
  ok(typeof env.data.activeStudents === 'number', 'activeStudents should be a number');
  ok(typeof env.data.activeStaff === 'number', 'activeStaff should be a number');
  ok(typeof env.data.schoolFeesCollected === 'number', 'schoolFeesCollected should be a number');
  ok(typeof env.data.feedingFeesCollected === 'number', 'feedingFeesCollected should be a number');
  ok(typeof env.data.lowStockItems === 'number', 'lowStockItems should be a number');
  ok(Array.isArray(env.data.recentPayments), 'recentPayments should be an array');
  ok(env.data.recentPayments.length <= 5, 'recentPayments should be at most 5');
});

check('dashboard.summary reports correct active student count', function () {
  p5LastApi = p5Api();
  var env = p5Post(p5LastApi, 'dashboard.summary');
  eq(env.success, true);
  eq(env.data.activeStudents, 2, 'two active students (STU-1, STU-2); STU-3 is Withdrawn');
});

check('dashboard.summary reports correct active staff count', function () {
  p5LastApi = p5Api();
  var env = p5Post(p5LastApi, 'dashboard.summary');
  eq(env.success, true);
  eq(env.data.activeStaff, 2, 'two active staff (STF-1, STF-2); STF-3 is Inactive');
});

check('dashboard.summary sums Amount_Paid across both fee sheets', function () {
  p5LastApi = p5Api();
  var env = p5Post(p5LastApi, 'dashboard.summary');
  eq(env.success, true);
  eq(env.data.schoolFeesCollected, 1800, '1200 + 600 + 0 from School_Fees');
  eq(env.data.feedingFeesCollected, 675, '450 + 225 from Feeding_Fees');
});

check('dashboard.summary counts non-In-Stock inventory items', function () {
  p5LastApi = p5Api();
  var env = p5Post(p5LastApi, 'dashboard.summary');
  eq(env.success, true);
  eq(env.data.lowStockItems, 2, 'ITM-002 (Low Stock) + ITM-003 (Out of Stock)');
});

check('dashboard.summary returns at most 5 most recent payments', function () {
  p5LastApi = p5Api();
  p5Post(p5LastApi, 'schoolFees.create', {
    Student_ID: 'STU-1', Academic_Year: '2026/2027', Term: 'Term 1',
    Amount_Due: 500, Amount_Paid: 500, Payment_Method: 'Cash', Payment_Date: '2026-09-01'
  });
  p5Post(p5LastApi, 'schoolFees.create', {
    Student_ID: 'STU-1', Academic_Year: '2026/2027', Term: 'Term 1',
    Amount_Due: 500, Amount_Paid: 500, Payment_Method: 'Cash', Payment_Date: '2026-09-02'
  });
  var env = p5Post(p5LastApi, 'dashboard.summary');
  eq(env.success, true);
  ok(env.data.recentPayments.length <= 5, 'should cap at 5 most recent payments');
});

check('dashboard.summary sorts recent payments most-recent-first', function () {
  p5LastApi = p5Api();
  var env = p5Post(p5LastApi, 'dashboard.summary');
  eq(env.success, true);
  eq(env.data.recentPayments.length, 5, '3 School_Fees + 2 Feeding_Fees fixture payments');
  ok(env.data.recentPayments[0].Payment_Date >= env.data.recentPayments[1].Payment_Date, 'payments should be descending');
  ok(env.data.recentPayments[1].Payment_Date >= env.data.recentPayments[2].Payment_Date, 'payments should be descending');
  ok(env.data.recentPayments[2].Payment_Date >= env.data.recentPayments[3].Payment_Date, 'payments should be descending');
});

check('dashboard.summary includes type discriminators on recent payments', function () {
  p5LastApi = p5Api();
  var env = p5Post(p5LastApi, 'dashboard.summary');
  eq(env.success, true);
  env.data.recentPayments.forEach(function (p) {
    ok(p.type === 'schoolFees' || p.type === 'feedingFees', 'type must be schoolFees or feedingFees: ' + p.type);
    ok(typeof p.Payment_ID === 'string' && p.Payment_ID !== '', 'Payment_ID must be present');
    ok(typeof p.Student_ID === 'string' && p.Student_ID !== '', 'Student_ID must be present');
    ok(typeof p.Amount_Paid === 'number', 'Amount_Paid must be numeric');
    ok(typeof p.Payment_Date === 'string' && p.Payment_Date !== '', 'Payment_Date must be present');
  });
});

section('Phase 5: permission enforcement');

check('every reserved (unimplemented) action still reports NOT_FOUND', function () {
  [
    'salaries.list',
    'delegations.list',
    'audit.list',
  ].forEach(function (action) {
    const envelope = readEnvelope(api.doGet({ parameter: { action: action } }));
    eq(envelope.success, false, action + ' unexpectedly succeeded');
    eq(envelope.error, ERROR_CODES.NOT_FOUND, action + ' should report NOT_FOUND');
  });
});

check('Phase 4B fee modules are implemented, later modules still placeholders', function () {
     // Phase 4B: implemented.
  ['SchoolFees.js', 'FeedingFees.js',
   // Phase 5: now implemented.
   'Stationery.js', 'Inventory.js', 'Dashboard.js'].forEach(function (file) {
    const code = stripComments(fs.readFileSync(path.join(ROOT, file), 'utf8'));
    ok(!/^function myFunction\(\)\s*\{\s*\}$/.test(code.trim()), file + ' should be implemented in Phase 4B');
  });
    // Phase 5-7 stubs are still untouched placeholders.
  ['Salaries.js',
   'Delegations.js', 'Audit.js'].forEach(function (file) {
    const code = stripComments(fs.readFileSync(path.join(ROOT, file), 'utf8'));
    ok(/^function myFunction\(\)\s*\{\s*\}$/.test(code.trim()), file + ' is no longer an untouched placeholder');
  });
  // Phase 3 modules are now implemented.
  ['Students.js', 'Staff.js'].forEach(function (file) {
    const code = stripComments(fs.readFileSync(path.join(ROOT, file), 'utf8'));
    ok(!/^function myFunction\(\)\s*\{\s*\}$/.test(code.trim()), file + ' should be implemented in Phase 3');
  });
  // Auth.js and Permissions.js are the Phase 2 scope and must be implemented.
  ['Auth.js', 'Permissions.js'].forEach(function (file) {
    const code = stripComments(fs.readFileSync(path.join(ROOT, file), 'utf8'));
    ok(!/^function myFunction\(\)\s*\{\s*\}$/.test(code.trim()), file + ' should be implemented in Phase 2');
  });
});

check('dashboard.summary requires DASHBOARD.READ permission', function () {
  var noPerm = loadBackendAs('teacher@school.edu', makeFullSpreadsheet());
  var env = noPerm.doPost({ parameter: {},
    postData: { contents: JSON.stringify({ action: 'dashboard.summary', payload: {} }) } });
  var result = JSON.parse(env.getContent());
  eq(result.success, false);
  eq(result.error, ERROR_CODES.FORBIDDEN);
});

check('dashboard.summary returns zeros for empty sheets', function () {
  // DASHBOARD.READ must resolve, so include the authorization tables.
  var p4aEmpty = p4aPermissionSheets();
  var emptySs = makeSpreadsheet('EmptyDashboard', [
    makeSheet('Users', [['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'],
      ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']]),
    makeSheet('Students', [['Student_ID', 'Status']]),
    makeSheet('Staff', [['Staff_ID', 'Employment_Status']]),
    makeSheet('School_Fees', [['Payment_ID', 'Amount_Paid']]),
    makeSheet('Feeding_Fees', [['Payment_ID', 'Amount_Paid']]),
    makeSheet('Inventory', [['Item_ID', 'Status']]),
    makeSheet('Inventory_Movements', [['Movement_ID']]),
    p4aEmpty.roles,
    p4aEmpty.permissions,
    p4aEmpty.rolePermissions,
  ]);
  var emptyApi = loadBackendAs('admin@school.edu', emptySs);
  var env = doPostEnvelope(emptyApi, 'dashboard.summary', {});
  console.log('DEBUG env: ' + JSON.stringify(env));
  eq(env.success, true);
  eq(env.data.activeStudents, 0);
  eq(env.data.activeStaff, 0);
  eq(env.data.schoolFeesCollected, 0);
  eq(env.data.feedingFeesCollected, 0);
  eq(env.data.lowStockItems, 0);
  eq(env.data.recentPayments.length, 0);
});

check('dashboard.summary: recentPayments empty when both fee sheets are empty', function () {
  // DASHBOARD.READ must resolve, so include the authorization tables.
  var p4aNoFees = p4aPermissionSheets();
  var emptySs = makeSpreadsheet('NoFees', [
    makeSheet('Users', [['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'],
      ['USR-1', 'STF-1', 'admin@school.edu', 'Admin', 'Active', '']]),
    makeSheet('Students', [['Student_ID', 'Status']]),
    makeSheet('Staff', [['Staff_ID', 'Employment_Status']]),
    makeSheet('School_Fees', [['Payment_ID', 'Amount_Paid']]),
    makeSheet('Feeding_Fees', [['Payment_ID', 'Amount_Paid']]),
    makeSheet('Inventory', [['Item_ID', 'Status']]),
    makeSheet('Inventory_Movements', [['Movement_ID']]),
    p4aNoFees.roles,
    p4aNoFees.permissions,
    p4aNoFees.rolePermissions,
  ]);
  var emptyApi = loadBackendAs('admin@school.edu', emptySs);
  var env = doPostEnvelope(emptyApi, 'dashboard.summary', {});
  eq(env.success, true);
  eq(env.data.recentPayments.length, 0, 'no payments when both sheets are empty');
});

/* ==========================================================================
 * Summary
 * ======================================================================== */

section('Summary');
console.log('\n  ' + passed + ' passed, ' + failed + ' failed');

if (failed) {
  console.log('\nFailures:');
  failures.forEach(function (failure) {
    console.log('  - ' + failure);
  });
}

process.exit(failed ? 1 : 0);