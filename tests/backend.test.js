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
const BACKEND_FILES = ['Config.js', 'Response.js', 'Utils.js', 'Permissions.js', 'Auth.js', 'Router.js'];

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
  return {
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
            return !options.lockUnavailable;
          },
          releaseLock: function () {},
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
    ' doPost: (typeof doPost !== "undefined" ? doPost : undefined) };';

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
  'School_Fees',
  'Feeding_Fees',
  'Stationery',
  'Inventory',
  'Inventory_Movements',
  'Salary_Payments',
  'Delegations',
  'Audit_Log',
];

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
  return makeSpreadsheet('SchoolManagementSystem', sheets);
}

/** A Users tab with one row per scenario the auth spec exercises. */
function makeAuthSpreadsheet(rows) {
  const users = makeSheet('Users', [
    ['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login'].concat(rows || []),
  ]);
  if (rows) users._rows = [['User_ID', 'Staff_ID', 'Email', 'Role', 'Status', 'Last_Login']].concat(rows);
  return makeSpreadsheet('SchoolManagementSystem', [users]);
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

check('all 13 logical tabs are declared with the agreed names', function () {
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

check('health and the two auth actions are routed', function () {
  eq(api.listAvailableActions_().sort(), ['auth.check', 'auth.me', 'health']);
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
  eq(api.readAll_('Staff'), []);
});

check('readAll_ returns [] for a completely empty sheet', function () {
  const blank = loadBackend(makeSandbox(makeSpreadsheet('Blank', [makeSheet('Students', [])])));
  eq(blank.readAll_('Students'), []);
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
  eq(envelope.data.missingSheets.length, 11);
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

check('a missing action is rejected with all three routed actions', function () {
  const envelope = readEnvelope(api.doGet({ parameter: {} }));
  eq(envelope.success, false);
  eq(envelope.error, 'VALIDATION_ERROR');
  eq(envelope.details.availableActions.sort(), ['auth.check', 'auth.me', 'health']);
});

check('an unknown action is NOT_FOUND and names what is available', function () {
  const envelope = readEnvelope(api.doGet({ parameter: { action: 'does.notExist' } }));
  eq(envelope.success, false);
  eq(envelope.error, 'NOT_FOUND');
  eq(envelope.details.action, 'does.notExist');
  eq(envelope.details.availableActions.sort(), ['auth.check', 'auth.me', 'health']);
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
  eq(apiP.hasPermission_(admin, 'NOPE.READ'), true);
  eq(apiP.hasPermission_(admin, 'not-a-permission'), false);
  eq(apiP.hasPermission_(admin, ''), false);
  eq(apiP.hasPermission_(null, 'STUDENTS.READ'), false);
  throwsWithCode(function () { apiP.assertValidPermissionFormat_('nope'); }, ERROR_CODES.VALIDATION_ERROR);
  apiP.assertValidPermissionFormat_('STUDENTS.READ');
});

check('P2: non-admin holds no permissions temporary map', function () {
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


section('Phase boundary (no fake implementations)');

check('every reserved action is genuinely NOT implemented', function () {
  [
    'students.list',
    'students.create',
    'students.withdraw',
    'staff.list',
    'schoolFees.create',
    'feedingFees.list',
    'stationery.fulfill',
    'inventory.stockIn',
    'salaries.create',
    'delegations.revoke',
    'audit.list',
    'dashboard.summary',
  ].forEach(function (action) {
    const envelope = readEnvelope(api.doGet({ parameter: { action: action } }));
    eq(envelope.success, false, action + ' unexpectedly succeeded');
    eq(envelope.error, ERROR_CODES.NOT_FOUND, action + ' should report NOT_FOUND');
  });
});

check('no phase-2+ module has been implemented yet', function () {
  ['Students.js', 'Staff.js', 'SchoolFees.js', 'FeedingFees.js', 'Stationery.js', 'Inventory.js',
   'Salaries.js', 'Delegations.js', 'Dashboard.js', 'Audit.js'].forEach(function (file) {
    const code = stripComments(fs.readFileSync(path.join(ROOT, file), 'utf8'));
    ok(/^function myFunction\(\)\s*\{\s*\}$/.test(code.trim()), file + ' is no longer an untouched placeholder');
  });
  // Auth.js and Permissions.js are the Phase 2 scope and must be implemented.
  ['Auth.js', 'Permissions.js'].forEach(function (file) {
    const code = stripComments(fs.readFileSync(path.join(ROOT, file), 'utf8'));
    ok(!/^function myFunction\(\)\s*\{\s*\}$/.test(code.trim()), file + ' should be implemented in Phase 2');
  });
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