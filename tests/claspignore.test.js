/**
 * tests/claspignore.test.js
 *
 * Verifies Task 1: the future React frontend and its node_modules can never be
 * pushed to Apps Script.
 *
 * HOW IT WORKS
 * `clasp push` selects files with:
 *     micromatch.not(files, ignorePatterns, { dot: true })
 * (see @google/clasp build/src/core/files.js in the installed version).
 *
 * This test reads the REAL .claspignore from the repository root, applies that
 * exact library call to a simulated repository containing the real backend
 * files plus a realistic Vite + React + TypeScript frontend, and asserts that
 * only backend files survive.
 *
 * It also demonstrates the counterfactual: with no .claspignore, clasp's
 * built-in DEFAULT_CLASP_IGNORE lets frontend files through, which is why the
 * file exists.
 *
 * Requires the locally installed clasp (for its bundled micromatch).
 * Run with:  node tests/claspignore.test.js
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const CLASPIGNORE = path.join(ROOT, '.claspignore');

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

/* ==========================================================================
 * Locate the micromatch that the installed clasp actually uses
 * ======================================================================== */

function resolveMicromatch() {
  const home = process.env.HOME || '';
  const claspDirs = [
    path.join(home, '.bun/install/global/node_modules/@google/clasp'),
    path.join(home, '.nvm/versions/node', process.version, 'lib/node_modules/@google/clasp'),
    path.join(home, '.bun/install/global/node_modules/@google/clasp/build'),
  ];

  for (let i = 0; i < claspDirs.length; i++) {
    try {
      return require(require.resolve('micromatch', { paths: [claspDirs[i]] }));
    } catch (err) {
      /* try the next candidate */
    }
  }
  try {
    return require('micromatch');
  } catch (err) {
    return null;
  }
}

const micromatch = resolveMicromatch();

/* ==========================================================================
 * Parse the real .claspignore the way clasp does
 * ======================================================================== */

/** Strip comments and blank lines, exactly as an ignore file is interpreted. */
function parseIgnoreFile(filePath) {
  return fs
    .readFileSync(filePath, 'utf8')
    .split('\n')
    .map(function (line) {
      return line.trim();
    })
    .filter(function (line) {
      return line !== '' && line.charAt(0) !== '#';
    });
}

/** The patterns clasp uses when there is NO .claspignore (v3.4.1). */
const DEFAULT_CLASP_IGNORE = [
  '**/**',
  '!**/appsscript.json',
  '!**/*.gs',
  '!**/*.js',
  '!**/*.ts',
  '!**/*.html',
  '.git/**',
  'node_modules/**',
];

/** Every .js file actually sitting in the repository root (the backend). */
function realBackendFiles() {
  return fs
    .readdirSync(ROOT, { withFileTypes: true })
    .filter(function (entry) {
      return entry.isFile() && /\.[jt]s$/.test(entry.name);
    })
    .map(function (entry) {
      return entry.name;
    });
}

/** A realistic tree: real backend + a stock Vite React TS frontend. */
function simulatedTree() {
  return realBackendFiles()
    .concat([
      'appsscript.json',
      '.clasp.json',
      '.gitignore',
      '.claspignore',
      'README.md',
      'package.json',
      // Vite scaffold files that really ship in a frontend directory
      'frontend/package.json',
      'frontend/index.html',
      'frontend/vite.config.ts',
      'frontend/eslint.config.js',
      'frontend/postcss.config.js',
      'frontend/tailwind.config.js',
      'frontend/src/main.tsx',
      'frontend/src/App.tsx',
      'frontend/src/services/api.ts',
      'frontend/src/utils/format.js',
      // dependency trees
      'node_modules/react/index.js',
      'frontend/node_modules/react/index.js',
      'frontend/node_modules/.vite/deps/chunk-ABC.js',
      'frontend/node_modules/@vitejs/plugin-react/dist/index.js',
      // other non-backend directories
      'tests/backend.test.js',
      'docs/design.md',
      '.git/config',
    ])
    .sort();
}

/* ==========================================================================
 * Tests
 * ======================================================================== */

console.log('claspignore verification (Task 1)\n=================================');

if (!micromatch) {
  console.log('\n  SKIPPED: could not locate micromatch from the installed clasp.');
  console.log('  This test needs clasp 3.x installed locally (it borrows clasp\'s own');
  console.log('  matcher so the verification uses identical logic).');
  console.log('  Nothing was verified. Install clasp, or run `clasp status` manually.');
  process.exit(2);
}

const patterns = parseIgnoreFile(CLASPIGNORE);
const tree = simulatedTree();
const kept = micromatch.not(tree, patterns, { dot: true });
const expectedKept = realBackendFiles().concat(['appsscript.json']).sort();

section('The .claspignore file itself');

check('.claspignore exists and declares a usable pattern set', function () {
  ok(fs.existsSync(CLASPIGNORE), '.claspignore is missing');
  ok(patterns.length >= 4, 'only ' + patterns.length + ' patterns parsed, expected more');
  ok(patterns.indexOf('**/**') !== -1, 'the ignore-everything pattern is missing');
  ok(patterns.indexOf('!appsscript.json') !== -1, 'appsscript.json is not re-included');
  ok(patterns.indexOf('!*.js') !== -1, 'root .js files are not re-included');
});

section('What clasp push would upload (using the real .claspignore)');

check('exactly the backend files and the manifest are uploaded', function () {
  eq(kept, expectedKept, 'the uploaded set is not exactly the Apps Script backend');
});

check('the uploaded set is non-empty and contains the real modules', function () {
  ok(kept.length >= 5, 'suspiciously few files would be pushed: ' + kept.length);
  ok(kept.indexOf('Config.js') !== -1, 'Config.js would not be pushed');
  ok(kept.indexOf('Router.js') !== -1, 'Router.js would not be pushed');
  ok(kept.indexOf('Utils.js') !== -1, 'Utils.js would not be pushed');
  ok(kept.indexOf('appsscript.json') !== -1, 'appsscript.json would not be pushed');
});

section('The frontend guarantee');

check('NO frontend file would ever be pushed', function () {
  const leaked = kept.filter(function (file) {
    return file.indexOf('frontend/') === 0;
  });
  eq(leaked, [], 'frontend files leaked into the push set: ' + JSON.stringify(leaked));
});

check('NO node_modules file would ever be pushed, at any depth', function () {
  const leaked = kept.filter(function (file) {
    return file.indexOf('node_modules/') !== -1;
  });
  eq(leaked, [], 'dependency files leaked into the push set: ' + JSON.stringify(leaked));
});

check('the nested frontend/node_modules case is covered (clasp default is not)', function () {
  const targets = ['frontend/node_modules/react/index.js', 'frontend/node_modules/.vite/deps/chunk-ABC.js'];
  targets.forEach(function (file) {
    ok(kept.indexOf(file) === -1, file + ' would be pushed');

    // Prove clasp's own default would have let it through.
    const underDefault = micromatch.not([file], DEFAULT_CLASP_IGNORE, { dot: true });
    eq(underDefault, [file], 'expected the clasp default to leak ' + file + ', so the test is not meaningful');
  });
});

check('frontend .js and .html files are excluded even though clasp allows that extension', function () {
  ['frontend/eslint.config.js', 'frontend/tailwind.config.js', 'frontend/index.html'].forEach(function (file) {
    ok(kept.indexOf(file) === -1, file + ' would be pushed');

    const underDefault = micromatch.not([file], DEFAULT_CLASP_IGNORE, { dot: true });
    ok(underDefault.length === 1, 'expected the clasp default to leak ' + file);
  });
});

section('Repository files that must not reach Apps Script');

check('docs, config, and test files are excluded', function () {
  ['README.md', '.gitignore', '.clasp.json', 'package.json', 'tests/backend.test.js', 'docs/design.md'].forEach(
    function (file) {
      ok(kept.indexOf(file) === -1, file + ' would be pushed');
    }
  );
});

check('the .git directory is excluded', function () {
  ok(kept.indexOf('.git/config') === -1, '.git contents would be pushed');
});

/* ==========================================================================
 * Summary
 * ======================================================================== */

section('Summary');
console.log('\n  ' + passed + ' passed, ' + failed + ' failed');
console.log('  (uploaded: ' + kept.length + ' files -> ' + kept.join(', ') + ')');

if (failed) {
  console.log('\nFailures:');
  failures.forEach(function (failure) {
    console.log('  - ' + failure);
  });
}

process.exit(failed ? 1 : 0);