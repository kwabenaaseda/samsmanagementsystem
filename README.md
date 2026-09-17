# Sam's Manage — School Management System (Google Apps Script backend)

Apps Script Web App backend serving a single `doGet` / `doPost` JSON API.

- **Script ID:** `14ORobhlgni5YFyvx3UeqI08OzySKPOZ_oMuaYm5l5NAoV1MhzSZ-WiIY`
- **Runtime:** V8, timezone `Africa/Abidjan`, exception logging to Stackdriver
- **Web app:** `executeAs: USER_DEPLOYING`, `access: MYSELF`

> Only a **Script ID** (`14ORob...`) can be used with `clasp clone`.
> A **Deployment ID** (`AKfycb...`) cannot — `clasp clone AKfycb...` fails with
> `Invalid script ID`.

## Files

| File | Purpose | Status |
| --- | --- | --- |
| `appsscript.json` | Manifest (runtime, web app access, timezone) | complete |
| `.clasp.json` | clasp target config (`scriptId`, extensions) | complete |
| `Config.js` | Central `CONFIG`: spreadsheet ID, sheet names, enums, action names, permission catalog | Phases 1-4A complete |
| `Router.js` | Action dispatcher (`doGet`/`doPost`), route table, `health` action | Phases 1-3 routed |
| `Response.js` | `success()`, `failure()`, `jsonResponse()` + error-code taxonomy | Phase 1 complete |
| `Auth.js` | login / session / credential checks | Phase 2 complete |
| `Permissions.js` | sheet-driven role→permission resolution + seed | Phase 4A complete |
| `Audit.js` | append-only audit logging | **empty stub** |
| `Students.js` | student list/get/create/update/withdraw | Phase 3 complete |
| `Staff.js` | staff list/get/create/update/deactivate | Phase 3 complete |
| `SchoolFees.js` | school fee billing/payments | **empty stub** |
| `FeedingFees.js` | feeding fee billing/payments | **empty stub** |
| `Stationery.js` | stationery sales | **empty stub** |
| `Inventory.js` | stock items + movements | **empty stub** |
| `Salaries.js` | salary payments | **empty stub** |
| `Delegations.js` | delegated permissions | **empty stub** |
| `Dashboard.js` | aggregate/reporting endpoints | **empty stub** |
| `Utils.js` | Foundation helpers: sheet access, reads/writes, IDs, dates, validation | Phase 1 complete |

"Empty stub" means the file currently contains only:

```js
function myFunction() {

}
```

That is the Apps Script default placeholder. **8 business modules are still
untouched at this level** (SchoolFees, FeedingFees, Stationery, Inventory,
Salaries, Delegations, Dashboard, Audit) and need implementing.

## Implementation status

- **Phase 1** (foundation): complete.
- **Phase 2** (authentication / authorization): complete.
- **Phase 3** (Students + Staff): implemented and covered by the test suite
  (route table, CRUD, soft delete, validation, locking, permission errors).
- **Phase 4A** (Role_Permissions authorization): implemented; grants now come
  from the `Role_Permissions` sheet (see the section below). **The seed has
  not been run on the live spreadsheet yet** — run `setupRolePermissions()`
  from the Apps Script editor right after the next push.
- The 8 remaining modules above are untouched stubs whose actions still
  return `NOT_FOUND`.

## Authorization (Phase 4A): Roles → Role_Permissions → Permissions

Authorization is driven entirely by the spreadsheet. There is deliberately no
in-code "Admin allows everything" rule and no hidden mapping table — if a role
has no active `Role_Permissions` rows, it has no permissions, including Admin.

Flow: `Users.Role` (role name) → `Roles` → `Role_Permissions` → `Permissions`.

### Role_Permissions schema

| Column | Meaning |
| --- | --- |
| `Role_Permission_ID` | server-generated row id (`RP-…`) |
| `Role_ID` | the role being granted (joins `Roles.Role_ID`) |
| `Permission_ID` | the granted permission (joins `Permissions.Permission_ID`) |
| `Status` | `Active` grants; anything else (e.g. `Inactive`) grants nothing |

`Roles` and `Permissions` keep their existing schemas — nothing is renamed.
The reader adapts to their actual columns: the role-name column is detected
as `Role_Name`, `Role` or `Name`; the permission-code column as
`Permission_Name`, `Permission`, `Permission_Code`, `Code` or `Name`. A
missing `Role_ID` / `Permission_ID` column is tolerated (the name itself then
serves as the join key), so no sheet restructuring is required.

### Canonical permission catalog

`CONFIG.PERMISSION_CODES` holds every permission name (34 today). `STUDENTS.*`
and `STAFF.*` are enforced by Phase 3 routes now; `SCHOOL_FEES.*`,
`FEEDING_FEES.*`, `STATIONERY.*`, `INVENTORY.*`, `SALARIES.*`,
`AUDIT_LOG.READ`, `DELEGATIONS.*` and `DASHBOARD.READ` are reserved so the
catalog does not change when later phases land.

### Setup / seed

`setupRolePermissions()` — run once from the Apps Script editor
(Run ▸ setupRolePermissions), NOT an API action. It is idempotent, runs under
the script lock, and:

1. creates the `Role_Permissions` tab with the four columns if absent;
2. adds a `Permissions` row for every catalog code that is missing
   (matched by name);
3. maps the `Admin` role (matched by name, then by `Role_ID`) to every
   catalog code with `Status = Active`;
4. grants NOTHING to any other role — deny-by-default. Granting
   Teacher/Accountant/etc. is a deliberate later decision made by editing
   the `Role_Permissions` sheet.

If `Roles` or `Permissions` is missing or unreadable it fails with
`SERVER_ERROR` instead of guessing. Everything it writes is visible in the
sheet — there are no hidden mappings.

### Fail-safe behavior

- missing `Roles` / `Permissions` / `Role_Permissions` → `SERVER_ERROR`
  (nothing is silently allowed);
- a mapping row whose `Permission_ID` has no `Permissions` row →
  `SERVER_ERROR` (`orphan-mapping`);
- duplicate `(Role_ID, Permission_ID)` rows with the same effective status
  de-duplicate; conflicting statuses (`Active` vs `Inactive`) → `CONFLICT`;
- unauthenticated → `UNAUTHORIZED`; inactive user → `UNAUTHORIZED`;
  authenticated without the permission (or unknown permission) → `FORBIDDEN`.

### Current role assumptions

- `Users.Role` holds a role NAME (e.g. `Admin` for `USR-001`,
  `mr.mensahgibson@gmail.com`). Role lookup matches the name
  case-insensitively, falling back to `Role_ID`.
- Only `Admin` is seeded. No Teacher/Finance roles are invented; if the live
  `Roles` sheet already contains other roles they simply have no grants
  until someone maps them deliberately.
- The live `Roles`/`Permissions` column layouts could not be read from this
  machine (the repo syncs code, not sheet data); the column detection above
  is what makes adapting safe. Verify the tabs after running the seed.

## API shape (action-based, not REST)

An Apps Script web app is one URL driven by `doGet`/`doPost`. It cannot serve
path segments (`GET /students/:id`) or `PUT`/`DELETE` verbs, so REST routes
cannot be expressed literally. The agreed equivalent:

```
GET  ?action=health
POST {"action":"students.create","payload":{...}}
```

Responses are always HTTP 200, so the frontend branches on the application-level
contract: `payload.success`, `payload.message`, `payload.data`, `payload.error`,
`payload.details`.

## Sheet tabs referenced by `Config.js`

`Students`, `Staff`, `Users`, `Roles`, `Permissions`, `Role_Permissions`,
`School_Fees`, `Feeding_Fees`, `Stationery`, `Inventory`, `Inventory_Movements`,
`Salary_Payments`, `Delegations`, `Audit_Log`

Only `Config.js` is asserted here; the tabs themselves are not created by this repo.

## Workflow

```
Cline  ->  local files  ->  Git  ->  GitHub  ->  clasp push  ->  Apps Script  ->  Web App
```

Pull remote edits before pushing, so the Apps Script editor and the local copy
do not diverge:

```bash
npx @google/clasp pull   # or: clasp pull
git add -A && git commit -m "..."
git push
clasp push               # local -> Apps Script
clasp deploy             # publish a new web app version
```

`clasp push` overwrites the remote project. Never edit in the Apps Script web
editor and locally at the same time without pulling first.

## Setup on a new machine

```bash
git clone <repo-url> && cd Sams_manage
clasp login              # writes ~/.clasprc.json (git-ignored)
clasp push               # .clasp.json supplies the scriptId
```

## Running the tests

No test framework is used. There are two zero-dependency Node scripts:

```bash
node tests/backend.test.js      # 161 tests: config, utils, response, router, health, auth, students, staff, role_permissions
node tests/claspignore.test.js  # 9 tests: proves frontend/ can never be pushed
```

`tests/backend.test.js` concatenates the backend files into one script and runs
it in a Node `vm` context with stubs for the Google services, mirroring how Apps
Script flattens files into a single global scope. This gives real feedback
without a framework, network access, or a deployed script.

`tests/claspignore.test.js` reads the real `.claspignore`, applies clasp's own
matcher to a simulated repository containing a full Vite frontend, and asserts
that only backend files would be pushed. It requires clasp installed locally,
because it borrows clasp's bundled `micromatch`.

## Why `.claspignore` exists

Without it, clasp falls back to a built-in default that allows `*.js`, `*.ts`
and `*.html` at **any depth**, and whose `node_modules/**` rule does **not**
cover a nested `frontend/node_modules`. A React frontend in this repository
would therefore have been uploaded into the Apps Script project — including its
dependency tree.

`.claspignore` ignores everything (`**/**`) and re-includes only root-level
backend files, so nothing in `frontend/`, `tests/`, `docs/` or any
`node_modules/` can ever be pushed.

## Known issues / outstanding items

1. **`clasp push` is currently blocked.** The Apps Script API write path returns
   `403 NOT_AUTHORIZED` ("User has not enabled the Apps Script API"). Reads work
   (`projects.get`/`getContent` return 200) but writes do not, so the Phase 1
   code is committed locally but **not yet uploaded**. Fix: enable it at
   <https://script.google.com/home/usersettings> — account owner, browser only.
   No local change can work around this.
2. **No git remote is configured**, so nothing can be pushed to GitHub yet:
   `git remote add origin git@github.com:kwabenaaseda/samsmanagementsystem.git`
3. **The web app is not publicly reachable.** `webapp.access: MYSELF` means only
   the owner can invoke the deployment, so the `health` action cannot yet be
   verified over HTTP from a browser or the frontend. Changing this needs
   explicit approval; the React frontend will require `ANYONE`.
4. **8 business modules are still untouched stubs** (see the table above).
   Every `module.verb` action name in `CONFIG.ACTIONS` is a reserved identifier
   only — the router returns `NOT_FOUND` for all of them.
5. **`CONFIG.PAYMENT_METHOD` values are an assumption** and need confirming with
   the school.
6. **The Role_Permissions seed has not been run on the live spreadsheet yet.**
   After the next `clasp push`, run `setupRolePermissions()` once from the
   Apps Script editor. Until then, permission-checked routes (students.*,
   staff.*) return `SERVER_ERROR` because the `Role_Permissions` sheet is
   missing; `health`, `auth.me` and `auth.check` without a permission
   argument are unaffected. Deployment order: push → run the seed → verify
   with `auth.check` (`permission: "STUDENTS.READ"`).