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
| `Config.js` | `CONFIG` object: sheet ID + `CONFIG.SHEETS` name map | complete |
| `Router.js` | `doGet(e)` / `doPost(e)` entry points | stub-level (echoes body) |
| `Response.js` | `success()`, `failure()`, `jsonResponse()` envelopes | complete |
| `Auth.js` | login / session / credential checks | **empty stub** |
| `Permissions.js` | role + permission resolution | **empty stub** |
| `Audit.js` | append-only audit logging | **empty stub** |
| `Students.js` | student CRUD | **empty stub** |
| `Staff.js` | staff CRUD | **empty stub** |
| `SchoolFees.js` | school fee billing/payments | **empty stub** |
| `FeedingFees.js` | feeding fee billing/payments | **empty stub** |
| `Stationery.js` | stationery sales | **empty stub** |
| `Inventory.js` | stock items + movements | **empty stub** |
| `Salaries.js` | salary payments | **empty stub** |
| `Delegations.js` | delegated permissions | **empty stub** |
| `Dashboard.js` | aggregate/reporting endpoints | **empty stub** |
| `Utils.js` | shared helpers | **empty stub** |

"Empty stub" means the file currently contains only:

```js
function myFunction() {

}
```

That is the Apps Script default placeholder, so these modules still need to be
implemented.

## Sheet tabs referenced by `Config.js`

`Students`, `Staff`, `Users`, `Roles`, `Permissions`, `School_Fees`,
`Feeding_Fees`, `Stationery`, `Inventory`, `Inventory_Movements`,
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

## Known issues to resolve

1. **`Config.js` evaluates at load time.** `SpreadsheetApp.getActiveSpreadsheet().getId()`
   is executed at global scope on every script run. When the script is *not*
   bound to a spreadsheet (web app request, time-driven trigger, standalone
   execution) `getActiveSpreadsheet()` returns `null` and `.getId()` throws
   `Cannot read properties of null (reading 'getId')`. Convert to a lazy
   accessor or use `SpreadsheetApp.openById(...)`.
2. **`doPost` is an echo, not a router.** It returns the parsed request body
   with no action dispatch, validation, auth, or per-module routing.
3. **Web app is not publicly reachable.** `access: MYSELF` + `executeAs:
   USER_DEPLOYING` means only the owner can invoke the deployment.
4. **13 module files are unimplemented stubs** (see table above).