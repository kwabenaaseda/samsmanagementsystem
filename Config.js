/**
 * Config.js
 * Centralized configuration and constants for the School Management System API.
 *
 * SPREADSHEET ACCESS
 * This script is container-bound to the spreadsheet below. However a web app
 * request (doGet/doPost) executes OUTSIDE the container context, where
 * SpreadsheetApp.getActiveSpreadsheet() returns null and .getId() throws.
 * The spreadsheet is therefore ALWAYS addressed by explicit ID, via
 * Utils.js -> getSpreadsheet_() -> SpreadsheetApp.openById(CONFIG.SHEET_ID).
 *
 * STORED-VALUE CONVENTION
 * Values persisted to sheets use Title Case ('Active', 'Paid', 'Voided').
 * Machine-only tokens (audit actions, error codes) use UPPER_SNAKE_CASE.
 */

const CONFIG = {
  /** Service name reported by the `health` action. */
  SERVICE_NAME: 'School Management System API',

  /**
   * Spreadsheet ID of the container spreadsheet this script is bound to.
   * Do NOT replace this with getActiveSpreadsheet(). See note above.
   */
  SHEET_ID: '17nqU2fim3e9txZaLLyo4DsOo46yuJUCJRV9ku144S6I',

  /** Milliseconds to wait for the script lock before failing a write. */
  LOCK_TIMEOUT_MS: 30000,

  /** Logical sheet (tab) names. Must match the real tab names exactly. */
  SHEETS: {
    STUDENTS: 'Students',
    STAFF: 'Staff',
    USERS: 'Users',
    ROLES: 'Roles',
    PERMISSIONS: 'Permissions',
    SCHOOL_FEES: 'School_Fees',
    FEEDING_FEES: 'Feeding_Fees',
    STATIONERY: 'Stationery',
    INVENTORY: 'Inventory',
    INVENTORY_MOVEMENTS: 'Inventory_Movements',
    SALARY_PAYMENTS: 'Salary_Payments',
    DELEGATIONS: 'Delegations',
    AUDIT_LOG: 'Audit_Log'
  },

  /**
   * Student status. 'Active' and 'Withdrawn' are the values V1 uses;
   * 'Graduated' and 'Suspended' are reserved so the stored vocabulary does not
   * have to change later. Students are never physically deleted -- withdrawal
   * preserves history.
   */
  STUDENT_STATUS: {
    ACTIVE: 'Active',
    WITHDRAWN: 'Withdrawn',
    GRADUATED: 'Graduated',
    SUSPENDED: 'Suspended'
  },

  /**
   * Payment methods used by School_Fees, Feeding_Fees and Stationery.
   * ASSUMPTION: this list still needs confirming with the school.
   */
  PAYMENT_METHOD: {
    CASH: 'Cash',
    MOBILE_MONEY: 'Mobile Money',
    BANK_TRANSFER: 'Bank Transfer',
    CHEQUE: 'Cheque',
    OTHER: 'Other'
  },

  /**
   * Payment status. 'Voided' implements the agreed correction mechanism:
   * financial rows are never deleted, they are voided with a reason so
   * history is preserved.
   */
  PAYMENT_STATUS: {
    UNPAID: 'Unpaid',
    PARTIAL: 'Partial',
    PAID: 'Paid',
    VOIDED: 'Voided'
  },

  /** Academic terms used by School_Fees and Feeding_Fees. */
  TERM: {
    TERM_1: 'Term 1',
    TERM_2: 'Term 2',
    TERM_3: 'Term 3'
  },
  /**
   * Module names used by the Permissions sheet (Module column) and the
   * Audit_Log sheet (Module column). Stored values must match these exactly.
   */
  MODULE: {
    STUDENTS: 'Students',
    STAFF: 'Staff',
    USERS: 'Users',
    ROLES: 'Roles',
    PERMISSIONS: 'Permissions',
    SCHOOL_FEES: 'School_Fees',
    FEEDING_FEES: 'Feeding_Fees',
    STATIONERY: 'Stationery',
    INVENTORY: 'Inventory',
    SALARIES: 'Salaries',
    DELEGATIONS: 'Delegations',
    AUDIT: 'Audit_Log',
    DASHBOARD: 'Dashboard',
    SYSTEM: 'System'
  },
  /**
   * API action names -- the frontend/backend contract.
   * Transcribed from the agreed endpoint list in the project brief.
   *
   * IMPLEMENTATION STATUS: Phase 1 routes ONLY `HEALTH`.
   * Every other action name below is a RESERVED identifier for a later phase
   * and is NOT implemented -- the router will return NOT_FOUND for it.
   */
  ACTIONS: {
    HEALTH: 'health',

    STUDENTS: {
      LIST: 'students.list',
      GET: 'students.get',
      CREATE: 'students.create',
      UPDATE: 'students.update',
      WITHDRAW: 'students.withdraw'
    },
    SCHOOL_FEES: {
      LIST: 'schoolFees.list',
      CREATE: 'schoolFees.create',
      BY_STUDENT: 'schoolFees.byStudent'
    },
    FEEDING_FEES: {
      LIST: 'feedingFees.list',
      CREATE: 'feedingFees.create'
    },
    STATIONERY: {
      LIST: 'stationery.list',
      CREATE: 'stationery.create',
      FULFILL: 'stationery.fulfill'
    },
    INVENTORY: {
      LIST: 'inventory.list',
      CREATE: 'inventory.create',
      STOCK_IN: 'inventory.stockIn',
      STOCK_OUT: 'inventory.stockOut',
      MOVEMENTS: 'inventory.movements'
    },
    STAFF: {
      LIST: 'staff.list',
      CREATE: 'staff.create',
      UPDATE: 'staff.update'
    },
    SALARIES: {
      LIST: 'salaries.list',
      CREATE: 'salaries.create'
    },
    DELEGATIONS: {
      LIST: 'delegations.list',
      CREATE: 'delegations.create',
      REVOKE: 'delegations.revoke'
    },
    AUDIT: {
      LIST: 'audit.list',
      GET: 'audit.get'
    },
    DASHBOARD: {
      SUMMARY: 'dashboard.summary'
    }
  }
};

/**
 * Derived allow-lists used by enum validation (see assertOneOf_ in Utils.js).
 * Derived from the enums above rather than duplicated, so the two can never
 * drift apart.
 */
CONFIG.VALUES = {
  STUDENT_STATUS: Object.values(CONFIG.STUDENT_STATUS),
  PAYMENT_METHOD: Object.values(CONFIG.PAYMENT_METHOD),
  PAYMENT_STATUS: Object.values(CONFIG.PAYMENT_STATUS),
  TERM: Object.values(CONFIG.TERM),
  MODULE: Object.values(CONFIG.MODULE)
};