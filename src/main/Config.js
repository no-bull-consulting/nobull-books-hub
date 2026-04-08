/**
 * NO~BULL BOOKS -- CONFIG
 * Global constants: spreadsheet ID, sheet names, roles & permissions
 * -------------------------------------------------------------
 */

// ============================================
// SPREADSHEET
// ============================================

// Fallback only -- each client request carries _sheetId in params.
var DEFAULT_SPREADSHEET_ID = '1V71QyGO6IFvU8_JBlc7FVybYq3Z80I3c0nd5gVqsw0M';

/**
 * _ss(params)
 * Smart spreadsheet selector -- used by functions that don't receive params
 * directly (e.g. legacy module-level helpers).
 * Prefer getDb(params) wherever params is available.
 */
function _ss(params) {
  var id = (params && params._sheetId) ? params._sheetId : null;

  if (id) {
    try {
      var ss = SpreadsheetApp.openById(id);
      if (ss) return ss;
    } catch(e) {
      Logger.log('_ss: could not open sheet by ID ' + id + ': ' + e);
    }
  }

  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch(e) {
    // Last resort -- hub default (dev/debug only)
    return SpreadsheetApp.openById(DEFAULT_SPREADSHEET_ID);
  }
}

// ============================================
// SHEET NAMES
// ============================================
var SHEETS = {
  SETTINGS:             'Settings',
  CHART_OF_ACCOUNTS:    'ChartOfAccounts',
  INVOICES:             'Invoices',
  INVOICE_LINES:        'InvoiceLines',
  CLIENTS:              'Clients',
  SUPPLIERS:            'Suppliers',
  BILLS:                'Bills',
  BILL_LINES:           'BillLines',
  BANK_ACCOUNTS:        'BankAccounts',
  BANK_TRANSACTIONS:    'BankTransactions',
  VAT_RETURNS:          'VATReturns',
  VAT_OBLIGATIONS:      'VATObligations',
  TRANSACTIONS:         'Transactions',
  INVOICE_HISTORY:      'InvoiceHistory',
  BILL_HISTORY:         'BillHistory',
  INVOICE_FILES:        'InvoiceFiles',
  BILL_FILES:           'BillFiles',
  PAYMENT_ALLOCATIONS:  'PaymentAllocations',
  BANK_RULES:           'BankRules',
  ITSA_OBLIGATIONS:     'ITSAObligations',
  ITSA_SUBMISSIONS:     'ITSASubmissions',
  USERS:                'Users',
  BAD_DEBTS:            'BadDebts',
  CREDIT_NOTES:         'CreditNotes',
  CREDIT_NOTE_LINES:    'CreditNoteLines',
  PURCHASE_ORDERS:      'PurchaseOrders',
  PURCHASE_ORDER_LINES: 'PurchaseOrderLines',
  CLIENT_CREDITS:       'ClientCredits',
  FIXED_ASSETS:         'FixedAssets',
  DEPRECIATION_RUNS:    'DepreciationRuns',
  FINANCIAL_YEARS:      'FinancialYears',
  RECURRING_INVOICES:   'RecurringInvoices',
  FX_RATES_LOG:         'FXRatesLog'
};

// ============================================
// INVOICE COLUMN INDICES  (1-based for getRange)
// ============================================
var INV_COLS = {
  ID:           1,
  NUMBER:       2,
  CLIENT_ID:    3,
  CLIENT_NAME:  4,
  CLIENT_EMAIL: 5,
  CLIENT_ADDR:  6,
  ISSUE_DATE:   7,
  DUE_DATE:     8,
  SUBTOTAL:     9,
  VAT_RATE:     10,
  VAT:          11,
  TOTAL:        12,
  AMOUNT_PAID:  13,
  AMOUNT_DUE:   14,
  STATUS:       15,
  PAYMENT_DATE: 16,
  NOTES:        17,
  PDF_URL:      18,
  BANK_ACCT:    19,
  VOID_DATE:    20,
  VOID_REASON:  21,
  VOIDED_BY:    22
};

// ============================================
// BILL COLUMN INDICES  (1-based)
// ============================================
var BILL_COLS = {
  ID:            1,
  NUMBER:        2,
  SUPPLIER_ID:   3,
  SUPPLIER_NAME: 4,
  ISSUE_DATE:    5,
  DUE_DATE:      6,
  SUBTOTAL:      7,
  VAT_RATE:      8,
  VAT:           9,
  TOTAL:         10,
  AMOUNT_PAID:   11,
  AMOUNT_DUE:    12,
  STATUS:        13,
  PAYMENT_DATE:  14,
  NOTES:         15,
  RECONCILED:    16,
  VOID_DATE:     17,
  VOID_REASON:   18,
  VOIDED_BY:     19,
  CURRENCY:      20,
  EXCHANGE_RATE: 21
};

// ============================================
// BANK TRANSACTION COLUMN INDICES  (1-based)
// ============================================
var BANK_TX_COLS = {
  ID:              1,
  DATE:            2,
  DESCRIPTION:     3,
  REFERENCE:       4,
  AMOUNT:          5,
  TYPE:            6,
  CATEGORY:        7,
  BANK_ACCOUNT:    8,
  STATUS:          9,
  RECONCILED_DATE: 10,
  MATCH_ID:        11,
  MATCH_TYPE:      12,
  NOTES:           13
};

// ============================================
// NOMINAL ACCOUNT CODES
// ============================================
var ACCOUNTS = {
  BANK:        '1000',
  DEBTORS:     '1100',
  CREDITORS:   '2100',
  VAT_CONTROL: '2200',
  SALES:       '4000',
  BAD_DEBT:    '8100',
  ROUNDING:    '9100'
};

// ============================================
// DOCUMENT STATUS VALUES
// ============================================
var STATUS = {
  INV_DRAFT:    'Draft',
  INV_SENT:     'Sent',
  INV_PARTIAL:  'Partial',
  INV_PAID:     'Paid',
  INV_VOID:     'Void',
  INV_BAD_DEBT: 'Bad Debt',
  BILL_PENDING: 'Pending',
  BILL_PARTIAL: 'Partial',
  BILL_PAID:    'Paid',
  BILL_VOID:    'Void',
  BD_ELIGIBLE:  'Eligible',
  BD_CLAIMED:   'Claimed',
  BD_NOT_ELIG:  'Not Eligible'
};

// ============================================
// ACL / RBAC -- ROLES & PERMISSIONS
// ============================================
var ROLE_PERMISSIONS = {
  'Owner':      ['invoices.*','clients.*','suppliers.*','bills.*','banking.*',
                 'reports.*','coa.*','settings.write',
                 'mtd.*','users.manage','users.view','credentials.manage',
                 'creditnotes.*','purchaseorders.*'],
  'Admin':      ['invoices.*','clients.*','suppliers.*','bills.*','banking.*',
                 'reports.*','coa.*','settings.write',
                 'mtd.*','users.view',
                 'creditnotes.*','purchaseorders.*'],
  'Accountant': ['invoices.*','clients.read','suppliers.read','bills.*',
                 'banking.read','banking.reconcile','reports.*','coa.read',
                 'mtd.submit','mtd.read','users.view',
                 'creditnotes.*','purchaseorders.read'],
  'Staff':      ['invoices.write','invoices.read','clients.write','clients.read',
                 'suppliers.read','bills.write','bills.read','banking.read',
                 'creditnotes.read','purchaseorders.write','purchaseorders.read'],
  'ReadOnly':   ['invoices.read','clients.read','suppliers.read','bills.read',
                 'banking.read','reports.read','coa.read',
                 'creditnotes.read','purchaseorders.read']
};

// -- Instance Registry ---------------------------------------------------------
var REGISTRY_URL = '';

// -- Superuser -----------------------------------------------------------------
// Platform superuser -- edward@nobull.consulting.
// Full system access across ALL deployed instances.
var SUPERUSER_EMAIL = (function() {
  try {
    var v = PropertiesService.getScriptProperties().getProperty('nb_superuser_email');
    return v || 'edward@nobull.consulting';
  } catch(e) { return 'edward@nobull.consulting'; }
})();

/** Call once per instance to set the superuser email in Script Properties */
function setSuperuserEmail(email) {
  PropertiesService.getScriptProperties().setProperty('nb_superuser_email', email);
  Logger.log('Superuser set to: ' + email);
  return { success: true, email: email };
}

// -- Application version -------------------------------------------------------
var APP_VERSION = '2.0.0';

// -- Google Identity Services -- OAuth Client ID --------------------------------
// Used for client-side Sign in with Google (verifies user identity in hub model)
// Created at console.cloud.google.com -> APIs & Services -> Credentials
var OAUTH_CLIENT_ID = '490062327176-b9q29mjh1dtd4cave1ct4c03ri90dkvc.apps.googleusercontent.com';