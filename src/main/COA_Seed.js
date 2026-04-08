/**
 * NO~BULL BOOKS -- UK STANDARD CHART OF ACCOUNTS SEEDER
 *
 * Installs a complete UK HMRC-compliant double-entry COA.
 * Account codes follow the standard UK nominal ledger convention.
 * Safe to run on new instances -- skips accounts that already exist.
 *
 * Account code ranges:
 *   0000-0999  Fixed Assets
 *   1000-1999  Current Assets
 *   2000-2999  Current Liabilities
 *   3000-3999  Equity / Capital
 *   4000-4999  Sales Revenue
 *   5000-5999  Cost of Sales
 *   6000-6999  Staff / Payroll Costs
 *   7000-7999  Operating Expenses
 *   8000-8999  Depreciation / Non-cash
 *   9000-9999  Other / Suspense
 *
 * USAGE: Call seedUKChartOfAccounts(params) from Api.gs or run directly
 * from the Apps Script editor for an existing instance.
 */

var UK_COA = [
  // -- Fixed Assets --------------------------------------------------------
  ['0010', 'Land & Buildings',             'Asset',   'Fixed Assets',          0],
  ['0020', 'Plant & Machinery',            'Asset',   'Fixed Assets',          0],
  ['0030', 'Office Equipment',             'Asset',   'Fixed Assets',          0],
  ['0040', 'Computer Equipment',           'Asset',   'Fixed Assets',          0],
  ['0050', 'Motor Vehicles',              'Asset',   'Fixed Assets',          0],
  ['0060', 'Furniture & Fittings',        'Asset',   'Fixed Assets',          0],
  ['0100', 'Accumulated Depreciation -- Plant',     'Asset', 'Fixed Assets',   0],
  ['0110', 'Accumulated Depreciation -- Equipment', 'Asset', 'Fixed Assets',   0],
  ['0120', 'Accumulated Depreciation -- Vehicles',  'Asset', 'Fixed Assets',   0],

  // -- Current Assets -------------------------------------------------------
  ['1000', 'Business Current Account',    'Asset',   'Bank Accounts',         0],
  ['1010', 'Business Savings Account',    'Asset',   'Bank Accounts',         0],
  ['1020', 'Petty Cash',                  'Asset',   'Bank Accounts',         0],
  ['1100', 'Trade Debtors (Receivables)', 'Asset',   'Current Assets',        0],
  ['1101', 'Other Debtors',               'Asset',   'Current Assets',        0],
  ['1200', 'Stock / Inventory',           'Asset',   'Current Assets',        0],
  ['1210', 'Work in Progress',            'Asset',   'Current Assets',        0],
  ['1300', 'Prepayments',                 'Asset',   'Current Assets',        0],
  ['1310', 'Accrued Income',              'Asset',   'Current Assets',        0],
  ['1400', 'VAT Recoverable',             'Asset',   'Current Assets',        0],
  ['1500', 'PAYE/NI Recoverable',         'Asset',   'Current Assets',        0],
  ['1600', 'Director\'s Loan Account',   'Asset',   'Current Assets',        0],

  // -- Current Liabilities --------------------------------------------------
  ['2100', 'Trade Creditors (Payables)',  'Liability','Current Liabilities',  0],
  ['2101', 'Other Creditors',             'Liability','Current Liabilities',  0],
  ['2200', 'VAT Control Account',         'Liability','Current Liabilities',  0],
  ['2201', 'VAT on Sales (Output)',       'Liability','Current Liabilities',  0],
  ['2202', 'VAT on Purchases (Input)',    'Liability','Current Liabilities',  0],
  ['2210', 'PAYE Payable',                'Liability','Current Liabilities',  0],
  ['2211', 'National Insurance Payable',  'Liability','Current Liabilities',  0],
  ['2220', 'Pension Payable',             'Liability','Current Liabilities',  0],
  ['2300', 'Accruals',                    'Liability','Current Liabilities',  0],
  ['2310', 'Deferred Income',             'Liability','Current Liabilities',  0],
  ['2400', 'Corporation Tax Payable',     'Liability','Current Liabilities',  0],
  ['2410', 'Dividends Payable',           'Liability','Current Liabilities',  0],
  ['2500', 'Bank Loan (Current)',         'Liability','Current Liabilities',  0],
  ['2600', 'Hire Purchase (Current)',     'Liability','Current Liabilities',  0],

  // -- Long-term Liabilities ------------------------------------------------
  ['2700', 'Bank Loan (Long-term)',        'Liability','Long-term Liabilities',0],
  ['2710', 'Hire Purchase (Long-term)',    'Liability','Long-term Liabilities',0],
  ['2800', 'Director\'s Loan (Long-term)','Liability','Long-term Liabilities',0],

  // -- Equity / Capital -----------------------------------------------------
  ['3000', 'Share Capital / Capital Introduced', 'Equity','Capital',          0],
  ['3010', 'Retained Earnings',           'Equity',  'Retained Earnings',     0],
  ['3020', 'Profit & Loss Account',       'Equity',  'Retained Earnings',     0],
  ['3100', 'Drawings',                    'Equity',  'Capital',               0],
  ['3200', 'Dividends Paid',              'Equity',  'Capital',               0],

  // -- Sales Revenue --------------------------------------------------------
  ['4000', 'Sales -- Standard Rated (20%)','Revenue', 'Sales Revenue',        0],
  ['4001', 'Sales -- Reduced Rate (5%)',   'Revenue', 'Sales Revenue',        0],
  ['4002', 'Sales -- Zero Rated',          'Revenue', 'Sales Revenue',        0],
  ['4003', 'Sales -- Exempt',              'Revenue', 'Sales Revenue',        0],
  ['4010', 'Export Sales (Outside UK)',   'Revenue', 'Sales Revenue',        0],
  ['4100', 'Service Income',              'Revenue', 'Service Revenue',      0],
  ['4110', 'Consulting Fees',             'Revenue', 'Service Revenue',      0],
  ['4200', 'Rental Income',               'Revenue', 'Other Revenue',        0],
  ['4300', 'Commission Received',         'Revenue', 'Other Revenue',        0],
  ['4400', 'Interest Received',           'Revenue', 'Other Revenue',        0],
  ['4900', 'Other Income',                'Revenue', 'Other Revenue',        0],

  // -- Cost of Sales --------------------------------------------------------
  ['5000', 'Purchases -- Standard Rated', 'Expense', 'Cost of Sales',         0],
  ['5001', 'Purchases -- Zero Rated',      'Expense', 'Cost of Sales',        0],
  ['5010', 'Import Purchases',            'Expense', 'Cost of Sales',        0],
  ['5100', 'Direct Materials',            'Expense', 'Cost of Sales',        0],
  ['5110', 'Direct Labour',               'Expense', 'Cost of Sales',        0],
  ['5120', 'Subcontractors',              'Expense', 'Cost of Sales',        0],
  ['5200', 'Opening Stock',               'Expense', 'Cost of Sales',        0],
  ['5210', 'Closing Stock',               'Expense', 'Cost of Sales',        0],
  ['5900', 'Other Cost of Sales',         'Expense', 'Cost of Sales',        0],

  // -- Staff Costs ----------------------------------------------------------
  ['6000', 'Gross Wages & Salaries',      'Expense', 'Payroll',              0],
  ['6010', 'Employer\'s NI',             'Expense', 'Payroll',              0],
  ['6020', 'Employer\'s Pension',         'Expense', 'Payroll',              0],
  ['6030', 'Staff Benefits',              'Expense', 'Payroll',              0],
  ['6100', 'Staff Training',              'Expense', 'Payroll',              0],
  ['6200', 'Staff Recruitment',           'Expense', 'Payroll',              0],

  // -- Operating Expenses ---------------------------------------------------
  ['7000', 'Rent & Rates',                'Expense', 'Operating Expenses',   0],
  ['7010', 'Water, Gas & Electricity',    'Expense', 'Operating Expenses',   0],
  ['7020', 'Building Insurance',          'Expense', 'Operating Expenses',   0],
  ['7100', 'Telephone & Internet',        'Expense', 'Operating Expenses',   0],
  ['7110', 'Postage & Courier',           'Expense', 'Operating Expenses',   0],
  ['7120', 'Stationery & Office Supplies','Expense', 'Operating Expenses',   0],
  ['7130', 'Computer Software & Subscriptions','Expense','Operating Expenses',0],
  ['7200', 'Motor Expenses & Fuel',       'Expense', 'Operating Expenses',   0],
  ['7210', 'Travel & Accommodation',      'Expense', 'Operating Expenses',   0],
  ['7220', 'Subsistence',                 'Expense', 'Operating Expenses',   0],
  ['7300', 'Advertising & Marketing',     'Expense', 'Operating Expenses',   0],
  ['7310', 'Website & Digital',           'Expense', 'Operating Expenses',   0],
  ['7400', 'Accountancy Fees',            'Expense', 'Operating Expenses',   0],
  ['7410', 'Legal & Professional Fees',   'Expense', 'Operating Expenses',   0],
  ['7420', 'Bank Charges & Interest',     'Expense', 'Operating Expenses',   0],
  ['7430', 'Merchant & Card Fees',        'Expense', 'Operating Expenses',   0],
  ['7500', 'Repairs & Maintenance',       'Expense', 'Operating Expenses',   0],
  ['7510', 'Cleaning',                    'Expense', 'Operating Expenses',   0],
  ['7600', 'Entertainment (Non-deductible)','Expense','Operating Expenses',  0],
  ['7700', 'Charitable Donations',        'Expense', 'Operating Expenses',   0],
  ['7800', 'Bad Debts Written Off',       'Expense', 'Operating Expenses',   0],
  ['7900', 'Sundry Expenses',             'Expense', 'Operating Expenses',   0],
  ['7910', 'Other Expenses',              'Expense', 'Operating Expenses',   0],

  // -- Depreciation ---------------------------------------------------------
  ['8000', 'Depreciation -- Plant & Machinery','Expense','Operating Expenses', 0],
  ['8010', 'Depreciation -- Office Equipment', 'Expense','Operating Expenses', 0],
  ['8020', 'Depreciation -- Computer Equipment','Expense','Operating Expenses',0],
  ['8030', 'Depreciation -- Motor Vehicles',   'Expense','Operating Expenses', 0],
  ['8040', 'Amortisation',                    'Expense','Operating Expenses', 0],
  ['8100', 'Bad Debt Expense',                'Expense','Operating Expenses', 0],

  // -- Suspense / Control ---------------------------------------------------
  ['9000', 'Suspense Account',            'Asset',   'Other Assets',          0],
  ['9100', 'Rounding Account',            'Expense', 'Other Expenses',        0],
  ['9900', 'Year-End Retained Earnings',  'Equity',  'Retained Earnings',     0]
];

/**
 * seedUKChartOfAccounts(params)
 *
 * Installs the standard UK COA into the client spreadsheet.
 * Skips accounts that already exist (by account code).
 * Returns { success, created, skipped, message }.
 *
 * Call from Api.gs: case 'seedCOA': return seedUKChartOfAccounts(params);
 * Or run directly from Apps Script editor: seedUKChartOfAccounts({ _sheetId: 'YOUR_ID' })
 */
function seedUKChartOfAccounts(params) {
  try {
    var db    = getDb(params || {});
    var sheet = db.getSheetByName(SHEETS.CHART_OF_ACCOUNTS);

    if (!sheet) {
      return { success: false, message: 'ChartOfAccounts sheet not found. Run initial setup first.' };
    }

    // Build set of existing codes
    var existing  = {};
    var lastRow   = sheet.getLastRow();
    if (lastRow > 1) {
      var existData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      existData.forEach(function(r) {
        if (r[0]) existing[r[0].toString().trim()] = true;
      });
    }

    var toAdd   = [];
    var skipped = 0;

    UK_COA.forEach(function(acc) {
      var code = acc[0].toString();
      if (existing[code]) { skipped++; return; }
      toAdd.push([
        code,          // AccountCode
        acc[1],        // AccountName
        acc[2],        // AccountType
        acc[3],        // Category (SubType)
        acc[4] || 0,   // OpeningBalance
        acc[4] || 0,   // CurrentBalance
        true,          // Active
        ''             // Notes
      ]);
    });

    if (toAdd.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, 8).setValues(toAdd);
    }

    // Sort by account code after seeding
    if (sheet.getLastRow() > 2) {
      var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8);
      range.sort(1); // sort by column 1 (account code)
    }

    Logger.log('seedUKChartOfAccounts: ' + toAdd.length + ' added, ' + skipped + ' skipped.');
    return {
      success: true,
      created: toAdd.length,
      skipped: skipped,
      total:   UK_COA.length,
      message: toAdd.length + ' accounts added, ' + skipped + ' already existed.'
    };
  } catch(e) {
    Logger.log('seedUKChartOfAccounts error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Run from Apps Script editor to seed a specific instance.
 * Replace the sheetId with your client spreadsheet ID.
 */
function debug_seedCOA() {
  var result = seedUKChartOfAccounts({ _sheetId: 'YOUR_SPREADSHEET_ID_HERE' });
  Logger.log(JSON.stringify(result));
}