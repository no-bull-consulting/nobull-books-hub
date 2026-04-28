/**
 * NO~BULL BOOKS — DEMO INSTANCE
 * ─────────────────────────────────────────────────────────────────────────────
 * Populates a sheet with realistic demo data for a generic small business.
 * Run _seedDemoData() once to set up, then install the nightly reset trigger.
 *
 * Demo company: Brightwell Trading Ltd
 * Sheet ID: set DEMO_SHEET_ID below before running
 */

var DEMO_SHEET_ID = 'YOUR_DEMO_SHEET_ID'; // ← replace after creating demo instance

// ─────────────────────────────────────────────────────────────────────────────
// SEED
// ─────────────────────────────────────────────────────────────────────────────

function _seedDemoData() {
  var ss = SpreadsheetApp.openById(DEMO_SHEET_ID);
  Logger.log('Seeding demo data for: ' + ss.getName());

  _clearDemoSheets(ss);
  _seedSettings(ss);
  _seedCOA(ss);
  _seedClients(ss);
  _seedSuppliers(ss);
  _seedBankAccounts(ss);
  _seedInvoices(ss);
  _seedBills(ss);
  _seedTransactions(ss);

  Logger.log('✅ Demo data seeded successfully');
}

function _clearDemoSheets(ss) {
  var sheets = ['Settings','Clients','Suppliers','Invoices','InvoiceLines',
                'Bills','BillLines','BankAccounts','BankTransactions',
                'Transactions','ChartOfAccounts','CreditNotes','PurchaseOrders',
                'BadDebts','AuditLog'];
  sheets.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) {
      sh.deleteRows(2, sh.getLastRow() - 1);
    }
  });
  Logger.log('Sheets cleared');
}

function _seedSettings(ss) {
  var sh = ss.getSheetByName('Settings');
  if (!sh) return;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var row = headers.map(function(h) { return ''; });
  var set = function(name, val) {
    var i = headers.indexOf(name);
    if (i >= 0) row[i] = val;
  };

  set('CompanyName',         'Brightwell Trading Ltd');
  set('CompanyAddress',      '14 Merchant Square, Bristol');
  set('CompanyPostcode',     'BS1 4RJ');
  set('CompanyPhone',        '0117 456 7890');
  set('CompanyEmail',        'accounts@brightwelltrading.co.uk');
  set('vatRegNumber',        'GB 123 4567 89');
  set('InvoicePrefix',       'INV-');
  set('NextInvoiceNumber',   5);
  set('BillPrefix',          'BILL-');
  set('NextBillNumber',      4);
  set('VATRegistered',       true);
  set('VATScheme',           'standard');
  set('VATRate',             20);
  set('VATFrequency',        'quarterly');
  set('FinancialYearStart',  '2025-04-01');
  set('FinancialYearEnd',    '2026-03-31');
  set('CurrentFinancialYear','2025/26');
  set('baseCurrency',        'GBP');
  set('paymentTerms',        30);
  set('templateAccentColor', '#1a3c6b');
  set('ownerEmail',          'demo@nobull.consulting');

  sh.getRange(2, 1, 1, row.length).setValues([row]);
  Logger.log('Settings seeded');
}

function _seedCOA(ss) {
  var sh = ss.getSheetByName('ChartOfAccounts');
  if (!sh) return;
  var accounts = [
    ['1000','Business Current Account','Asset','Bank Accounts',0,14820.50,true,''],
    ['1100','Trade Debtors','Asset','Current Assets',0,3600,true,''],
    ['1200','Stock / Inventory','Asset','Current Assets',0,0,true,''],
    ['2100','Trade Creditors','Liability','Current Liabilities',0,1840,true,''],
    ['2200','VAT Control','Liability','Current Liabilities',0,0,true,''],
    ['3000','Share Capital','Equity','Capital',10000,10000,true,''],
    ['3100','Retained Earnings','Equity','Retained Earnings',0,0,true,''],
    ['4000','Sales — Goods','Revenue','Sales Revenue',0,18000,true,''],
    ['4100','Sales — Services','Revenue','Service Revenue',0,0,true,''],
    ['5000','Cost of Goods Sold','Expense','Cost of Sales',0,4200,true,''],
    ['6000','Rent & Rates','Expense','Operating Expenses',0,1800,true,''],
    ['6100','Utilities','Expense','Operating Expenses',0,420,true,''],
    ['6200','Office Supplies','Expense','Operating Expenses',0,180,true,''],
    ['6300','Professional Fees','Expense','Operating Expenses',0,960,true,''],
    ['6400','Marketing & Advertising','Expense','Operating Expenses',0,480,true,''],
    ['6500','Bank Charges','Expense','Operating Expenses',0,60,true,''],
    ['7000','Wages & Salaries','Expense','Payroll',0,0,true,''],
  ];
  accounts.forEach(function(a) { sh.appendRow(a); });
  Logger.log('COA seeded — ' + accounts.length + ' accounts');
}

function _seedClients(ss) {
  var sh = ss.getSheetByName('Clients');
  if (!sh) return;
  var clients = [
    ['CLI_DEMO_001','Apex Solutions Ltd','finance@apexsolutions.co.uk','020 7123 4567','22 City Road, London','EC1V 2PY','UK','GB987654321','Sarah Mitchell','Key account',new Date('2024-01-15'),true],
    ['CLI_DEMO_002','Green Valley Retailers','orders@greenvalley.co.uk','0161 234 5678','8 Market Street, Manchester','M1 1PW','UK','','James Cooper','',new Date('2024-03-01'),true],
    ['CLI_DEMO_003','Harlow & Partners','accounts@harlowpartners.co.uk','0113 345 6789','5 Park Lane, Leeds','LS1 2AB','UK','GB112233445','Emma Harlow','',new Date('2024-06-10'),true],
  ];
  clients.forEach(function(c) { sh.appendRow(c); });
  Logger.log('Clients seeded — ' + clients.length);
}

function _seedSuppliers(ss) {
  var sh = ss.getSheetByName('Suppliers');
  if (!sh) return;
  var suppliers = [
    ['SUP_DEMO_001','Meridian Supplies Ltd','invoices@meridiansupplies.co.uk','01234 567890','Industrial Estate, Birmingham','B6 5RJ','UK','GB445566778','Tom Richards','',new Date('2024-01-01'),true],
    ['SUP_DEMO_002','FastTrack Logistics','billing@fasttracklogistics.co.uk','0800 123 456','Warehouse Way, Coventry','CV1 2GH','UK','','','',new Date('2024-02-01'),true],
    ['SUP_DEMO_003','Digital Edge Marketing','hello@digitaledge.co.uk','020 8765 4321','Creative Quarter, London','SE1 7PB','UK','GB998877665','Anna Brooks','',new Date('2024-04-01'),true],
  ];
  suppliers.forEach(function(s) { sh.appendRow(s); });
  Logger.log('Suppliers seeded — ' + suppliers.length);
}

function _seedBankAccounts(ss) {
  var sh = ss.getSheetByName('BankAccounts');
  if (!sh) return;
  sh.appendRow([
    'BA_DEMO_001',
    'Business Current Account',
    'Barclays',
    'Current',
    '20-45-67',
    '12345678',
    10000,
    14820.50,
    '2026-03-15',
    true,
    '1000'
  ]);
  Logger.log('Bank account seeded');
}

function _seedInvoices(ss) {
  var invSh  = ss.getSheetByName('Invoices');
  var lineSh = ss.getSheetByName('InvoiceLines');
  if (!invSh || !lineSh) return;

  var invoices = [
    // [InvoiceId, Number, ClientId, ClientName, ClientEmail, ClientAddress, IssueDate, DueDate,
    //  Subtotal, VATRate, VAT, Total, AmountPaid, AmountDue, Status, PaymentDate, Notes, PDFURL, Currency, FXRate, BaseTotal]
    ['INV_DEMO_001','INV-0001','CLI_DEMO_001','Apex Solutions Ltd','finance@apexsolutions.co.uk','22 City Road, London, EC1V 2PY',
     new Date('2026-01-10'), new Date('2026-02-09'),
     3000,20,600,3600,3600,0,'Paid',new Date('2026-02-05'),'Q4 goods supply','','GBP',1,3600],

    ['INV_DEMO_002','INV-0002','CLI_DEMO_002','Green Valley Retailers','orders@greenvalley.co.uk','8 Market Street, Manchester, M1 1PW',
     new Date('2026-02-01'), new Date('2026-03-03'),
     2500,20,500,3000,3000,0,'Paid',new Date('2026-02-28'),'February stock order','','GBP',1,3000],

    ['INV_DEMO_003','INV-0003','CLI_DEMO_001','Apex Solutions Ltd','finance@apexsolutions.co.uk','22 City Road, London, EC1V 2PY',
     new Date('2026-03-01'), new Date('2026-03-31'),
     2000,20,400,2400,0,2400,'Sent','',' March delivery — payment pending','','GBP',1,2400],

    ['INV_DEMO_004','INV-0004','CLI_DEMO_003','Harlow & Partners','accounts@harlowpartners.co.uk','5 Park Lane, Leeds, LS1 2AB',
     new Date('2026-03-15'), new Date('2026-04-14'),
     1000,20,200,1200,0,1200,'Approved','','Consultancy retainer March','','GBP',1,1200],
  ];

  invoices.forEach(function(inv) { invSh.appendRow(inv); });

  var lines = [
    ['LN_DEMO_001','INV_DEMO_001','Office furniture supply',1,3000,20,'4000'],
    ['LN_DEMO_002','INV_DEMO_002','Retail stock — mixed goods',1,2500,20,'4000'],
    ['LN_DEMO_003','INV_DEMO_003','Product supply — March',1,2000,20,'4000'],
    ['LN_DEMO_004','INV_DEMO_004','Consultancy services',1,1000,20,'4100'],
  ];
  lines.forEach(function(l) { lineSh.appendRow(l); });
  Logger.log('Invoices seeded — ' + invoices.length);
}

function _seedBills(ss) {
  var billSh = ss.getSheetByName('Bills');
  var lineSh = ss.getSheetByName('BillLines');
  if (!billSh || !lineSh) return;

  var bills = [
    ['BILL_DEMO_001','BILL-0001','SUP_DEMO_001','Meridian Supplies Ltd',
     new Date('2026-01-05'), new Date('2026-02-04'),
     3500,20,700,4200,4200,0,'Paid',new Date('2026-01-30'),'Stock purchase Jan',false,'','',''],

    ['BILL_DEMO_002','BILL-0002','SUP_DEMO_002','FastTrack Logistics',
     new Date('2026-02-10'), new Date('2026-03-12'),
     350,0,0,350,350,0,'Paid',new Date('2026-03-05'),'Delivery charges Feb',false,'','',''],

    ['BILL_DEMO_003','BILL-0003','SUP_DEMO_003','Digital Edge Marketing',
     new Date('2026-03-01'), new Date('2026-03-31'),
     800,20,160,960,0,960,'Pending','','Q1 marketing campaign',false,'','',''],
  ];
  bills.forEach(function(b) { billSh.appendRow(b); });

  var lines = [
    ['BL_DEMO_001','BILL_DEMO_001','Stock purchase — mixed goods',1,3500,20],
    ['BL_DEMO_002','BILL_DEMO_002','Delivery charges',1,350,0],
    ['BL_DEMO_003','BILL_DEMO_003','Marketing campaign Q1',1,800,20],
  ];
  lines.forEach(function(l) { lineSh.appendRow(l); });
  Logger.log('Bills seeded — ' + bills.length);
}

function _seedTransactions(ss) {
  var sh = ss.getSheetByName('Transactions');
  var bsh = ss.getSheetByName('BankTransactions');
  if (!sh) return;

  var txns = [
    // Invoice payments received
    ['TXN_DEMO_001',new Date('2026-02-05'),'Payment','INV-0001','1000','1100',3600,'Payment — Apex Solutions Ltd','INV_DEMO_001','',true],
    ['TXN_DEMO_002',new Date('2026-02-28'),'Payment','INV-0002','1000','1100',3000,'Payment — Green Valley Retailers','INV_DEMO_002','',true],
    // Invoice journal entries
    ['TXN_DEMO_003',new Date('2026-01-10'),'Invoice','INV-0001','1100','4000',3000,'Invoice — Apex Solutions Ltd','INV_DEMO_001','',false],
    ['TXN_DEMO_004',new Date('2026-02-01'),'Invoice','INV-0002','1100','4000',2500,'Invoice — Green Valley Retailers','INV_DEMO_002','',false],
    ['TXN_DEMO_005',new Date('2026-03-01'),'Invoice','INV-0003','1100','4000',2000,'Invoice — Apex Solutions Ltd','INV_DEMO_003','',false],
    ['TXN_DEMO_006',new Date('2026-03-15'),'Invoice','INV-0004','1100','4100',1000,'Invoice — Harlow & Partners','INV_DEMO_004','',false],
    // Bill payments made
    ['TXN_DEMO_007',new Date('2026-01-30'),'Payment','BILL-0001','2100','1000',4200,'Payment — Meridian Supplies','','BILL_DEMO_001',true],
    ['TXN_DEMO_008',new Date('2026-03-05'),'Payment','BILL-0002','2100','1000',350,'Payment — FastTrack Logistics','','BILL_DEMO_002',true],
    // Bill journal entries
    ['TXN_DEMO_009',new Date('2026-01-05'),'Bill','BILL-0001','5000','2100',3500,'Stock purchase — Meridian','','BILL_DEMO_001',false],
    ['TXN_DEMO_010',new Date('2026-02-10'),'Bill','BILL-0002','6000','2100',350,'Delivery — FastTrack','','BILL_DEMO_002',false],
    ['TXN_DEMO_011',new Date('2026-03-01'),'Bill','BILL-0003','6400','2100',800,'Marketing — Digital Edge','','BILL_DEMO_003',false],
    // Other expenses
    ['TXN_DEMO_012',new Date('2026-01-01'),'Expense','RENT-JAN','6000','1000',600,'Office rent January','','',true],
    ['TXN_DEMO_013',new Date('2026-02-01'),'Expense','RENT-FEB','6000','1000',600,'Office rent February','','',true],
    ['TXN_DEMO_014',new Date('2026-03-01'),'Expense','RENT-MAR','6000','1000',600,'Office rent March','','',true],
    ['TXN_DEMO_015',new Date('2026-01-15'),'Expense','UTILS-JAN','6100','1000',140,'Utilities January','','',true],
    ['TXN_DEMO_016',new Date('2026-02-15'),'Expense','UTILS-FEB','6100','1000',140,'Utilities February','','',true],
    ['TXN_DEMO_017',new Date('2026-03-15'),'Expense','UTILS-MAR','6100','1000',140,'Utilities March','','',true],
  ];
  txns.forEach(function(t) { sh.appendRow(t); });

  if (bsh) {
    var bankTxns = [
      ['BTX_DEMO_001',new Date('2026-02-05'),'Payment received for invoice INV-0001','INV-0001',3600,'Credit','Sales','BA_DEMO_001','Reconciled','','','',''],
      ['BTX_DEMO_002',new Date('2026-02-28'),'Payment received for invoice INV-0002','INV-0002',3000,'Credit','Sales','BA_DEMO_001','Reconciled','','','',''],
      ['BTX_DEMO_003',new Date('2026-01-30'),'Payment for bill BILL-0001','BILL-0001',-4200,'Debit','Expenses','BA_DEMO_001','Reconciled','','','',''],
      ['BTX_DEMO_004',new Date('2026-03-05'),'Payment for bill BILL-0002','BILL-0002',-350,'Debit','Expenses','BA_DEMO_001','Reconciled','','','',''],
      ['BTX_DEMO_005',new Date('2026-01-01'),'Office rent January','RENT-JAN',-600,'Debit','Expenses','BA_DEMO_001','Reconciled','','','',''],
      ['BTX_DEMO_006',new Date('2026-02-01'),'Office rent February','RENT-FEB',-600,'Debit','Expenses','BA_DEMO_001','Reconciled','','','',''],
      ['BTX_DEMO_007',new Date('2026-03-01'),'Office rent March','RENT-MAR',-600,'Debit','Expenses','BA_DEMO_001','Unreconciled','','','',''],
      ['BTX_DEMO_008',new Date('2026-01-15'),'Utilities January','UTILS-JAN',-140,'Debit','Expenses','BA_DEMO_001','Reconciled','','','',''],
      ['BTX_DEMO_009',new Date('2026-02-15'),'Utilities February','UTILS-FEB',-140,'Debit','Expenses','BA_DEMO_001','Reconciled','','','',''],
      ['BTX_DEMO_010',new Date('2026-03-15'),'Utilities March','UTILS-MAR',-140,'Debit','Expenses','BA_DEMO_001','Unreconciled','','','',''],
    ];
    bankTxns.forEach(function(t) { bsh.appendRow(t); });
  }

  Logger.log('Transactions seeded — ' + txns.length + ' journal, ' + (bsh ? '10' : '0') + ' bank');
}

// ─────────────────────────────────────────────────────────────────────────────
// NIGHTLY RESET TRIGGER
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Install the nightly reset trigger.
 * Run once from GAS editor.
 */
function _installDemoResetTrigger() {
  // Remove existing demo triggers first
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === '_nightlyDemoReset') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('_nightlyDemoReset')
    .timeBased()
    .atHour(3)       // 3am
    .everyDays(1)
    .inTimezone('Europe/London')
    .create();

  Logger.log('✅ Nightly demo reset trigger installed — runs at 3am London time');
}

/**
 * Nightly reset — called by trigger.
 */
function _nightlyDemoReset() {
  Logger.log('🔄 Nightly demo reset starting...');
  _seedDemoData();
  Logger.log('✅ Nightly demo reset complete');
}

/**
 * Remove the nightly trigger (call if demo is taken down).
 */
function _removeDemoResetTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === '_nightlyDemoReset') {
      ScriptApp.deleteTrigger(t);
      Logger.log('Trigger removed');
    }
  });
}