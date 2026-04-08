/**
 * NO~BULL BOOKS -- INITIALIZER
 *
 * Provisions all required sheets on a blank client spreadsheet.
 * Safe to run multiple times -- only creates sheets that are missing.
 * Called automatically on first boot via api('runInitialSetup').
 */
function checkAndInitSheet(params) {
  try {
    var ss      = getDb(params);
    var created = [];

    // -- Full schema -- must match column indices in Config.gs -----------------
    var SCHEMA = {
      'Settings': [
        'CompanyName','CompanyAddress','CompanyPostcode','CompanyPhone','CompanyEmail',
        'vatRegNumber','InvoicePrefix','NextInvoiceNumber','BillPrefix','NextBillNumber',
        'LogoURL','BankName','AccountName','SortCode','AccountNumber',
        'FinancialYearStart','FinancialYearEnd','CurrentFinancialYear','VATRegistered',
        'VATScheme','VATRate','VATFrequency','MTDEnabled',
        'hmrcClientID','hmrcClientSecret','hmrcAccessToken','hmrcTestMode','hmrcTokenExpiry',
        'hmrcNINO','mtdBusinessId','cnPrefix','nextCNNumber','poPrefix','nextPONumber',
        'lockedBefore','emailSubject','emailBody','paymentTerms','invoiceFooter',
        'templateAccentColor','templateLogoPosition','templateShowReference','templateFont',
        'baseCurrency','enabledCurrencies','businessStartDate','yearEndDay',
        'ownerEmail'
      ],
      'Clients': [
        'ClientId','Name','Email','Phone','Address','Postcode','Country',
        'VATNumber','ContactName','Notes','CreatedDate','Active'
      ],
      'Suppliers': [
        'SupplierId','Name','Email','Phone','Address','Postcode','Country',
        'VATNumber','ContactName','Notes','CreatedDate','Active'
      ],
      'Invoices': [
        'InvoiceId','InvoiceNumber','ClientId','ClientName','ClientEmail','ClientAddress',
        'IssueDate','DueDate','Subtotal','VATRate','VAT','Total',
        'AmountPaid','AmountDue','Status','PaymentDate','Notes','PDFURL',
        'Currency','ExchangeRate','BaseTotal',
        'BankAccount','VoidDate','VoidReason','VoidedBy'
      ],
      'InvoiceLines': [
        'LineId','InvoiceId','Description','Quantity','UnitPrice','VATRate','LineTotal'
      ],
      'Bills': [
        'BillId','BillNumber','SupplierId','SupplierName','IssueDate','DueDate',
        'Subtotal','VATRate','VAT','Total','AmountPaid','AmountDue',
        'Status','PaymentDate','Notes','Reconciled','VoidDate','VoidReason','VoidedBy',
        'Currency','ExchangeRate'
      ],
      'BillLines': [
        'LineId','BillId','Description','Quantity','UnitPrice','VATRate','LineTotal'
      ],
      'BankAccounts': [
        'AccountId','AccountName','BankName','SortCode','AccountNumber',
        'Currency','OpeningBalance','CurrentBalance','Active','Notes'
      ],
      'BankTransactions': [
        'TxId','Date','Description','Reference','Amount','Type','Category',
        'BankAccount','Status','ReconciledDate','MatchId','MatchType','Notes'
      ],
      'Transactions': [
        'TxId','Date','Description','Reference','Debit','Credit',
        'AccountCode','DocumentId','DocumentType','CreatedBy'
      ],
      'ChartOfAccounts': [
        'AccountCode','AccountName','AccountType','SubType','Description','Active','OpeningBalance'
      ],
      'CreditNotes': [
        'CNId','CNNumber','InvoiceId','ClientId','ClientName','IssueDate',
        'Subtotal','VAT','Total','Status','Reason','AppliedDate','AppliedInvoiceId'
      ],
      'CreditNoteLines': [
        'LineId','CNId','Description','Quantity','UnitPrice','VATRate','LineTotal'
      ],
      'PurchaseOrders': [
        'POId','PONumber','SupplierId','SupplierName','IssueDate','ExpectedDelivery',
        'Subtotal','VAT','Total','Status','Notes','ApprovedBy','BillId'
      ],
      'PurchaseOrderLines': [
        'LineId','POId','Description','Quantity','UnitPrice','VATRate','LineTotal'
      ],
      'BadDebts': [
        'BadDebtId','InvoiceId','InvoiceNumber','ClientId','ClientName',
        'WriteOffDate','AmountWrittenOff','VATElement','VATReclaimStatus',
        'VATClaimDate','Reason','WrittenOffBy'
      ],
      'VATReturns': [
        'ReturnId','PeriodStart','PeriodEnd','Box1','Box2','Box3','Box4',
        'Box5','Box6','Box7','Box8','Box9','Status','SubmittedDate','Reference'
      ],
      'VATObligations': [
        'ObligationId','VRN','PeriodKey','Start','End','Due','Status','Received'
      ],
      'Users': [
        'Email','Role','AddedBy','AddedDate','Active','Notes'
      ],
      'InvoiceHistory': [
        'HistoryId','InvoiceId','Action','Detail','ChangedBy','ChangedDate'
      ],
      'BillHistory': [
        'HistoryId','BillId','Action','Detail','ChangedBy','ChangedDate'
      ],
      'InvoiceFiles': [
        'FileId','InvoiceId','FileName','FileURL','FileType',
        'UploadedBy','UploadedDate','Description'
      ],
      'BillFiles': [
        'FileId','BillId','FileName','FileURL','FileType',
        'UploadedBy','UploadedDate','Description'
      ],
      'PaymentAllocations': [
        'AllocationId','InvoiceId','Amount','PaymentDate','BankAccountId','Notes','CreatedBy'
      ],
      'RecurringInvoices': [
        'RecurringId','ClientId','ClientName','Frequency','NextDate',
        'InvoicePrefix','Lines','Status','CreatedBy','CreatedDate'
      ],
      'FXRatesLog': [
        'LogId','Date','BaseCurrency','TargetCurrency','Rate','Source'
      ],
      'FixedAssets': [
        'AssetId','Name','Category','PurchaseDate','Cost','DepreciationMethod',
        'UsefulLifeYears','ResidualValue','AccumulatedDepreciation','NetBookValue','Status','Notes'
      ],
      'DepreciationRuns': [
        'RunId','PeriodEndDate','PeriodMonths','AssetsProcessed','TotalDepreciation',
        'PostedToLedger','RunDate','RunBy'
      ],
      'FinancialYears': [
        'YearId','YearLabel','StartDate','EndDate','Status','ClosedDate','ClosedBy'
      ],
      'BankRules': [
        'RuleId','Pattern','Category','AccountCode','Direction','Priority','Active'
      ],
      'ClientCredits': [
        'CreditId','ClientId','Amount','Source','SourceId','AppliedDate','Notes'
      ],
      'ITSAObligations': [
        'ObligationId','NINO','TaxYear','Quarter','Start','End','Due','Status'
      ],
      'ITSASubmissions': [
        'SubmissionId','NINO','TaxYear','Quarter','SubmittedDate','Reference','Status','Detail'
      ]
    };

    for (var sheetName in SCHEMA) {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        var headers     = SCHEMA[sheetName];
        var neededCols  = headers.length;
        var currentCols = sheet.getMaxColumns();
        if (currentCols < neededCols) {
          sheet.insertColumnsAfter(currentCols, neededCols - currentCols);
        }
        sheet.getRange(1, 1, 1, neededCols).setValues([headers]);
        sheet.setFrozenRows(1);
        created.push(sheetName);
      }
    }

    // -- Seed Users sheet with the caller as Owner (bootstrap only) ------------
    var usersSheet = ss.getSheetByName('Users');
    if (usersSheet && usersSheet.getLastRow() < 2) {
      // Prefer the email passed from SetupService (real client email)
      // over Session.getActiveUser() which always returns edward in hub model
      var ownerEmail = (params && params._ownerEmail) ? params._ownerEmail : '';
      if (!ownerEmail) {
        try { ownerEmail = Session.getActiveUser().getEmail(); } catch(e) {}
      }
      if (ownerEmail) {
        usersSheet.appendRow([ownerEmail, 'Owner', 'system', new Date(), true, 'Initial setup']);
        Logger.log('Initializer: seeded Owner -- ' + ownerEmail);

        // -- Store ownerEmail in Settings so Auth.gs can identify the user ------
        // Since the hub runs as edward (USER_DEPLOYING), we cannot use
        // Session.getActiveUser() to identify clients. We store their email
        // in the Settings sheet so every API call can resolve identity.
        var settingsSheet = ss.getSheetByName('Settings');
        if (settingsSheet && settingsSheet.getLastRow() >= 1) {
          var headers = settingsSheet.getRange(1, 1, 1, settingsSheet.getLastColumn()).getValues()[0];
          var ownerEmailCol = headers.indexOf('ownerEmail');
          if (ownerEmailCol >= 0) {
            // Ensure row 2 exists
            if (settingsSheet.getLastRow() < 2) {
              settingsSheet.appendRow([]);
            }
            settingsSheet.getRange(2, ownerEmailCol + 1).setValue(ownerEmail);
            Logger.log('Initializer: stored ownerEmail in Settings -- ' + ownerEmail);
          }
        }
      }
    }

    Logger.log('checkAndInitSheet: created ' + created.length + ' sheets: ' + created.join(', '));
    return {
      success: true,
      created: created,
      message: created.length
        ? 'Provisioned ' + created.length + ' sheets.'
        : 'All sheets already exist -- no changes made.'
    };

  } catch(e) {
    Logger.log('checkAndInitSheet ERROR: ' + e.toString());
    return { success: false, message: 'Setup failed: ' + e.toString() };
  }
}

// -- Dev helper -- run from Apps Script editor to test against a blank sheet ---
function debug_testInit() {
  var testParams = { _sheetId: '11G7eOUSefvUCR95evsx0A_pauub5LwaE10DNtcium8M' };
  var result = checkAndInitSheet(testParams);
  Logger.log(JSON.stringify(result));
}