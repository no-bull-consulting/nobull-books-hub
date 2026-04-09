/**
 * NO~BULL BOOKS -- API LAYER
 * ------------------------------------------------------------------------------
 * Single entry point: handleApiCall(action, paramsJson)
 * Called via google.script.run.handleApiCall() from the frontend.
 */

var API_VERSION = '2.0';

/**
 * handleApiCall(action, paramsJson)
 * Single google.script.run entry point.
 */
function handleApiCall(action, paramsJson) {
  try {
    var params = JSON.parse(paramsJson || '{}');
    var ctx    = _getCurrentUserContext(params);

    if (action === 'askGemini') return handleGeminiRequest(params, ctx);

    var result = _route(action, params, ctx);
    return JSON.parse(JSON.stringify(result));
  } catch(e) {
    Logger.log('handleApiCall ERROR [' + action + ']: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * _route -- maps action strings to module functions.
 * params is always passed through so _sheetId reaches every data function.
 */
function _route(action, params, ctx) {
  switch (action) {

    // -- STARTUP --------------------------------------------------------------
    case 'sendOTP':             return sendOTP(params.email, params);
    case 'verifyOTP':           return verifyOTP(params.email, params.otp, params);
    case 'getStartupData': {
      // OTP-verified email overrides the identity from Settings sheet
      if (params._verifiedEmail) {
        var _vEmail = params._verifiedEmail.toString().toLowerCase().trim();
        // Superuser bypass
        if (typeof SUPERUSER_EMAIL !== 'undefined' && _vEmail === SUPERUSER_EMAIL.toLowerCase()) {
          ctx = { email: _vEmail, role: 'Superuser', canDo: function() { return true; } };
        } else {
          // Look up role from Users sheet
          try {
            var _uSheet = getDb(params).getSheetByName(SHEETS.USERS);
            if (_uSheet && _uSheet.getLastRow() >= 2) {
              var _uRows = _uSheet.getDataRange().getValues();
              var _found = false;
              for (var _ui = 1; _ui < _uRows.length; _ui++) {
                var _uEmail  = _uRows[_ui][0] ? _uRows[_ui][0].toString().toLowerCase().trim() : '';
                var _uActive = _uRows[_ui][4] !== false && _uRows[_ui][4] !== 'FALSE';
                if (_uEmail === _vEmail && _uActive) {
                  var _uRole = _uRows[_ui][1].toString() || 'ReadOnly';
                  ctx = { email: _vEmail, role: _uRole, canDo: function(a) { return _canDoPermission(_uRole, a); } };
                  _found = true;
                  break;
                }
              }
              if (!_found) {
                return { success: false, accessDenied: true, email: _vEmail,
                  error: 'Your account (' + _vEmail + ') is not registered for this instance.' };
              }
            } else {
              ctx = { email: _vEmail, role: 'Owner', canDo: function() { return true; } };
            }
          } catch(_ue) { Logger.log('verifiedEmail lookup error: ' + _ue); }
        }
      } else if (ctx.role === null) {
        return {
          success:      false,
          accessDenied: true,
          email:        ctx.email,
          error:        'Your account is not registered for this no~bull books instance.'
        };
      }
      var _inv = [], _bil = [], _settings = {}, _ba = [], _cli = [], _sup = [],
          _cn  = [], _po  = [], _bd  = [];
      try { _settings = getSettings(params)           || {}; } catch(e) { Logger.log('settings err: '+e); }
      try { _ba  = (getBankAccounts(params) ||{}).accounts       || []; } catch(e) { Logger.log('ba err: '+e); }
      try { _cli = (getAllClients(params)   ||{}).clients         || []; } catch(e) { Logger.log('cli err: '+e); }
      try { _sup = (getAllSuppliers(params) ||{}).suppliers       || []; } catch(e) { Logger.log('sup err: '+e); }
      try { _inv = (getAllInvoices(params)  ||{}).invoices        || []; } catch(e) { Logger.log('inv err: '+e); }
      try { _bil = (getAllBills(params)     ||{}).bills           || []; } catch(e) { Logger.log('bil err: '+e); }
      try { _cn  = (getCreditNotes(params) ||{}).creditNotes      || []; } catch(e) { Logger.log('cn err: '+e); }
      try { _po  = (getPurchaseOrders(null, params)||{}).purchaseOrders || []; } catch(e) { Logger.log('po err: '+e); }
      try { _bd  = (getBadDebts(params)    ||{}).badDebts         || []; } catch(e) { Logger.log('bd err: '+e); }
      var _licence = _checkLicence(params._sheetId);
      return {
        success:        true,
        user:           { email: ctx.email, role: ctx.role, permissions: _buildPermissions(ctx) },
        settings:       _settings,
        bankAccounts:   _ba,
        clients:        _cli,
        suppliers:      _sup,
        invoices:       _inv,
        bills:          _bil,
        creditNotes:    _cn,
        purchaseOrders: _po,
        badDebts:       _bd,
        licence:        _licence,
        voidedInvoices: _inv.filter(function(i){ return i.status==='Void'; }),
        voidedBills:    _bil.filter(function(b){ return b.status==='Void'; })
      };
    }

    case 'getSecondaryData':
      return {
        success:        true,
        creditNotes:    _safe(function(){ return getCreditNotes(params); }).creditNotes || [],
        purchaseOrders: _safe(function(){ return getPurchaseOrders(null, params); }).purchaseOrders || [],
        badDebts:       _safe(function(){ return getBadDebts(params); }).badDebts || [],
        recurringInvoices: _safe(function(){ return getAllRecurringInvoices(params); }).recurringInvoices || []
      };

    // -- AUTH / USERS ----------------------------------------------------------
    case 'getCurrentUser':  return { success: true, email: ctx.email, role: ctx.role };
    case 'getAllUsers':      _auth('users.view', params); return getAllUsers(params);
    case 'manageUser':      _auth('users.manage', params); return manageUser(params.action, params.email, params.role, params.notes, params);
    case 'runInitialSetup': return checkAndInitSheet(params);

    // -- SETTINGS -------------------------------------------------------------
    case 'getSettings':      return { success: true, settings: getSettings(params) };
    case 'updateSettings':   _auth('settings.write', params); return updateSettings(params, params);
    case 'uploadLogo':       _auth('settings.write', params); return uploadLogo(params);
    case 'getSecurityStatus': return getSecurityStatus(params);
    case 'getIntegrityStatus': _auth('maintenance.run', params); return getIntegrityStatus(params);
    case 'getAuditLog':      _auth('reports.read', params); return getAuditLog(params);

    // -- CLIENTS ---------------------------------------------------------------
    case 'getAllClients':     return getAllClients(params);
    case 'createClient':     _auth('clients.write', params); return createClient(params);
    case 'updateClient':     _auth('clients.write', params); return updateClient(params.clientId, params);
    case 'deleteClient':     _auth('clients.write', params); return deleteClient(params.clientId, params);
    case 'eraseClient':      _auth('clients.write', params); return eraseClient(params.clientId, params.retainFinancial);
    case 'exportClientData': _auth('clients.read',  params); return exportClientData(params.clientId);
    case 'generateClientStatement': _auth('invoices.read', params); return generateClientStatement(params.clientId, params.startDate, params.endDate, params);

    // -- INVOICES --------------------------------------------------------------
    case 'getAllInvoices':    return getAllInvoices(params);
    case 'createInvoice':    _auth('invoices.write', params); return createInvoice(params.clientId, params.lines, params.dueDate, params.notes, params.currency, params.exchangeRate, params.issueDate, params);
    case 'editInvoice':      _auth('invoices.write', params); return editInvoice(params.invoiceId, params, params);
    case 'deleteInvoice':    _auth('invoices.write', params); return deleteInvoice(params.invoiceId, params);
    case 'approveInvoice':   _auth('invoices.write', params); return approveInvoice(params.invoiceId, params);
    case 'markInvoiceSent':  _auth('invoices.write', params); return markInvoiceSent(params.invoiceId, params);
    case 'voidInvoice':      _auth('invoices.write', params); return voidInvoice(params.invoiceId, params.reason, params);
    case 'recordPayment':    _auth('invoices.write', params); return recordPaymentWithBank(params.invoiceId, params.amount, params.paymentDate, params.bankAccountId, params.notes, params);
    case 'generateInvoicePDF':  return generateInvoicePDF(params.invoiceId, params);
    case 'sendInvoiceEmail':    _auth('invoices.write', params); return sendInvoiceEmail(params.invoiceId, params.overrides||{}, params);
    case 'getInvoiceLines':     return getInvoiceLines(params.invoiceId, params);
    case 'getInvoiceFiles':     return getInvoiceFiles(params.invoiceId, params);
    case 'uploadInvoiceFile':   _auth('invoices.write', params); return uploadInvoiceFile(params);
    case 'deleteInvoiceFile':   _auth('invoices.write', params); return deleteInvoiceFile(params);
    case 'adjustInvoiceBalance': _auth('invoices.write', params); return adjustInvoiceBalance(params.invoiceId, params.amount, params);
    case 'writeOffBadDebt':  _auth('invoices.write', params); return writeOffBadDebt(params.invoiceId, params.reason, params);
    case 'checkBadDebtEligibility': return checkBadDebtVATEligibility(params.invoiceId, params);
    case 'getWhatsAppLink':  _auth('invoices.read', params); return getWhatsAppLink(params);
    case 'sendInvoiceWhatsApp': _auth('invoices.write', params); return sendInvoiceWhatsApp(params);

    // -- RECURRING INVOICES ----------------------------------------------------
    case 'getAllRecurring':           return getAllRecurringInvoices(params);
    case 'createRecurring':           _auth('invoices.write', params); return createRecurringInvoice(params);
    case 'updateRecurring':           _auth('invoices.write', params); return updateRecurringInvoice(params.recurringId, params);
    case 'deleteRecurring':           _auth('invoices.write', params); return deleteRecurringInvoice(params.recurringId, params);
    case 'processRecurringInvoices':  _auth('invoices.write', params); return processRecurringInvoices(params);
    case 'installRecurringTrigger':   _auth('maintenance.run', params); return installRecurringTrigger();

    // -- CREDIT NOTES ----------------------------------------------------------
    case 'getCreditNotes':   return getCreditNotes(params);
    case 'createCreditNote': _auth('invoices.write', params); return createCreditNote(params.invoiceId, params.lines, params.reason, params.issueDate, params);
    case 'applyCreditNote':  _auth('invoices.write', params); return applyCreditNote(params.cnId, params.invoiceId, params);
    case 'voidCreditNote':   _auth('invoices.write', params); return voidCreditNote(params.cnId, params);

    // -- SUPPLIERS -------------------------------------------------------------
    case 'getAllSuppliers':   return getAllSuppliers(params);
    case 'createSupplier':   _auth('suppliers.write', params); return createSupplier(params);
    case 'updateSupplier':   _auth('suppliers.write', params); return updateSupplier(params.supplierId, params);
    case 'deleteSupplier':   _auth('suppliers.write', params); return deleteSupplier(params.supplierId, params);

    // -- BILLS -----------------------------------------------------------------
    case 'getAllBills':       return getAllBills(params);
    case 'createBill':       _auth('bills.write', params); return createBill(params.supplierId, params.lines, params.issueDate, params.dueDate, params.notes, params);
    case 'getBillLines':     return getBillLines(params.billId, params);
    case 'editBill':         _auth('bills.write', params); return editBill(params.billId, params, params);
    case 'approveBill':      _auth('bills.write', params); return approveBill(params.billId, params);
    case 'deleteBill':       _auth('bills.write', params); return deleteBill(params.billId, params);
    case 'voidBill':         _auth('bills.write', params); return voidBill(params.billId, params.reason, params);
    case 'recordBillPayment': _auth('bills.write', params); return recordBillPaymentWithBank(params.billId, params.amount, params.paymentDate, params.bankAccountId, params.notes, params);
    case 'adjustBillBalance': _auth('bills.write', params); return adjustBillBalance(params.billId, params.amount, params);
    case 'getBillFiles':     return getBillFiles(params.billId, params);
    case 'uploadBillFile':   _auth('bills.write', params); return uploadBillFile(params);
    case 'deleteBillFile':   _auth('bills.write', params); return deleteBillFile(params);
    case 'generateRemittanceAdvice': _auth('bills.read', params); return generateRemittanceAdvice(params.billIds, params);

    // -- PURCHASE ORDERS -------------------------------------------------------
    case 'getPurchaseOrders':            return getPurchaseOrders(params.statusFilter, params);
    case 'getPurchaseOrderLines':        return getPurchaseOrderLines(params.poId, params);
    case 'createPurchaseOrder':          _auth('purchaseorders.write', params); return createPurchaseOrder(params.supplierId, params.lines, params.expectedDelivery, params.notes, params);
    case 'submitPurchaseOrderForApproval': _auth('purchaseorders.write', params); return submitPurchaseOrderForApproval(params.poId, params);
    case 'approvePurchaseOrder':         _auth('purchaseorders.write', params); return approvePurchaseOrder(params.poId, params);
    case 'markPurchaseOrderSent':        _auth('purchaseorders.write', params); return markPurchaseOrderSent(params.poId, params);
    case 'markPurchaseOrderPartial':     _auth('purchaseorders.write', params); return markPurchaseOrderPartial(params.poId, params);
    case 'receivePurchaseOrder':         _auth('purchaseorders.write', params); return receivePurchaseOrder(params.poId, params);
    case 'cancelPurchaseOrder':          _auth('purchaseorders.write', params); return cancelPurchaseOrder(params.poId, params);

    // -- BAD DEBTS -------------------------------------------------------------
    case 'getBadDebts':      return getBadDebts(params);
    case 'checkBadDebtVATEligibility': return checkBadDebtVATEligibility(params.invoiceId, params);

    // -- BANKING ---------------------------------------------------------------
    case 'getBankAccounts':             return getBankAccounts(params);
    case 'createBankAccount':           _auth('banking.write', params); return createBankAccount(params, params);
    case 'updateBankAccount':           _auth('banking.write', params); return updateBankAccount(params.accountId, params, params);
    case 'deleteBankAccount':           _auth('banking.write', params); return deleteBankAccount(params.accountId, params);
    case 'getBankTransactions':         return getBankTransactions(params.accountId, params.fromDate, params.toDate, params);
    case 'getUnreconciledTransactions': return getUnreconciledTransactions(params.accountId, params);
    case 'getReconciliationSummary':    return getReconciliationSummary(params.accountId, params);
    case 'reconcileTransaction':        _auth('banking.reconcile', params); return reconcileTransaction(params.transactionId, params.allocations, params);
    case 'createReconAdjustment':       _auth('banking.write', params); return createReconAdjustment(params);
    case 'spendMoney':                  _auth('banking.reconcile', params); return spendMoney(params, params);
    case 'receiveMoney':                _auth('banking.reconcile', params); return receiveMoney(params, params);
    case 'transferMoney':               _auth('banking.reconcile', params); return transferMoney(params, params);
    case 'importBankStatement':         _auth('banking.reconcile', params); return importBankStatement(params.accountId, params.csvData, params);
    case 'getUnallocatedInvoices':      return getUnallocatedInvoices(params.clientId, params);
    case 'getUnallocatedBills':         return getUnallocatedBills(params.supplierId, params);
    case 'quickMarkReconciled':         _auth('banking.reconcile', params); return quickMarkReconciled(params.transactionId, params);
    case 'matchSelected':               _auth('banking.reconcile', params); return matchSelected(params);

    // -- CHART OF ACCOUNTS -----------------------------------------------------
    case 'getAccounts':      return getAccounts(params.filters, params);
    case 'getAccountTypes':  return getAccountTypes(params);
    case 'getGeneralLedger': return getGeneralLedger(params.filters, params);
    case 'getTrialBalance':  return getTrialBalance(params.dateFrom, params.dateTo, params);
    case 'createAccount':    _auth('coa.write', params); return createAccount(params);
    case 'updateAccount':    _auth('coa.write', params); return updateAccount(params);
    case 'deleteAccount':    _auth('coa.write', params); return deleteAccount(params.accountCode, params);
    case 'seedCOA':          _auth('coa.write', params); return seedUKChartOfAccounts(params);

    // -- REPORTS ---------------------------------------------------------------
    case 'generateProfitLoss':   _auth('reports.read', params); return generateProfitLoss(params.startDate, params.endDate, params);
    case 'generateBalanceSheet': _auth('reports.read', params); return generateBalanceSheet(params);
    case 'generateCashFlow':     _auth('reports.read', params); return generateCashFlow(params.startDate, params.endDate, params);
    case 'getAllTransactions':   _auth('reports.read', params); return getAllTransactions(params.startDate, params.endDate, params);
    case 'getCurrencyBreakdown': return getCurrencyBreakdown(params.startDate, params.endDate, params);
    case 'getAvailableRates':    return getAvailableRates(params);
    case 'getExchangeRates':     return getExchangeRates(params);

    // -- VAT / MTD -------------------------------------------------------------
    case 'calculateVATReturn':  Logger.log('ROUTE calculateVATReturn _sheetId=' + params._sheetId + ' keys=' + Object.keys(params).join(',')); return calculateVATReturn(params.periodStart||params.fromDate, params.periodEnd||params.toDate, params);
    case 'saveVATReturn':       _auth('reports.tax', params); return saveVATReturn(params);
    case 'getVATReturns':       _auth('reports.tax', params); return getVATReturns(params);
    case 'getVATObligations': {
      var _vrn = params.vrn || (function(){ var s=getSettings(params); return (s.vatRegNumber||'').replace(/[^0-9]/g,''); })();
      return getVATObligations(_vrn, params.fromDate, params.toDate, params);
    }
    case 'submitVATReturn':     { var _v1=params.vrn||(function(){ var s=getSettings(params); return (s.vatRegNumber||'').replace(/[^0-9]/g,''); })(); return submitVATReturn(_v1, params.periodKey, params); }
    case 'getVATLiabilities':   { var _v2=params.vrn||(function(){ var s=getSettings(params); return (s.vatRegNumber||'').replace(/[^0-9]/g,''); })(); return getVATLiabilities(_v2, params.fromDate, params.toDate, params); }
    case 'getVATPayments':      { var _v3=params.vrn||(function(){ var s=getSettings(params); return (s.vatRegNumber||'').replace(/[^0-9]/g,''); })(); return getVATPayments(_v3, params.fromDate, params.toDate, params); }
    case 'getHMRCAuthStatus':   return getHMRCAuthStatus(params);
    case 'getHMRCAuthUrl':      return getHMRCManualAuthUrl(params);
    case 'exchangeHMRCCode':    _auth('credentials.manage', params); return exchangeHMRCCode(params.code, params);
    case 'testHMRCConnection':  return testHMRCConnection(params);

    // -- SA103 / SELF ASSESSMENT -----------------------------------------------
    case 'getSA103Data':           _auth('reports.tax', params); return getSA103Data(params);
    case 'getCapitalAllowances':   _auth('reports.tax', params); return getCapitalAllowances(params);
    case 'saveCapitalAllowance':   _auth('reports.tax', params); return saveCapitalAllowance(params);
    case 'deleteCapitalAllowance': _auth('reports.tax', params); return deleteCapitalAllowance(params);
    case 'saveSATaxAdjustments':   _auth('reports.tax', params); return saveSATaxAdjustments(params);

    // -- ITSA ------------------------------------------------------------------
    case 'getITSAObligationsFromSheet': return getITSAObligationsFromSheet(params);
    case 'getITSASubmissions':          return getITSASubmissions(params);
    case 'submitITSAQuarterlyUpdate':   _auth('mtd.submit', params); return submitQuarterlyUpdate(params.nino||'', params.businessId||'', params.taxYear, params.quarter, { turnover:params.turnover, expenses:params.expenses }, params);
    case 'triggerITSACalculation':      _auth('mtd.submit', params); return triggerAndGetCalculation(params.nino||'', params.taxYear||'', params);

    // -- FINANCIAL YEAR --------------------------------------------------------
    case 'getFinancialYears':       return getFinancialYears(params);
    case 'runPreCloseChecks':       _auth('settings.write', params); return runPreCloseChecks(params.yearEndDate, params);
    case 'previewImport':          return previewImport(params);
    case 'importContacts':         return importContacts(params);
    case 'importInvoices':         return importInvoices(params);
    case 'importBills':            return importBills(params);
    case 'importOpeningBalances':  return importOpeningBalances(params);
    case 'getImportTemplate':      return getImportTemplate(params);
    case 'calculateCT600':         return calculateCT600(params);
    case 'saveCT600Draft':          return saveCT600Draft(params);
    case 'getCT600Returns':         return getCT600Returns(params);
    case 'getYearEndSummary':       _auth('settings.write', params); return getYearEndSummary(params.yearStart, params.yearEnd, params);
    case 'closeFinancialYear':      _auth('settings.write', params); return closeFinancialYear(params);
    case 'reopenFinancialYear':     _auth('settings.write', params); return reopenFinancialYear(params.yearId, params.reason, params);

    // -- FIXED ASSETS ----------------------------------------------------------
    case 'getAllFixedAssets':        return getAllFixedAssets(params);
    case 'createFixedAsset':         _auth('coa.write', params); return createFixedAsset(params);
    case 'updateFixedAsset':         _auth('coa.write', params); return updateFixedAsset(params.assetId, params);
    case 'disposeFixedAsset':        _auth('coa.write', params); return disposeFixedAsset(params.assetId, params.disposalDate, params.disposalProceeds, params.notes, params);
    case 'runDepreciationRun':       _auth('coa.write', params); return runDepreciationRun(params.periodEndDate, params.periodMonths, params.postToLedger, params);
    case 'getDepreciationSchedule':  return getDepreciationSchedule(params.assetId, params);
    case 'getDepreciationRuns':      return getDepreciationRuns(params);

    // -- MAINTENANCE -----------------------------------------------------------
    case 'createBackup':          _auth('maintenance.run', params); return createBackup(params);
    case 'verifyIntegrity':       _auth('maintenance.run', params); return verifyIntegrity(params);
    case 'diagnoseSheets':        _auth('maintenance.run', params); return diagnoseSheets(params);
    case 'rebuildAccountBalances': _auth('maintenance.run', params); return rebuildAccountBalances(params);
    case 'cleanDuplicateTxns':    _auth('maintenance.run', params); return cleanDuplicateTransactions(params);
    case 'verifySchemaIntegrity': _auth('maintenance.run', params); return verifySchemaIntegrity(params);
    case 'runMaintenance':        _auth('maintenance.run', params); return runMaintenance(params.action, params);
    case 'installBackupTrigger':  _auth('maintenance.run', params); return installBackupTrigger();
    case 'runManualBackup':       _auth('maintenance.run', params); return runManualBackup(params);
    case 'getBackupStatus':       return getBackupStatus(params);
    case 'getInstanceMeta':       return getInstanceMeta(params);
    case 'getAllInstances':        return getAllInstances(params);
    case 'getAdminStats':         return getInstanceMeta(params);

    // -- REGISTRY / ADMIN ------------------------------------------------------
    case 'getAllRegistryClients':   return getAllRegistryClients(params);
    case 'registerClient':          _auth('settings.write', params); return registerClient(params, params);
    case 'updateRegistryClient':    _auth('settings.write', params); return updateRegistryClient(params.registryId, params, params);
    case 'deactivateRegistryClient':_auth('settings.write', params); return deactivateRegistryClient(params.registryId, params.reason, params);
    case 'provisionNewClient':      _auth('settings.write', params); return provisionNewClient(params);
    case 'activateClient':          _auth('settings.write', params); return activateClient(params.registryId, params);
    case 'getAdminStats':           return getAdminStats(params);

    // -- WHATSAPP --------------------------------------------------------------
    case 'getWhatsAppStatus':   return getWhatsAppStatus(params);
    case 'getTwilioStatus':     return getWhatsAppStatus(params);

    // -- VOID LOG --------------------------------------------------------------
    case 'getVoidLog':       return getVoidLog(params);

    default:
      return { success: false, error: 'Unknown action: ' + action, code: 404 };
  }
}

/**
 * _buildPermissions -- builds the permissions object from user context
 */
function _buildPermissions(ctx) {
  return {
    canWrite:             ctx.canDo('invoices.write'),
    canManageUsers:       ctx.canDo('users.manage'),
    canViewUsers:         ctx.canDo('users.view'),
    canManageSettings:    ctx.canDo('settings.write'),
    canViewReports:       ctx.canDo('reports.read'),
    canViewTaxReports:    ctx.canDo('reports.tax'),
    canManageBanking:     ctx.canDo('banking.reconcile'),
    canSubmitMTD:         ctx.canDo('mtd.submit'),
    canManageCredentials: ctx.canDo('credentials.manage'),
    canRunMaintenance:    ctx.canDo('maintenance.run'),
    canViewCOA:           ctx.canDo('coa.read'),
    canManageCOA:         ctx.canDo('coa.write'),
    canManageSuppliers:   ctx.canDo('suppliers.write')
  };
}

/**
 * _safe -- calls a function and returns empty object on error
 */
function _safe(fn) {
  try { return fn() || {}; } catch(e) { Logger.log('_safe: ' + e); return {}; }
}

/**
 * getWhatsAppStatus()
 * Returns WhatsApp/Twilio configuration status.
 */
function getWhatsAppStatus(params) {
  try {
    var props = PropertiesService.getScriptProperties().getProperties();
    var metaConfigured   = !!(props['META_PHONE_ID'] && props['META_ACCESS_TOKEN']);
    var twilioConfigured = !!(props['TWILIO_SID'] && props['TWILIO_TOKEN']);
    if (metaConfigured) {
      return { success: true, configured: true, provider: 'meta',
        meta: { phoneId: props['META_PHONE_ID'].substring(0, 6) + '...' } };
    }
    if (twilioConfigured) {
      return { success: true, configured: true, provider: 'twilio',
        twilio: { fromNumber: props['TWILIO_FROM'] || 'configured' } };
    }
    return { success: true, configured: false, provider: null };
  } catch(e) {
    return { success: false, configured: false, error: e.toString() };
  }
}
