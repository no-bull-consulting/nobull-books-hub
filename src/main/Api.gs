/**
 * NO~BULL BOOKS — API LAYER
 * ──────────────────────────────────────────────────────────────────────────────
 * Single entry point: handleApiCall(action, paramsJson)
 * Called via google.script.run.handleApiCall() from the frontend.
 *
 * KEY CHANGES (hub model):
 *  - params._sheetId is threaded through to every function that reads/writes data
 *  - _auth(action, params) now passes params so Users sheet lookup hits the right sheet
 *  - duplicate updateSettings() removed — canonical version is in Settings.gs
 *  - uploadLogo now receives full params object (not positional args)
 */

var API_VERSION = '1.0';

/**
 * handleApiCall(action, paramsJson)
 * The single google.script.run entry point.
 */
function handleApiCall(action, paramsJson) {
  try {
    var params = JSON.parse(paramsJson || '{}');
    var ctx    = _getCurrentUserContext(params);

    // Gemini route
    if (action === 'askGemini') return handleGeminiRequest(params, ctx);

    var result = _route(action, params, ctx);
    return JSON.parse(JSON.stringify(result));
  } catch(e) {
    Logger.log('handleApiCall ERROR [' + action + ']: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * _route — maps action strings to module functions.
 * params is always passed through so _sheetId reaches every data function.
 */
function _route(action, params, ctx) {
  switch (action) {

    // ── STARTUP ──────────────────────────────────────────────────────────────
    case 'getStartupData': {
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
      // ── Registry / licence check ─────────────────────────────────────────
      var _licence = _checkLicence(params._sheetId);

      return {
        success:        true,
        licence:        _licence,
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
        voidedInvoices: _inv.filter(function(i){ return i.status==='Void'||i.status==='Voided'; }),
        voidedBills:    _bil.filter(function(b){ return b.status==='Void'||b.status==='Voided'; })
      };
    }

    case 'getSecondaryData':
      return {
        success:        true,
        creditNotes:    _safe(function(){ return getCreditNotes(params); }).creditNotes     || [],
        purchaseOrders: _safe(function(){ return getPurchaseOrders(null, params); }).purchaseOrders || [],
        badDebts:       _safe(function(){ return getBadDebts(params); }).badDebts           || [],
        voidLog:        _safe(function(){ return getVoidLog(params); })
      };

    // ── INIT / SETUP ──────────────────────────────────────────────────────────
    // No _auth here — bootstrap runs before Users sheet exists
    case 'runInitialSetup':  return checkAndInitSheet(params);

    // ── AUTH / USERS ──────────────────────────────────────────────────────────
    case 'getCurrentUser':
      return getCurrentUserWithRole(params);
    case 'getAllUsers':
      _auth('users.view', params);
      return getAllUsers(params);
    case 'manageUser':
      _auth('users.manage', params);
      return manageUser(params.action, params.email, params.role, params.notes, params);

    // ── SETTINGS ─────────────────────────────────────────────────────────────
    case 'getSettings':      return getSettings(params);
    case 'updateSettings':   _auth('settings.write', params); return updateSettings(params);
    case 'uploadLogo':       _auth('settings.write', params); return uploadLogo(params);
    case 'getSecurityStatus': return getSecurityStatus();

    // ── CLIENTS ───────────────────────────────────────────────────────────────
    case 'getAllClients':     return getAllClients(params);
    case 'createClient':     _auth('clients.write', params); return createClient(params);
    case 'updateClient':     _auth('clients.write', params); return updateClient(params.clientId, params);
    case 'eraseClient':      _auth('clients.write', params); return eraseClient(params.clientId, params.retainFinancial);
    case 'exportClientData': _auth('clients.read',  params); return exportClientData(params.clientId);

    // ── INVOICES ──────────────────────────────────────────────────────────────
    case 'getAllInvoices':    return getAllInvoices(params);
    case 'createInvoice':    _auth('invoices.write', params); return createInvoice(params.clientId, params.lines, params.dueDate, params.notes, params.currency, params.exchangeRate, params);
    case 'editInvoice':      _auth('invoices.write', params); return editInvoice(params.invoiceId, params);
    case 'deleteInvoice':    _auth('invoices.write', params); return deleteInvoice(params.invoiceId, params);
    case 'approveInvoice':   _auth('invoices.write', params); return approveInvoice(params.invoiceId, params);
    case 'markInvoiceSent':  _auth('invoices.write', params); return markInvoiceSent(params.invoiceId, params);
    case 'recordPayment':    _auth('invoices.write', params); return recordPaymentWithBank(params.invoiceId, params.amount, params.paymentDate, params.bankAccountId, params.notes, params);
    case 'voidInvoice':      _auth('invoices.write', params); return voidInvoice(params.invoiceId, params.reason, params);
    case 'writeOffInvoice':  _auth('invoices.write', params); return writeOffInvoice(params.invoiceId, params.writeOffDate, params.reason, params);
    case 'getInvoiceLines':  return getInvoiceLines(params.invoiceId, params);
    case 'getInvoiceFiles':  return getInvoiceFiles(params.invoiceId, params);
    case 'uploadInvoiceFile':_auth('invoices.write', params); return uploadInvoiceFile(params.invoiceId, params.base64Data, params.fileName, params.fileType, params.description, params);
    case 'deleteInvoiceFile':_auth('invoices.write', params); return deleteInvoiceFile(params.fileId, params);
    case 'generateInvoicePDF': return generateInvoicePDF(params.invoiceId, params);
    case 'sendInvoiceEmail': _auth('invoices.write', params); return sendInvoiceEmail(params.invoiceId, params.overrides||{}, params);
    case 'generateClientStatement': _auth('invoices.read', params); return generateClientStatement(params.clientId, params.startDate, params.endDate, params);
    case 'adjustInvoiceBalance':    _auth('invoices.write', params); return adjustInvoiceBalance(params.invoiceId, params.amount, params.adjustmentType, params.reason, params.accountCode, params);

    // ── RECURRING INVOICES ────────────────────────────────────────────────────
    case 'getAllRecurring':           return getAllRecurring(params);
    case 'createRecurring':          _auth('invoices.write', params); return createRecurring(params);
    case 'updateRecurring':          _auth('invoices.write', params); return updateRecurring(params.recurringId, params);
    case 'deleteRecurring':          _auth('invoices.write', params); return deleteRecurring(params.recurringId, params);
    case 'processRecurringInvoices': _auth('invoices.write', params); return processRecurringInvoices(params);
    case 'installRecurringTrigger':  _auth('maintenance.run', params); return installRecurringTrigger();

    // ── CREDIT NOTES ──────────────────────────────────────────────────────────
    case 'getCreditNotes':   return getCreditNotes(params);
    case 'createCreditNote': _auth('invoices.write', params); return createCreditNote(params.invoiceId, params.lines, params.reason, params.issueDate, params);
    case 'applyCreditNote':  _auth('invoices.write', params); return applyCreditNote(params.cnId, params.invoiceId, params);
    case 'voidCreditNote':   _auth('invoices.write', params); return voidCreditNote(params.cnId, params.reason, params);

    // ── SUPPLIERS ─────────────────────────────────────────────────────────────
    case 'getAllSuppliers':   return getAllSuppliers(params);
    case 'createSupplier':   _auth('suppliers.write', params); return createSupplier(params);
    case 'updateSupplier':   _auth('suppliers.write', params); return updateSupplier(params.supplierId, params);
    case 'deleteClient':   _auth('clients.write', params); return deleteClient(params.clientId, params);
    case 'deleteSupplier':   _auth('suppliers.write', params); return deleteSupplier(params.supplierId, params);

    // ── BILLS ─────────────────────────────────────────────────────────────────
    case 'getAllBills':        return getAllBills(params);
    case 'createBill':        _auth('bills.write', params); return createBill(params.supplierId, params.lines, params.issueDate, params.dueDate, params.notes, params);
    case 'editBill':          _auth('bills.write', params); return editBill(params.billId, params);
    case 'deleteBill':        _auth('bills.write', params); return deleteBill(params.billId, params);
    case 'approveBill':       _auth('bills.write', params); return approveBill(params.billId, params);
    case 'recordBillPayment': _auth('bills.write',  params); return recordBillPaymentWithBank(params.billId, params.amount, params.paymentDate, params.bankAccountId, params.notes, params);
    case 'voidBill':          _auth('bills.write',  params); return voidBill(params.billId, params.reason, params);
    case 'getBillLines':      return getBillLines(params.billId, params);
    case 'getBillFiles':      return getBillFiles(params.billId, params);
    case 'uploadBillFile':    _auth('bills.write', params); return uploadBillFile(params.billId, params.base64Data, params.fileName, params.fileType, params.description, params);
    case 'deleteBillFile':    _auth('bills.write', params); return deleteBillFile(params.fileId, params);
    case 'adjustBillBalance': _auth('bills.write', params); return adjustBillBalance(params.billId, params.amount, params.adjustmentType, params.reason, params.accountCode, params);
    case 'generateRemittanceAdvice': _auth('bills.read', params); return generateRemittanceAdvice(params.billIds, params.paymentDate, params.paymentRef, params.bankAccountId, params);

    // ── PURCHASE ORDERS ───────────────────────────────────────────────────────
    case 'getPurchaseOrders':        return getPurchaseOrders(params.statusFilter||null, params);
    case 'createPurchaseOrder':      _auth('purchaseorders.write', params); return createPurchaseOrder(params.supplierId, params.lines, params.expectedDelivery, params.notes, params);
    case 'getPOLines':               return getPurchaseOrderLines(params.poId, params);
    case 'submitPO':                 _auth('purchaseorders.write',   params); return submitPurchaseOrderForApproval(params.poId, params);
    case 'approvePO':                _auth('purchaseorders.approve', params); return approvePurchaseOrder(params.poId, params.notes, params);
    case 'approvePurchaseOrder':     _auth('purchaseorders.approve', params); return approvePurchaseOrder(params.poId, params.notes, params);
    case 'markPOSent':               _auth('purchaseorders.write',   params); return markPurchaseOrderSent(params.poId, params);
    case 'markPurchaseOrderSent':    _auth('purchaseorders.write',   params); return markPurchaseOrderSent(params.poId, params);
    case 'markPurchaseOrderPartial': _auth('purchaseorders.write',   params); return markPurchaseOrderPartial(params.poId, params.notes, params);
    case 'receivePO':                _auth('purchaseorders.write',   params); return receivePurchaseOrder(params.poId, params.invoiceRef, params.receiptDate, params.notes, params);
    case 'cancelPO':                 _auth('purchaseorders.write',   params); return cancelPurchaseOrder(params.poId, params.reason, params);

    // ── BAD DEBTS ─────────────────────────────────────────────────────────────
    case 'getBadDebts':              return getBadDebts(params);
    case 'writeOffBadDebt':          _auth('invoices.write', params); return writeOffBadDebt(params.invoiceId, params.reason, params.writeOffDate, params);
    case 'checkBadDebtEligibility':  return checkBadDebtVATEligibility(params.invoiceId, params);
    case 'markBadDebtVATClaimed':    _auth('invoices.write', params); return markBadDebtVATClaimed(params.badDebtId, params.claimDate, params);

    // ── VOID LOG ──────────────────────────────────────────────────────────────
    case 'getVoidLog':               return getVoidLog(params);

    // ── BANKING ───────────────────────────────────────────────────────────────
    case 'getBankAccounts':             return getBankAccounts(params);
    case 'createBankAccount':          _auth('banking.write', params); return createBankAccount(params, params);
    case 'getBankTransactions':         return getBankTransactions(params.accountId, params.fromDate, params.toDate, params);
    case 'getUnreconciledTransactions': return getUnreconciledTransactions(params.accountId, params);
    case 'getReconciliationSummary':    return getReconciliationSummary(params.accountId, params);
    case 'reconcileTransaction':       _auth('banking.reconcile', params); return reconcileTransaction(params.transactionId, params.allocations, params);
    case 'createReconAdjustment':      _auth('banking.write',     params); return createReconAdjustment(params);
    case 'spendMoney':                 _auth('banking.reconcile', params); return spendMoney(params);
    case 'receiveMoney':               _auth('banking.reconcile', params); return receiveMoney(params);
    case 'transferMoney':              _auth('banking.reconcile', params); return transferMoney(params);
    case 'importBankStatement':        _auth('banking.reconcile', params); return importBankStatement(params.accountId, params.csvData, params);
    case 'getUnallocatedInvoices':      return getUnallocatedInvoices(params.clientId, params);
    case 'getUnallocatedBills':         return getUnallocatedBills(params.supplierId, params);

    // ── CHART OF ACCOUNTS ─────────────────────────────────────────────────────
    case 'getAccounts':      return getAccounts(params.filters, params);
    case 'getAccountTypes':  return getAccountTypes(params);
    case 'getGeneralLedger': return getGeneralLedger(params.filters, params);
    case 'getTrialBalance':  return getTrialBalance(params.dateFrom, params.dateTo, params);
    case 'createAccount':   _auth('coa.write', params); return createAccount(params);
    case 'updateAccount':   _auth('coa.write', params); return updateAccount(params);
    case 'deleteAccount':   _auth('coa.write', params); return deleteAccount(params.accountCode, params);

    // ── REPORTS ───────────────────────────────────────────────────────────────
    case 'getAvailableRates':    return getAvailableRates(params);
    case 'getExchangeRates':     return getExchangeRates(params);
    case 'getCurrencyBreakdown': return getCurrencyBreakdown(params.startDate, params.endDate, params);
    case 'generateCashFlow':    _auth('reports.read', params); return generateCashFlow(params.startDate, params.endDate, params);
    case 'generateProfitLoss':  _auth('reports.read', params); return generateProfitLoss(params.startDate, params.endDate, params);
    case 'generateBalanceSheet':_auth('reports.read', params); return generateBalanceSheet(params);
    case 'getAllTransactions':   _auth('reports.read', params); return getAllTransactions(params.startDate, params.endDate, params);
    case 'getAuditLog':          _auth('reports.read', params); return getAuditLog(params);

    // ── VAT / MTD ─────────────────────────────────────────────────────────────
    case 'calculateVATReturn':  _auth('reports.tax', params); return calculateVATReturn(params.periodStart, params.periodEnd, params);
    case 'saveVATReturn':       _auth('reports.tax', params); return saveVATReturn(params);
    case 'getVATReturns':       _auth('reports.tax', params); return getVATReturns(params);
    case 'getVATObligations':   _auth('reports.tax', params); return getVATObligations(params.vrn, params.fromDate, params.toDate, params);
    case 'submitVATReturn':     _auth('mtd.submit',  params); return submitVATReturn(params.vrn, params.periodKey, params);
    case 'getVATLiabilities':   _auth('reports.tax', params); return getVATLiabilities(params.vrn, params.fromDate, params.toDate, params);
    case 'getVATPayments':      _auth('reports.tax', params); return getVATPayments(params.vrn, params.fromDate, params.toDate, params);
    case 'getHMRCAuthStatus':   return getHMRCAuthStatus(params);
    case 'getHMRCAuthUrl':      return getHMRCManualAuthUrl(params);
    case 'exchangeHMRCCode':   _auth('credentials.manage', params); return exchangeHMRCCode(params.code, params);
    case 'testHMRCConnection':  return testHMRCConnection(params);

    // ── WHATSAPP / EMAIL ──────────────────────────────────────────────────────
    case 'sendInvoiceWhatsApp': _auth('invoices.write', params); return sendInvoiceWhatsApp(params.invoiceId, params.overrides||{}, params);
    case 'getWhatsAppLink':     _auth('invoices.read',  params); return getWhatsAppLink(params.invoiceId, params);
    case 'getWhatsAppStatus':   return getWhatsAppStatus();
    case 'getTwilioStatus':     return getWhatsAppStatus(); // backwards compat

    // ── SA103 / SELF ASSESSMENT ───────────────────────────────────────────────
    case 'getSA103Data':           _auth('reports.tax', params); return getSA103Data(params.taxYear, params);
    case 'getCapitalAllowances':   _auth('reports.tax', params); return getCapitalAllowances(params.taxYear, params);
    case 'saveCapitalAllowance':   _auth('reports.tax', params); return saveCapitalAllowance(params);
    case 'deleteCapitalAllowance': _auth('reports.tax', params); return deleteCapitalAllowance(params.assetId, params);
    case 'saveSATaxAdjustments':   _auth('reports.tax', params); return saveSATaxAdjustments(params.taxYear, params.ownUse, params.otherAdj, params);

    // ── FINANCIAL YEAR CLOSE ──────────────────────────────────────────────────
    case 'getFinancialYears':       return getFinancialYears(params);
    case 'runPreCloseChecks':      _auth('settings.write', params); return runPreCloseChecks(params.yearEndDate, params);
    case 'getYearEndSummary':      _auth('settings.write', params); return getYearEndSummary(params.yearStart, params.yearEnd, params);
    case 'closeFinancialYear':     _auth('settings.write', params); return closeFinancialYear(params);
    case 'reopenFinancialYear':    return reopenFinancialYear(params.yearId, params.reason, params);
    case 'initFinancialYearSheets':_auth('maintenance.run', params); return initFinancialYearSheets(params);

    // ── FIXED ASSETS ──────────────────────────────────────────────────────────
    case 'getAllFixedAssets':        return getAllFixedAssets(params);
    case 'createFixedAsset':        _auth('coa.write', params); return createFixedAsset(params);
    case 'updateFixedAsset':        _auth('coa.write', params); return updateFixedAsset(params.assetId, params);
    case 'disposeFixedAsset':       _auth('coa.write', params); return disposeFixedAsset(params.assetId, params.disposalDate, params.disposalProceeds, params.notes, params);
    case 'runDepreciationRun':      _auth('coa.write', params); return runDepreciationRun(params.periodEndDate, params.periodMonths, params.postToLedger, params);
    case 'getDepreciationSchedule': return getDepreciationSchedule(params.assetId, params);
    case 'getDepreciationRuns':     return getDepreciationRuns(params);
    case 'initFixedAssetSheets':   _auth('maintenance.run', params); return initFixedAssetSheets(params);

    // ── ITSA ──────────────────────────────────────────────────────────────────
    case 'getITSAObligationsFromSheet': return getITSAObligationsFromSheet(params);
    case 'getITSASubmissions':          return getITSASubmissions(params);
    case 'submitITSAQuarterlyUpdate':  _auth('mtd.submit', params); return submitQuarterlyUpdate(params.nino||'', params.businessId||'', params.taxYear, params.quarter, { turnover:params.turnover, expenses:params.expenses }, params);
    case 'triggerITSACalculation':     _auth('mtd.submit', params); return triggerAndGetCalculation(params.nino||'', params.taxYear||'', params);

    // ── MAINTENANCE (superuser only) ──────────────────────────────────────────
    case 'initializeSystem':       _auth('maintenance.run', params); return initializeSystem(params);
    case 'createBackup':           _auth('maintenance.run', params); return createBackup(params);
    case 'verifyIntegrity':        _auth('maintenance.run', params); return verifyIntegrity(params);
    case 'diagnoseSheets':         _auth('maintenance.run', params); return diagnoseSheets(params);
    case 'rebuildAccountBalances': _auth('maintenance.run', params); return rebuildAccountBalances(params);
    case 'cleanDuplicateTxns':     _auth('maintenance.run', params); return cleanDuplicateTransactions(params);
    case 'verifySchemaIntegrity':  _auth('maintenance.run', params); return verifySchemaIntegrity(params);
    case 'getIntegrityStatus':     _auth('maintenance.run', params); return getIntegrityStatus(params);
    case 'installBackupTrigger':   _auth('maintenance.run', params); return installBackupTrigger();
    case 'removeBackupTrigger':    _auth('maintenance.run', params); return removeBackupTrigger();
    case 'runManualBackup':        _auth('maintenance.run', params); return runManualBackup();
    case 'getBackupStatus':        _auth('maintenance.run', params); return getBackupStatus();
    case 'runSandboxValidation':   _auth('maintenance.run', params); return runSandboxValidation();
    case 'sandboxVATDryRun':       _auth('maintenance.run', params); return sandboxVATSubmitDryRun();
    case 'getAllInstances':         _auth('maintenance.run', params); return getAllInstances();
    case 'getInstanceMeta':         _auth('maintenance.run', params); return getInstanceMeta();
    case 'protectSensitiveSheets': _auth('maintenance.run', params); return protectSensitiveSheets();

    // ── CLIENT STATEMENT / REMITTANCE ─────────────────────────────────────────
    case 'generateClientStatement':  _auth('invoices.read', params); return generateClientStatement(params.clientId, params.startDate, params.endDate, params);
    case 'generateRemittanceAdvice': _auth('bills.read',    params); return generateRemittanceAdvice(params.billIds, params.paymentDate, params.paymentRef, params.bankAccountId, params);

    case 'seedCOA': _auth('coa.write', params); return seedUKChartOfAccounts(params);

    // Onboarding & provisioning
    case 'provisionNewClient':   _auth('settings.write', params); return provisionNewClient(params);
    case 'sendWelcomeEmail':     _auth('settings.write', params); return sendWelcomeEmail(params);
    case 'resendWelcomeEmail':   _auth('settings.write', params); return sendWelcomeEmail(params);

    // Registry
    case 'getRegistrySummary':     return getRegistrySummary(params);
    case 'getAllRegistryClients':   return getAllRegistryClients(params);
    case 'registerClient':          _auth('settings.write', params); return registerClient(params, params);
    case 'updateRegistryClient':    _auth('settings.write', params); return updateRegistryClient(params.registryId, params, params);
    case 'deactivateRegistryClient':_auth('settings.write', params); return deactivateRegistryClient(params.registryId, params.reason, params);

    default:
      return { success: false, error: 'Unknown action: ' + action, code: 404 };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────

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

/** Calls fn() and returns empty object on error — used in getSecondaryData */
function _safe(fn) {
  try { return fn() || {}; } catch(e) { Logger.log('_safe: ' + e); return {}; }
}

/**
 * getWhatsAppStatus()
 * Returns WhatsApp/Twilio configuration status.
 * Full implementation pending — returns not-configured stub for now.
 */
function getWhatsAppStatus() {
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