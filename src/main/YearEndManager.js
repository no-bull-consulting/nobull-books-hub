/**
 * NO~BULL BOOKS -- YEAR END MANAGER
 * Automates the closing of the financial year.
 */

function closeFinancialYear(params) {
  try {
    _auth('settings.write', params); //
    var ss = getDb(params);
    var yearStart = params.yearStart;
    var yearEnd = params.yearEnd;
    var reCode = params.retainedEarningsCode || '3000'; //

    // 1. Run Final P&L for the closing year
    var plResult = generateProfitLoss(yearStart, yearEnd, params);
    if (!plResult.success) throw new Error("P&L Calculation failed: " + plResult.message);
    var pl = plResult.report;

    // 2. Create Closing Journals
    // We must post entries that bring Revenue (Credit) and Expenses (Debit) to zero.
    var journalDate = new Date(yearEnd);
    var journalRef = "YEC-" + params.label;

    // Post Revenue Reversal: Debit Revenue, Credit Retained Earnings
    if (pl.totalRevenue > 0) {
      Object.keys(pl.revenue).forEach(function(code) {
        var amt = pl.revenue[code].amount || pl.revenue[code];
        if (amt === 0) return;
        createDoubleEntry(journalDate, 'YearEndClose', journalRef, code, reCode, amt, "Closing Revenue to Retained Earnings", null, null, params);
      });
    }

    // Post Expense Reversal: Debit Retained Earnings, Credit Expenses
    if (pl.totalExpenses > 0) {
      Object.keys(pl.expenses).forEach(function(code) {
        var amt = pl.expenses[code].amount || pl.expenses[code];
        if (amt === 0) return;
        createDoubleEntry(journalDate, 'YearEndClose', journalRef, reCode, code, amt, "Closing Expenses to Retained Earnings", null, null, params);
      });
    }

    // 3. Update the Financial Years Archive Sheet
    var yearId = generateId('FY');
    var fySheet = ss.getSheetByName(SHEETS.FINANCIAL_YEARS);
    if (fySheet) {
      fySheet.appendRow([
        yearId,
        params.label,
        yearStart,
        yearEnd,
        'Closed',
        new Date(),
        _getCurrentUserContext(params).email
      ]);
    }

    // 4. Update System Settings: Advance the year and LOCK the period
    var settings = getSettings(params);
    settings.lockedBefore = yearEnd; // All transactions on/before this date are now read-only
    
    // Advance next FY dates (defaulting to next calendar year)
    var nextStart = new Date(yearStart);
    nextStart.setFullYear(nextStart.getFullYear() + 1);
    var nextEnd = new Date(yearEnd);
    nextEnd.setFullYear(nextEnd.getFullYear() + 1);
    
    settings.financialYearStart = nextStart.toISOString().split('T')[0];
    settings.financialYearEnd = nextEnd.toISOString().split('T')[0];
    
    updateSettings(settings, params);

    logAudit('YEAR_END_CLOSE', 'FinancialYear', yearId, { label: params.label, profit: pl.netProfit }, params);

    return { 
      success: true, 
      label: params.label, 
      netProfit: pl.netProfit, 
      lockedBefore: yearEnd 
    };
  } catch(e) {
    return apiError("Year-end close failed", e);
  }
}