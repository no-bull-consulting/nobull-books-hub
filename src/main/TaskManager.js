/**
 * NO~BULL BOOKS -- TAX MANAGER
 * Implements the logic for SA103 (Self Assessment).
 */

function getSA103Data(params) {
  try {
    _auth('reports.tax', params); //
    var taxYear = params.taxYear || '2025-26';
    
    // 1. Calculate dates for the tax year (6 April to 5 April)
    var yearParts = taxYear.split('-');
    var start = yearParts[0] + '-04-06';
    var end = '20' + yearParts[1] + '-04-05';

    // 2. Fetch P&L for the tax period
    var plResult = generateProfitLoss(start, end, params);
    if (!plResult.success) throw new Error(plResult.message);
    var pl = plResult.report;

    // 3. Simple UK Tax Estimation (2025/26 logic)
    var taxableProfit = pl.netProfit;
    var personalAllowance = 12570;
    var taxableIncome = Math.max(0, taxableProfit - personalAllowance);
    
    var incomeTax = 0;
    if (taxableIncome > 0) {
      var basicRate = Math.min(taxableIncome, 37700);
      incomeTax += basicRate * 0.20;
      // ... Higher rate logic as needed ...
    }

    // 4. Build Expense Categories mapped to SA103 boxes
    var expenseCats = {
      'car': { label: 'Car/Van/Travel', box: 'Box 12', total: 0 },
      'office': { label: 'Office/Phone/Stationery', box: 'Box 17', total: 0 }
      // ... map your COA categories here ...
    };

    return {
      success: true,
      taxYear: taxYear,
      grossIncome: pl.totalRevenue,
      totalExpenses: pl.totalExpenses,
      netProfitBeforeAllowances: pl.netProfit,
      taxableProfit: taxableProfit,
      personalAllowance: personalAllowance,
      taxableIncome: taxableIncome,
      incomeTax: incomeTax,
      totalTax: incomeTax, // Simplified for now
      effectiveRate: taxableProfit > 0 ? Math.round((incomeTax / taxableProfit) * 100) : 0,
      expenseCategories: expenseCats,
      assets: [] //
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Maps COA Nominal Codes to SA103 boxes
 */
function _mapSA103Categories(ledgerEntries) {
  var cats = {
    'car':    { label: 'Car, van and travel expenses', box: 'Box 12', total: 0 },
    'office': { label: 'Rent, rates, power and insurance costs', box: 'Box 14', total: 0 },
    'admin':  { label: 'Office expenses', box: 'Box 17', total: 0 }
  };

  ledgerEntries.forEach(function(t) {
    var code = String(t.accountDebit);
    
    // Map Car/Travel (COA 7200-7299)
    if (code.indexOf('72') === 0) cats.car.total += t.amount;
    
    // Map Rent/Power (COA 7000-7099)
    if (code.indexOf('70') === 0) cats.office.total += t.amount;
    
    // Map Admin/Stationery (COA 7100-7199)
    if (code.indexOf('71') === 0) cats.admin.total += t.amount;
  });

  return cats;
}