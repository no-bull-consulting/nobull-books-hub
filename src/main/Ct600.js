/**
 * NO~BULL BOOKS -- CORPORATION TAX CT600
 * -----------------------------------------------------------------------------
 * Calculates CT600 figures from the Chart of Accounts and Transactions.
 * Produces a draft for accountant review -- NOT a direct HMRC submission.
 *
 * CT600 boxes covered:
 *   Box 1   -- Company name
 *   Box 30  -- Turnover (total revenue)
 *   Box 35  -- Trading profit before deductions
 *   Box 40  -- Capital allowances
 *   Box 45  -- Trading profit after capital allowances
 *   Box 155 -- Profits chargeable to CT
 *   Box 160 -- Franked investment income
 *   Box 185 -- CT at main rate (25% from Apr 2023)
 *   Box 188 -- Marginal relief (if profits GBP 50k-GBP 250k)
 *   Box 190 -- CT chargeable
 *   Box 235 -- Income tax deducted
 *   Box 440 -- Tax payable
 *
 * UK CT Rates (FY2023 onwards):
 *   Small profits rate:  19% on profits up to GBP 50,000
 *   Main rate:           25% on profits over GBP 250,000
 *   Marginal relief:     Between GBP 50,000 and GBP 250,000
 */

var CT_RATES = {
  SMALL_RATE:       0.19,   // 19% -- profits up to GBP 50k
  MAIN_RATE:        0.25,   // 25% -- profits over GBP 250k
  SMALL_LIMIT:      50000,
  UPPER_LIMIT:      250000,
  MRF:              11/400  // Marginal Relief Fraction
};

var CT600_SHEET_NAME = 'CT600Returns';

// -----------------------------------------------------------------------------
// MAIN CALCULATION
// -----------------------------------------------------------------------------

function calculateCT600(params) {
  try {
    var periodStart = params.periodStart;
    var periodEnd   = params.periodEnd;
    if (!periodStart || !periodEnd) return { success: false, message: 'Accounting period start and end dates are required.' };

    var ss       = getDb(params || {});
    var settings = getSettings(params);

    // -- Get P&L data ----------------------------------------------------------
    var plResult = generateProfitLoss(periodStart, periodEnd, params);
    Logger.log('CT600 plResult: ' + JSON.stringify(plResult ? {success:plResult.success, message:plResult.message} : null));
    if (!plResult) return { success: false, message: 'generateProfitLoss returned null' };
    if (!plResult.success) return { success: false, message: 'P&L error: ' + (plResult.message || plResult.error || 'unknown') };
    var pl = plResult.report || {};
    Logger.log('CT600 pl keys: ' + Object.keys(pl).join(','));

    // -- Revenue ---------------------------------------------------------------
    var turnover = 0;
    if (pl.revenue) {
      Object.keys(pl.revenue).forEach(function(k) {
        turnover += parseFloat(pl.revenue[k].amount || pl.revenue[k] || 0);
      });
    }

    // -- Expenses --------------------------------------------------------------
    var totalExpenses    = 0;
    var depreciation     = 0;
    var otherExpenses    = 0;

    if (pl.expenses) {
      Object.keys(pl.expenses).forEach(function(k) {
        var amt = parseFloat(pl.expenses[k].amount || pl.expenses[k] || 0);
        totalExpenses += amt;
        // Separate depreciation (account codes 8xxx typically)
        var code = (pl.expenses[k].code || k || '').toString();
        if (code.match(/^8[0-9]/) || k.toLowerCase().indexOf('depreciat') >= 0) {
          depreciation += amt;
        } else {
          otherExpenses += amt;
        }
      });
    }

    // -- Capital Allowances ----------------------------------------------------
    // Pull from FixedAssets sheet -- Annual Investment Allowance + WDA
    var capitalAllowances = _getCapitalAllowances(ss, periodStart, periodEnd);

    // -- Trading Profit Computation --------------------------------------------
    var netProfit          = turnover - totalExpenses;
    // Add back depreciation (not allowable), deduct capital allowances
    var tradingProfitAdj   = netProfit + depreciation - capitalAllowances;
    var profitsChargeable  = Math.max(0, tradingProfitAdj);

    // -- CT Computation --------------------------------------------------------
    var ctComputation      = _computeCT(profitsChargeable);

    // -- Prior Year ------------------------------------------------------------
    var priorYear = null;
    if (params.includePriorYear) {
      var pyStart = _shiftYear(periodStart, -1);
      var pyEnd   = _shiftYear(periodEnd, -1);
      try {
        var pyResult = calculateCT600(Object.assign({}, params, {
          periodStart: pyStart, periodEnd: pyEnd, includePriorYear: false
        }));
        if (pyResult.success) priorYear = pyResult.ct600;
      } catch(e) { Logger.log('Prior year CT600 error: ' + e); }
    }

    var ct600 = {
      // Period
      periodStart:        periodStart,
      periodEnd:          periodEnd,
      companyName:        settings.companyName || '',
      companyNumber:      settings.companyNumber || '',
      utr:                settings.utr || '',

      // Box 30 -- Turnover
      box30_turnover:     _r(turnover),

      // Box 35 -- Trading profit before adjustments
      box35_tradingProfit: _r(netProfit),

      // Adjustments
      addBackDepreciation: _r(depreciation),
      lessCapitalAllowances: _r(capitalAllowances),

      // Box 45 -- Adjusted trading profit
      box45_adjProfit:    _r(tradingProfitAdj),

      // Box 155 -- Profits chargeable to CT
      box155_chargeable:  _r(profitsChargeable),

      // CT computation
      box185_ctMainRate:  _r(ctComputation.taxAtMainRate),
      box188_marginalRelief: _r(ctComputation.marginalRelief),
      box190_ctChargeable: _r(ctComputation.ctChargeable),
      box440_taxPayable:   _r(ctComputation.ctChargeable),

      // Rates applied
      rateApplied:        ctComputation.rateDescription,
      effectiveRate:      profitsChargeable > 0 ? _r((ctComputation.ctChargeable / profitsChargeable) * 100) : 0,

      // Supporting figures
      totalRevenue:       _r(turnover),
      totalExpenses:      _r(totalExpenses),
      netProfit:          _r(netProfit),
      depreciation:       _r(depreciation),
      capitalAllowances:  _r(capitalAllowances),

      // Status
      status:             'Draft',
      calculatedAt:       new Date().toISOString(),
      priorYear:          priorYear
    };

    return { success: true, ct600: ct600 };
  } catch(e) {
    Logger.log('calculateCT600 error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// -----------------------------------------------------------------------------
// CT RATE COMPUTATION
// -----------------------------------------------------------------------------

function _computeCT(profits) {
  if (profits <= 0) {
    return { taxAtMainRate: 0, marginalRelief: 0, ctChargeable: 0, rateDescription: 'No profits' };
  }

  if (profits <= CT_RATES.SMALL_LIMIT) {
    // Small profits rate -- 19%
    var tax = profits * CT_RATES.SMALL_RATE;
    return {
      taxAtMainRate:   _r(tax),
      marginalRelief:  0,
      ctChargeable:    _r(tax),
      rateDescription: '19% small profits rate'
    };
  }

  if (profits >= CT_RATES.UPPER_LIMIT) {
    // Main rate -- 25%
    var tax2 = profits * CT_RATES.MAIN_RATE;
    return {
      taxAtMainRate:   _r(tax2),
      marginalRelief:  0,
      ctChargeable:    _r(tax2),
      rateDescription: '25% main rate'
    };
  }

  // Marginal relief band -- GBP 50k to GBP 250k
  var taxAtMain = profits * CT_RATES.MAIN_RATE;
  var mrf       = CT_RATES.MRF * (CT_RATES.UPPER_LIMIT - profits) * (profits / profits);
  var netTax    = taxAtMain - mrf;
  return {
    taxAtMainRate:   _r(taxAtMain),
    marginalRelief:  _r(mrf),
    ctChargeable:    _r(netTax),
    rateDescription: '25% main rate with marginal relief'
  };
}

// -----------------------------------------------------------------------------
// CAPITAL ALLOWANCES
// -----------------------------------------------------------------------------

function _getCapitalAllowances(ss, periodStart, periodEnd) {
  try {
    var sheet = ss.getSheetByName('FixedAssets');
    if (!sheet || sheet.getLastRow() < 2) return 0;

    var data  = sheet.getDataRange().getValues();
    var start = new Date(periodStart);
    var end   = new Date(periodEnd);
    var total = 0;

    for (var i = 1; i < data.length; i++) {
      var purchaseDate = data[i][4] ? new Date(data[i][4]) : null;
      if (!purchaseDate || purchaseDate < start || purchaseDate > end) continue;
      if (data[i][9] === 'Disposed') continue;

      // Annual Investment Allowance -- 100% of cost up to AIA limit
      var cost = parseFloat(data[i][5]) || 0;
      total += cost; // simplified -- full AIA assumed (GBP 1m limit from 2023)
    }
    return total;
  } catch(e) {
    Logger.log('_getCapitalAllowances error: ' + e);
    return 0;
  }
}

// -----------------------------------------------------------------------------
// SAVE / LOAD DRAFTS
// -----------------------------------------------------------------------------

function saveCT600Draft(params) {
  try {
    var ss    = getDb(params || {});
    var sheet = _ensureCT600Sheet(ss);
    var d     = params;

    // Check if draft for this period already exists
    var existingRow = -1;
    if (sheet.getLastRow() > 1) {
      var existing = sheet.getDataRange().getValues();
      for (var i = 1; i < existing.length; i++) {
        if (existing[i][1] === d.periodStart && existing[i][2] === d.periodEnd) {
          existingRow = i + 1;
          break;
        }
      }
    }

    var returnId = existingRow > 0
      ? sheet.getRange(existingRow, 1).getValue().toString()
      : generateId('CT6');

    var row = [
      returnId,
      d.periodStart          || '',
      d.periodEnd            || '',
      d.companyName          || '',
      d.companyNumber        || '',
      d.utr                  || '',
      d.box30_turnover       || 0,
      d.box35_tradingProfit  || 0,
      d.addBackDepreciation  || 0,
      d.lessCapitalAllowances|| 0,
      d.box45_adjProfit      || 0,
      d.box155_chargeable    || 0,
      d.box185_ctMainRate    || 0,
      d.box188_marginalRelief|| 0,
      d.box190_ctChargeable  || 0,
      d.box440_taxPayable    || 0,
      d.rateApplied          || '',
      d.effectiveRate        || 0,
      d.status               || 'Draft',
      d.calculatedAt         || new Date().toISOString(),
      d.notes                || ''
    ];

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    logAudit('SAVE', 'CT600', returnId, { period: d.periodStart + ' to ' + d.periodEnd });
    return { success: true, returnId: returnId, message: 'CT600 draft saved.' };
  } catch(e) {
    Logger.log('saveCT600Draft error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getCT600Returns(params) {
  try {
    var ss    = getDb(params || {});
    var sheet = ss.getSheetByName(CT600_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, returns: [] };

    var data    = sheet.getDataRange().getValues();
    var returns = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      returns.push({
        returnId:          data[i][0].toString(),
        periodStart:       safeSerializeDate(data[i][1]),
        periodEnd:         safeSerializeDate(data[i][2]),
        companyName:       data[i][3].toString(),
        companyNumber:     data[i][4].toString(),
        utr:               data[i][5].toString(),
        box30_turnover:    parseFloat(data[i][6])  || 0,
        box155_chargeable: parseFloat(data[i][11]) || 0,
        box440_taxPayable: parseFloat(data[i][15]) || 0,
        rateApplied:       data[i][16].toString(),
        effectiveRate:     parseFloat(data[i][17]) || 0,
        status:            data[i][18].toString(),
        calculatedAt:      safeSerializeDate(data[i][19]),
        notes:             data[i][20] ? data[i][20].toString() : ''
      });
    }
    returns.sort(function(a,b){ return b.periodEnd > a.periodEnd ? 1 : -1; });
    return { success: true, returns: returns };
  } catch(e) {
    return { success: false, message: e.toString(), returns: [] };
  }
}

// -----------------------------------------------------------------------------
// HELPERS
// -----------------------------------------------------------------------------

function _r(n) { return Math.round((parseFloat(n)||0) * 100) / 100; }

function _shiftYear(dateStr, years) {
  var d = new Date(dateStr);
  d.setFullYear(d.getFullYear() + years);
  return d.toISOString().split('T')[0];
}

function _ensureCT600Sheet(ss) {
  var sheet = ss.getSheetByName(CT600_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CT600_SHEET_NAME);
    sheet.getRange(1, 1, 1, 21).setValues([[
      'ReturnId','PeriodStart','PeriodEnd','CompanyName','CompanyNumber','UTR',
      'Box30_Turnover','Box35_TradingProfit','AddBackDepreciation','LessCapitalAllowances',
      'Box45_AdjProfit','Box155_Chargeable','Box185_CTMainRate','Box188_MarginalRelief',
      'Box190_CTChargeable','Box440_TaxPayable','RateApplied','EffectiveRate%',
      'Status','CalculatedAt','Notes'
    ]]);
  }
  return sheet;
}