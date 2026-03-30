/**
 * NO~BULL BOOKS — REPORTS
 * P&L, balance sheet, VAT return, transactions, cash flow, currency breakdown
 *
 * KEY FIX: params threaded through every function so getDb(params) can
 * target the correct client spreadsheet in the hub model.
 * ─────────────────────────────────────────────────────────────────────────────
 */

function generateProfitLoss(startDate, endDate, params) {
  try {
    var txnResult = getAllTransactions(startDate, endDate, params);
    if (!txnResult.success) return txnResult;

    var coaResult = getAccounts({}, params);
    if (!coaResult.success) return coaResult;

    var revenue      = {};
    var expenses     = {};
    var totalRevenue = 0;
    var totalExpenses= 0;

    var accountMap = {};
    coaResult.accounts.forEach(function(a) { accountMap[a.accountCode] = a; });

    txnResult.transactions.forEach(function(t) {
      var amount = t.amount;

      if (accountMap[t.accountCredit] && accountMap[t.accountCredit].accountType === 'Revenue') {
        if (!revenue[t.accountCredit]) {
          revenue[t.accountCredit] = { name: accountMap[t.accountCredit].accountName, amount: 0 };
        }
        revenue[t.accountCredit].amount += amount;
        totalRevenue += amount;
      }

      if (accountMap[t.accountDebit] && accountMap[t.accountDebit].accountType === 'Expense') {
        if (!expenses[t.accountDebit]) {
          expenses[t.accountDebit] = { name: accountMap[t.accountDebit].accountName, amount: 0 };
        }
        expenses[t.accountDebit].amount += amount;
        totalExpenses += amount;
      }
    });

    return {
      success: true,
      report: {
        revenue:       revenue,
        expenses:      expenses,
        totalRevenue:  totalRevenue,
        totalExpenses: totalExpenses,
        netProfit:     totalRevenue - totalExpenses
      }
    };
  } catch(e) {
    Logger.log('generateProfitLoss error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function generateBalanceSheet(params) {
  try {
    var coaResult = getAccounts({}, params);
    if (!coaResult.success) return coaResult;

    var assets      = {};
    var liabilities = {};
    var equity      = {};

    coaResult.accounts.forEach(function(acc) {
      if (!acc.active) return;
      var balance = parseFloat(acc.currentBalance) || 0;
      if (Math.abs(balance) < 0.01) balance = 0;

      if (acc.accountType === 'Asset') {
        assets[acc.accountCode]      = { name: acc.accountName, amount: balance };
      } else if (acc.accountType === 'Liability') {
        liabilities[acc.accountCode] = { name: acc.accountName, amount: balance };
      } else if (acc.accountType === 'Equity') {
        equity[acc.accountCode]      = { name: acc.accountName, amount: balance };
      }
    });

    var totalAssets      = Object.keys(assets).reduce(function(s,k)     { return s + assets[k].amount;      }, 0);
    var totalLiabilities = Object.keys(liabilities).reduce(function(s,k){ return s + liabilities[k].amount; }, 0);
    var totalEquity      = Object.keys(equity).reduce(function(s,k)      { return s + equity[k].amount;      }, 0);

    return {
      success: true,
      report: {
        assets:           assets,
        liabilities:      liabilities,
        equity:           equity,
        totalAssets:      totalAssets,
        totalLiabilities: totalLiabilities,
        totalEquity:      totalEquity
      }
    };
  } catch(e) {
    Logger.log('generateBalanceSheet error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function generateVATReturn(startDate, endDate, params) {
  try {
    var start = new Date(startDate);
    var end   = new Date(endDate);

    var invoicesResult = getAllInvoices(params);
    var billsResult    = getAllBills(params);

    if (!invoicesResult.success || !billsResult.success) {
      return { success: false, message: 'Failed to load invoice/bill data' };
    }

    var outputVAT      = 0;
    var inputVAT       = 0;
    var totalSales     = 0;
    var totalPurchases = 0;

    // Only Approved/Sent/Paid/Partial — Draft (Pro-Forma) excluded from VAT
    var vatableStatuses = ['Approved', 'Sent', 'Paid', 'Partial', 'Overdue'];
    invoicesResult.invoices.forEach(function(inv) {
      var issueDate = new Date(inv.issueDate);
      if (issueDate >= start && issueDate <= end && vatableStatuses.indexOf(inv.status) >= 0) {
        outputVAT  += parseFloat(inv.vatAmount) || parseFloat(inv.vat) || parseFloat(inv.vatTotal) || 0;
        totalSales += parseFloat(inv.subtotal)  || 0;
      }
    });

    billsResult.bills.forEach(function(bill) {
      var issueDate = new Date(bill.issueDate);
      if (issueDate >= start && issueDate <= end && bill.status !== 'Cancelled' && bill.status !== 'Void') {
        inputVAT       += parseFloat(bill.vatTotal) || parseFloat(bill.vat) || 0;
        totalPurchases += parseFloat(bill.subtotal) || 0;
      }
    });

    var netVAT = outputVAT - inputVAT;

    return {
      success: true,
      data: {
        periodStart:    startDate,
        periodEnd:      endDate,
        totalSales:     totalSales,
        totalPurchases: totalPurchases,
        outputVAT:      outputVAT,
        inputVAT:       inputVAT,
        netVAT:         netVAT,
        isPayable:      netVAT > 0,
        amountDue:      Math.abs(netVAT)
      }
    };
  } catch(e) {
    Logger.log('generateVATReturn error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getAllTransactions(startDate, endDate, params) {
  try {
    var sheet = getDb(params || {}).getSheetByName(SHEETS.TRANSACTIONS);
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, transactions: [] };
    }

    var data = sheet.getDataRange().getValues();
    var transactions = [];

    var start = startDate ? new Date(startDate) : null;
    var end   = endDate   ? new Date(endDate)   : null;
    if (end) end.setHours(23, 59, 59, 999);

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;

      var txnDate = row[1] ? new Date(row[1]) : null;
      if (start && txnDate && txnDate < start) continue;
      if (end   && txnDate && txnDate > end)   continue;

      transactions.push({
        transactionId: row[0].toString(),
        date:          safeSerializeDate(row[1]),
        type:          row[2]  ? row[2].toString()  : '',
        reference:     row[3]  ? row[3].toString()  : '',
        accountDebit:  row[4]  ? row[4].toString()  : '',
        accountCredit: row[5]  ? row[5].toString()  : '',
        amount:        parseFloat(row[6]) || 0,
        description:   row[7]  ? row[7].toString()  : '',
        invoiceId:     row[8]  ? row[8].toString()  : '',
        billId:        row[9]  ? row[9].toString()  : '',
        reconciled:    row[10] === true || row[10] === 'TRUE'
      });
    }

    return { success: true, transactions: transactions };
  } catch(e) {
    Logger.log('getAllTransactions error: ' + e.toString());
    return { success: false, message: e.toString(), transactions: [] };
  }
}

function generateCashFlow(startDate, endDate, params) {
  try {
    var plResult = generateProfitLoss(startDate, endDate, params);
    if (!plResult.success) return plResult;

    var txResult = getAllTransactions(startDate, endDate, params);
    if (!txResult.success) return txResult;

    var pl   = plResult.report || {};
    var txns = txResult.transactions || [];

    var revenue   = Object.keys(pl.revenue  || {}).reduce(function(s,k){ return s + (pl.revenue[k].amount  || 0); }, 0);
    var expenses2 = Object.keys(pl.expenses || {}).reduce(function(s,k){ return s + (pl.expenses[k].amount || 0); }, 0);
    var netProfit = revenue - expenses2;

    // ── Operating activities ───────────────────────────────────────────────
    var depreciation = txns
      .filter(function(t){ return t.accountDebit && String(t.accountDebit).match(/^81/); })
      .reduce(function(s,t){ return s + t.amount; }, 0);

    var debtorMovement =
      txns.filter(function(t){ return t.accountDebit === '1100'; })
          .reduce(function(s,t){ return s + t.amount; }, 0) -
      txns.filter(function(t){ return t.accountCredit === '1100'; })
          .reduce(function(s,t){ return s + t.amount; }, 0);

    var creditorMovement =
      txns.filter(function(t){ return t.accountCredit === '2100'; })
          .reduce(function(s,t){ return s + t.amount; }, 0) -
      txns.filter(function(t){ return t.accountDebit === '2100'; })
          .reduce(function(s,t){ return s + t.amount; }, 0);

    var operating = netProfit + depreciation - debtorMovement + creditorMovement;

    // ── Investing activities ───────────────────────────────────────────────
    var assetPurchases = txns
      .filter(function(t){ return t.accountDebit && String(t.accountDebit).match(/^0/); })
      .reduce(function(s,t){ return s + t.amount; }, 0);

    var assetDisposals = txns
      .filter(function(t){ return t.accountCredit && String(t.accountCredit).match(/^0/); })
      .reduce(function(s,t){ return s + t.amount; }, 0);

    var investing = assetDisposals - assetPurchases;

    // ── Financing activities ───────────────────────────────────────────────
    var drawings = txns
      .filter(function(t){ return t.accountDebit && String(t.accountDebit).match(/^3/) && t.type !== 'YearEndClose'; })
      .reduce(function(s,t){ return s + t.amount; }, 0);

    var capitalIn = txns
      .filter(function(t){ return t.accountCredit && String(t.accountCredit).match(/^3/) && t.type !== 'YearEndClose'; })
      .reduce(function(s,t){ return s + t.amount; }, 0);

    var financing = capitalIn - drawings;
    var netCash   = operating + investing + financing;

    return {
      success:   true,
      from:      startDate,
      to:        endDate,
      netProfit: netProfit,
      operating: {
        netProfit:        netProfit,
        depreciation:     depreciation,
        debtorMovement:   -debtorMovement,
        creditorMovement: creditorMovement,
        total:            operating
      },
      investing: {
        assetPurchases: -assetPurchases,
        assetDisposals:  assetDisposals,
        total:           investing
      },
      financing: {
        capitalIn:  capitalIn,
        drawings:   -drawings,
        total:      financing
      },
      netCash: netCash
    };
  } catch(e) {
    Logger.log('generateCashFlow error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function getCurrencyBreakdown(startDate, endDate, params) {
  try {
    var settings = getSettings(params);
    var baseCurr = settings.baseCurrency || 'GBP';
    var ss       = getDb(params || {});
    var start    = startDate ? new Date(startDate) : null;
    var end      = endDate   ? new Date(endDate)   : null;
    if (end) end.setHours(23, 59, 59, 999);

    var result = {};

    // ── Invoices ────────────────────────────────────────────────────────────
    var invSheet = ss.getSheetByName(SHEETS.INVOICES);
    if (invSheet && invSheet.getLastRow() > 1) {
      var invData = invSheet.getDataRange().getValues();
      for (var i = 1; i < invData.length; i++) {
        var row    = invData[i];
        var status = (row[14] || '').toString();
        if (status === 'Draft' || status === 'Void' || status === 'Voided') continue;
        var issueDate = row[6] ? new Date(row[6]) : null;
        if (start && issueDate && issueDate < start) continue;
        if (end   && issueDate && issueDate > end)   continue;
        var curr      = (row[19] || baseCurr).toString() || baseCurr;
        var total     = parseFloat(row[11]) || 0;
        var baseTotal = parseFloat(row[21]) || total;
        var amtDue    = parseFloat(row[13]) || 0;
        if (!result[curr]) result[curr] = {
          currency: curr, isBase: curr === baseCurr,
          invoices: { count:0, total:0, baseTotal:0, outstanding:0 },
          bills:    { count:0, total:0, baseTotal:0, outstanding:0 }
        };
        result[curr].invoices.count++;
        result[curr].invoices.total      += total;
        result[curr].invoices.baseTotal  += baseTotal;
        result[curr].invoices.outstanding+= amtDue;
      }
    }

    // ── Bills ───────────────────────────────────────────────────────────────
    var bilSheet = ss.getSheetByName(SHEETS.BILLS);
    if (bilSheet && bilSheet.getLastRow() > 1) {
      var bilData = bilSheet.getDataRange().getValues();
      for (var j = 1; j < bilData.length; j++) {
        var brow    = bilData[j];
        var bstatus = (brow[12] || '').toString();
        if (bstatus === 'Void' || bstatus === 'Voided') continue;
        var bdate = brow[4] ? new Date(brow[4]) : null;
        if (start && bdate && bdate < start) continue;
        if (end   && bdate && bdate > end)   continue;
        var bcurr    = (brow[16] || baseCurr).toString() || baseCurr;
        var btotal   = parseFloat(brow[9])  || 0;
        var bbase    = parseFloat(brow[18]) || btotal;
        var bdue     = parseFloat(brow[11]) || 0;
        if (!result[bcurr]) result[bcurr] = {
          currency: bcurr, isBase: bcurr === baseCurr,
          invoices: { count:0, total:0, baseTotal:0, outstanding:0 },
          bills:    { count:0, total:0, baseTotal:0, outstanding:0 }
        };
        result[bcurr].bills.count++;
        result[bcurr].bills.total      += btotal;
        result[bcurr].bills.baseTotal  += bbase;
        result[bcurr].bills.outstanding+= bdue;
      }
    }

    // Round all totals
    Object.keys(result).forEach(function(curr) {
      var r = result[curr];
      ['invoices', 'bills'].forEach(function(type) {
        r[type].total       = Math.round(r[type].total       * 100) / 100;
        r[type].baseTotal   = Math.round(r[type].baseTotal   * 100) / 100;
        r[type].outstanding = Math.round(r[type].outstanding * 100) / 100;
      });
    });

    return {
      success:      true,
      baseCurrency: baseCurr,
      breakdown:    Object.values(result).sort(function(a, b) {
        if (a.isBase) return -1;
        if (b.isBase) return  1;
        return a.currency.localeCompare(b.currency);
      }),
      period: { start: startDate || '', end: endDate || '' }
    };
  } catch(e) {
    Logger.log('getCurrencyBreakdown error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

function exportReport(reportType, fromDate, toDate, format, params) {
  try {
    var ss        = getDb(params || {});
    var tempSheet = ss.insertSheet('Temp_Report_' + new Date().getTime());
    var data      = [];

    if (reportType === 'pnl') {
      var report = generateProfitLoss(fromDate, toDate, params);
      if (!report.success) throw new Error(report.message);
      data.push(['PROFIT & LOSS STATEMENT', '']);
      data.push(['Period:', fromDate + ' to ' + toDate]);
      data.push(['', '']);
      data.push(['REVENUE', '']);
      Object.keys(report.report.revenue).forEach(function(acc) {
        data.push([report.report.revenue[acc].name, report.report.revenue[acc].amount]);
      });
      data.push(['Total Revenue', report.report.totalRevenue]);
      data.push(['', '']);
      data.push(['EXPENSES', '']);
      Object.keys(report.report.expenses).forEach(function(acc) {
        data.push([report.report.expenses[acc].name, report.report.expenses[acc].amount]);
      });
      data.push(['Total Expenses', report.report.totalExpenses]);
      data.push(['', '']);
      data.push(['NET PROFIT', report.report.netProfit]);
    }

    if (data.length > 0) {
      tempSheet.getRange(1, 1, data.length, 2).setValues(data);
    }

    if (format === 'pdf') {
      var ssId = ss.getId();
      var url  = 'https://docs.google.com/spreadsheets/d/' + ssId +
                 '/export?format=pdf&gid=' + tempSheet.getSheetId() +
                 '&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false';
      var response = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
      });
      var blob   = response.getBlob().setName(reportType + '_report_' + fromDate + '.pdf');
      var folder = DriveApp.getRootFolder();
      var file   = folder.createFile(blob);
      ss.deleteSheet(tempSheet);
      return { success: true, url: file.getUrl() };
    }

    ss.deleteSheet(tempSheet);
    return { success: false, message: 'Unsupported format: ' + format };
  } catch(e) {
    Logger.log('exportReport error: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}
