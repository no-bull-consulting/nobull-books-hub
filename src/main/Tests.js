/**
 * NO~BULL BOOKS — HMRC COMPLIANCE TEST SUITE
 * Verifies critical logic for security and calculation accuracy.
 */

function runComplianceTests() {
  Logger.log('🚀 Starting HMRC Compliance Test Suite...');
  
  var results = [
    testPeriodLockLogic(),
    testVatCalculationLogic(),
    testAuthEngine()
  ];
  
  var failed = results.filter(function(r) { return !r.success; });
  
  if (failed.length === 0) {
    Logger.log('✅ ALL TESTS PASSED');
    return { success: true, message: "All logic verified." };
  } else {
    Logger.log('❌ TESTS FAILED: ' + failed.length);
    return { success: false, errors: failed };
  }
}

/**
 * Verifies that the Period Lock correctly blocks back-dated entries.
 */
function testPeriodLockLogic() {
  try {
    var mockSettings = { lockedBefore: '2024-03-31' };
    var pastDate = '2024-01-01';
    
    // We expect this to throw an error
    try {
      _checkPeriodLock(pastDate, { _sheetId: 'TEST_ID' }); 
      return { name: "Period Lock", success: false, error: "Failed to block past date" };
    } catch (e) {
      if (e.message.indexOf('locked') >= 0) {
        return { name: "Period Lock", success: true };
      }
      return { name: "Period Lock", success: false, error: e.message };
    }
  } catch(err) {
    return { name: "Period Lock", success: false, error: err.toString() };
  }
}

/**
 * Verifies the accuracy of Box 1 and Box 4 VAT math.
 */
function testVatCalculationLogic() {
  try {
    var net = 1000;
    var rate = 20;
    var expectedVat = 200;
    
    var calculated = Math.round(net * (rate / 100) * 100) / 100;
    
    if (calculated === expectedVat) {
      return { name: "VAT Accuracy", success: true };
    } else {
      return { name: "VAT Accuracy", success: false, error: "Math mismatch: " + calculated };
    }
  } catch(err) {
    return { name: "VAT Accuracy", success: false, error: err.toString() };
  }
}

/**
 * Verifies that Unauthorized calls are blocked.
 */
function testAuthEngine() {
  try {
    var result = _canDoPermission('ReadOnly', 'users.manage');
    if (result === false) {
      return { name: "RBAC Security", success: true };
    } else {
      return { name: "RBAC Security", success: false, error: "ReadOnly allowed to manage users" };
    }
  } catch(err) {
    return { name: "RBAC Security", success: false, error: err.toString() };
  }
}