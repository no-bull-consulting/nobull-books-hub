/**
 * NO~BULL BOOKS -- VALIDATION MIDDLEWARE
 * Strictly types and sanitizes all incoming API parameters.
 */
var Validators = {
  // Common sanitizers
  cleanString: function(val, maxLen) {
    return String(val || '').replace(/<[^>]*>?/gm, '').substring(0, maxLen || 255).trim();
  },
  
  cleanEmail: function(val) {
    var email = String(val || '').toLowerCase().trim();
    var re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return re.test(email) ? email : '';
  },

  cleanAmount: function(val) {
    var num = parseFloat(val);
    return isNaN(num) ? 0 : Math.round(num * 100) / 100;
  },

  // Schema Definitions
  schemas: {
    'createClient': function(p) {
      return {
        clientName:   Validators.cleanString(p.clientName || p.name, 100),
        email:        Validators.cleanEmail(p.email),
        phone:        Validators.cleanString(p.phone, 20).replace(/[^\d+ ]/g, ''),
        address:      Validators.cleanString(p.address, 500),
        postcode:     Validators.cleanString(p.postcode, 10).toUpperCase(),
        vatRegNumber: Validators.cleanString(p.vatRegNumber, 20),
        notes:        Validators.cleanString(p.notes, 1000)
      };
    },
    
    'createInvoice': function(p) {
      if (!p.clientId) throw new Error("Missing Client ID");
      return {
        clientId:     Validators.cleanString(p.clientId, 50),
        issueDate:    Validators.cleanString(p.issueDate, 10),
        dueDate:      Validators.cleanString(p.dueDate, 10),
        notes:        Validators.cleanString(p.notes, 1000),
        currency:     /^[A-Z]{3}$/.test(p.currency) ? p.currency : 'GBP',
        exchangeRate: parseFloat(p.exchangeRate) || 1.0,
        lines: (p.lines || []).map(function(l) {
          return {
            description: Validators.cleanString(l.description, 255),
            quantity:    parseFloat(l.quantity) || 1,
            unitPrice:   Validators.cleanAmount(l.unitPrice),
            vatRate:     parseFloat(l.vatRate) || 0,
            accountCode: Validators.cleanString(l.accountCode, 10)
          };
        })
      };
    }
  }
};

/**
 * Global Validate function to be called by Api.gs
 */
function validateParams(action, params) {
  if (Validators.schemas[action]) {
    var cleaned = Validators.schemas[action](params);
    // Merge back the technical sheet ID which isn't part of business logic validation
    cleaned._sheetId = params._sheetId;
    cleaned._verifiedEmail = params._verifiedEmail;
    return cleaned;
  }
  return params; // Fallback for actions without specific schemas
}