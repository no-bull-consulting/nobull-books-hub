/**
 * NO~BULL BOOKS — GEMINI AI SERVICE
 *
 * Handles all AI assistant requests from the frontend.
 * API key stored in Script Properties as: GEMINI_API_KEY
 *
 * Supports:
 *  - params.userMessage  : the user's question (string)
 *  - params.reportData   : financial context object built by _buildGeminiContext()
 *  - params.history      : array of {role, text} for multi-turn conversation
 */
function handleGeminiRequest(params, ctx) {

  // ── Permission check ────────────────────────────────────────────────────────
  if (!ctx || !ctx.canDo('reports.read')) {
    return { success: false, error: 'Unauthorized — reports.read permission required.' };
  }

  // ── API key ─────────────────────────────────────────────────────────────────
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return {
      success: false,
      error: 'Gemini API key not configured. Run setGeminiKey() in the Apps Script editor.'
    };
  }

  // ── Inputs ──────────────────────────────────────────────────────────────────
  var userMessage = params.question || params.userMessage || 'Give me a financial health summary of this business.';
  var reportData  = params.reportData  || {};
  var contextStr  = params.context || '';
  var history     = Array.isArray(params.history) ? params.history : [];

  // ── System prompt ───────────────────────────────────────────────────────────
  var systemPrompt =
    'You are the no~bull books AI Financial Assistant — a friendly, knowledgeable accounting assistant ' +
    'built into the no~bull books accounting system for UK small businesses and sole traders. ' +
    '\n\nYou have two roles:\n' +
    '1. FINANCIAL ADVISOR: Answer questions about the user\'s actual financial data (invoices, bills, cash position, VAT, reports).\n' +
    '2. APP GUIDE: Explain how to use any feature of no~bull books (invoicing, banking, reconciliation, VAT/MTD, reports, settings).\n' +
    '\nIMPORTANT RULES:\n' +
    '- Always use GBP (£) for all monetary figures. Never use USD or any other currency unless the user\'s data shows a foreign currency invoice.\n' +
    '- Use British English spelling throughout.\n' +
    '- Be warm, friendly and practical. Avoid jargon where possible.\n' +
    '- Format responses clearly: use **bold** for key figures, bullet points for lists.\n' +
    '- When you have real financial data, always reference it specifically. Never say you don\'t have data when the context below contains it.\n' +
    '- If the context shows £0 for something, say so honestly rather than pretending you have no data.\n' +
    '- Keep answers concise but complete. For how-to questions, give step-by-step guidance.\n' +
    '\nTHE USER\'S LIVE FINANCIAL DATA:\n' + (contextStr || JSON.stringify(reportData, null, 2)) +
    '\n\nIf the financial data above contains figures, use them directly in your answers. ' +
    'The currency is GBP (£) unless stated otherwise.';

  // ── Build contents array (conversation history + new message) ───────────────
  var contents = [];

  // Add history turns (capped at last 10 exchanges = 20 entries)
  var recentHistory = history.slice(-20);
  recentHistory.forEach(function(h) {
    var txt = h.text || h.content || '';
    if (!txt) return; // skip empty entries
    contents.push({
      role:  h.role === 'model' || h.role === 'assistant' ? 'model' : 'user',
      parts: [{ text: txt }]
    });
  });

  // Always inject system context with the user message
  // Context is compact enough to include every time for accuracy
  var messageText = systemPrompt + '\n\nUser question: ' + userMessage;

  contents.push({ role: 'user', parts: [{ text: messageText }] });

  // ── Call Gemini API ─────────────────────────────────────────────────────────
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;

  var payload = {
    contents: contents,
    generationConfig: {
      temperature:     0.7,
      maxOutputTokens: 2048,
      topP:            0.9
    },
    safetySettings: [
      { category: 'HARM_CATEGORY_HARASSMENT',        threshold: 'BLOCK_MEDIUM_AND_ABOVE' },
      { category: 'HARM_CATEGORY_HATE_SPEECH',       threshold: 'BLOCK_MEDIUM_AND_ABOVE' },
      { category: 'HARM_CATEGORY_SEXUALLY_EXPLICIT', threshold: 'BLOCK_MEDIUM_AND_ABOVE' },
      { category: 'HARM_CATEGORY_DANGEROUS_CONTENT', threshold: 'BLOCK_MEDIUM_AND_ABOVE' }
    ]
  };

  try {
    var response = UrlFetchApp.fetch(url, {
      method:           'post',
      contentType:      'application/json',
      payload:          JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var json = JSON.parse(response.getContentText());

    if (json.error) {
      Logger.log('Gemini API error: ' + JSON.stringify(json.error));
      throw new Error(json.error.message || 'API error');
    }

    if (!json.candidates || !json.candidates.length) {
      throw new Error('No response generated — the prompt may have been blocked by safety filters.');
    }

    var text = json.candidates[0].content.parts[0].text;
    return { success: true, analysis: text };

  } catch(e) {
    Logger.log('handleGeminiRequest error: ' + e.toString());
    return { success: false, error: 'AI Error: ' + e.toString() };
  }
}

/**
 * setGeminiKey(key)
 * Run once from the Apps Script editor to store your Gemini API key.
 * The key is stored in Script Properties and never written to the spreadsheet.
 *
 * Usage: setGeminiKey('AIza...');
 */
function setGeminiKey(key) {
  if (!key) { Logger.log('No key provided'); return; }
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  Logger.log('Gemini API key saved to Script Properties.');
}

/**
 * testGemini()
 * Run from Apps Script editor to verify the key works.
 */
function testGemini() {
  var ctx = { canDo: function() { return true; } };
  var result = handleGeminiRequest({
    userMessage: 'Say hello in one sentence.',
    reportData:  { businessName: 'Test Business', cashPosition: 5000 },
    history:     []
  }, ctx);
  Logger.log(JSON.stringify(result));
}
