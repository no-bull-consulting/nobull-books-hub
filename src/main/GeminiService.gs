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
  var userMessage = params.userMessage || 'Give me a financial health summary of this business.';
  var reportData  = params.reportData  || {};
  var history     = Array.isArray(params.history) ? params.history : [];

  // ── System prompt ───────────────────────────────────────────────────────────
  var systemPrompt =
    'You are the no~bull books AI Financial Assistant for a UK business. ' +
    'You have been given real-time financial context from the user\'s accounting system below. ' +
    'Answer the user\'s question clearly, concisely and practically. ' +
    'Use British English spelling. Format your response with Markdown ' +
    '(use **bold** for key figures, bullet points for lists, ## for section headings). ' +
    'Focus on actionable insight. If the data is insufficient to answer, say so clearly. ' +
    'Do not invent figures.\n\n' +
    'FINANCIAL CONTEXT:\n' + JSON.stringify(reportData, null, 2);

  // ── Build contents array (conversation history + new message) ───────────────
  var contents = [];

  // Add history turns (capped at last 10 exchanges = 20 entries)
  var recentHistory = history.slice(-20);
  recentHistory.forEach(function(h) {
    contents.push({
      role:  h.role === 'model' ? 'model' : 'user',
      parts: [{ text: h.text }]
    });
  });

  // Add current user message (with system context injected on first turn only)
  var messageText = history.length === 0
    ? systemPrompt + '\n\nUser question: ' + userMessage
    : userMessage;

  contents.push({ role: 'user', parts: [{ text: messageText }] });

  // ── Call Gemini API ─────────────────────────────────────────────────────────
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;

  var payload = {
    contents: contents,
    generationConfig: {
      temperature:     0.7,
      maxOutputTokens: 1024,
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
