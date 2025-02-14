// Configuration
const RATE_LIMIT = {
  REQUESTS_PER_MIN: 300,
  DELAY_MS: 50,
  BATCH_SIZE: 20
};

// Cache for API key
let cachedApiKey;

function setup() {
  SpreadsheetApp.getUi()
    .createMenu('Gemini')
    .addItem('First Time Setup', 'firstTimeSetup')
    .addItem('Set/Update API Key', 'setApiKey')
    .addItem('Convert Formulas to Text', 'convertFormulasToValues')
    .addItem('Test Response Time', 'testGeminiLatency')
    .addToUi();
    
  SpreadsheetApp.getUi().alert(
    'Setup Started\n\n' +
    '1. Refresh sheet\n' +
    '2. Click "Gemini" menu\n' +
    '3. Click "First Time Setup"\n' +
    '4. Get API key from aistudio.google.com'
  );
}

function firstTimeSetup() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Get API Key\n\n' +
    '1. Go to aistudio.google.com\n' +
    '2. Click "Get API Key"\n' +
    '3. Create new key\n' +
    '4. Copy key\n' +
    '5. Click "Set/Update API Key" here'
  );
}

function getApiKey() {
  if (!cachedApiKey) {
    cachedApiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  }
  return cachedApiKey;
}

function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'API Key Setup',
    'Paste API Key from aistudio.google.com:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    
    try {
      const testUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${apiKey}`;
      const testStart = Date.now();
      const testResponse = UrlFetchApp.fetch(testUrl, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{ parts: [{ text: "test" }] }]
        }),
        muteHttpExceptions: true
      });
      const testDuration = Date.now() - testStart;

      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', apiKey);
      cachedApiKey = apiKey;
      ui.alert('Success', 
        `API key working!\nTest response time: ${testDuration}ms\n\nTry: =GEMINI("Write a thank you email")`, 
        ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('Error', 
        'Invalid API key. Check and try again.', 
        ui.ButtonSet.OK);
    }
  }
}

function processWithRateLimit() {
  const cache = CacheService.getScriptCache();
  const now = Date.now();
  const windowKey = Math.floor(now / 60000); // 1-minute window
  const requestCount = Number(cache.get(`requests_${windowKey}`) || 0);
  
  if (requestCount >= RATE_LIMIT.REQUESTS_PER_MIN) {
    Utilities.sleep(RATE_LIMIT.DELAY_MS);
  }
  
  cache.put(`requests_${windowKey}`, requestCount + 1, 60);
}

/**
 * @param {string} prompt The prompt to send to Gemini
 * @param {string=} systemPrompt Optional instructions
 * @param {number=} temperature Optional creativity (0-1)
 * @customfunction
 */
function GEMINI(prompt, systemPrompt = "", temperature = 0.7) {
  const timing = {
    start: Date.now()
  };

  if (!prompt) return "⚠️ Enter prompt";
  
  const apiKey = getApiKey();
  if (!apiKey) return "⚠️ Set API key";

  timing.afterGetKey = Date.now();

  try {
    processWithRateLimit();
    timing.afterRateLimit = Date.now();

    const finalPrompt = systemPrompt ? `${systemPrompt}\n${prompt}` : prompt;
    const options = {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: finalPrompt }] }],
        generationConfig: { temperature: temperature }
      }),
      muteHttpExceptions: true,
      timeout: 30
    };

    const response = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=' + apiKey,
      options
    );
    timing.afterFetch = Date.now();

    if (response.getResponseCode() === 429) {
      Utilities.sleep(100);
      return GEMINI(prompt, systemPrompt, temperature);
    }

    const result = JSON.parse(response.getContentText());
    if (!result.candidates?.[0]?.content?.parts?.[0]?.text) {
      throw new Error("Invalid response");
    }
    
    timing.end = Date.now();
    Logger.log(`
      Get Key: ${timing.afterGetKey - timing.start}ms
      Rate Limit: ${timing.afterRateLimit - timing.afterGetKey}ms
      API Call: ${timing.afterFetch - timing.afterRateLimit}ms
      Parse: ${timing.end - timing.afterFetch}ms
      Total: ${timing.end - timing.start}ms
    `);
    
    return result.candidates[0].content.parts[0].text;

  } catch (error) {
    const errorMsg = error.toString().toLowerCase();
    if (errorMsg.includes("429")) return "⚠️ Rate limit - try again";
    if (errorMsg.includes("timeout")) return "⚠️ Timeout - try again";
    if (errorMsg.includes("invalid")) return "⚠️ API key invalid";
    return "⚠️ Error - try again";
  }
}

function testGeminiLatency() {
  const ui = SpreadsheetApp.getUi();
  const startTime = Date.now();
  
  const result = GEMINI("Hello world");
  
  const endTime = Date.now();
  const latency = endTime - startTime;
  
  Logger.log(`Response time: ${latency}ms`);
  Logger.log(`Response: ${result}`);
  
  ui.alert('Latency Test', 
    `Response time: ${latency}ms\n\nResponse: ${result}`, 
    ui.ButtonSet.OK);
  
  return latency;
}

function convertFormulasToValues() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  
  if (!range) {
    ui.alert('Error', 'Please select cells first.', ui.ButtonSet.OK);
    return;
  }

  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();
  
  if (numRows * numCols > 1000) {
    const response = ui.alert(
      'Large Selection',
      'Selected ' + (numRows * numCols) + ' cells. Continue?',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  }

  try {
    const BATCH_SIZE = 100;
    for (let startRow = 0; startRow < numRows; startRow += BATCH_SIZE) {
      const batchRows = Math.min(BATCH_SIZE, numRows - startRow);
      const batchRange = range.offset(startRow, 0, batchRows, numCols);
      const values = batchRange.getValues();
      batchRange.setValues(values);
      
      if (batchRows === BATCH_SIZE) {
        SpreadsheetApp.flush();
        Utilities.sleep(100);
      }
    }

    ui.alert(
      'Success',
      'Converted ' + (numRows * numCols) + ' cells.',
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert(
      'Error',
      'Failed: ' + error.toString(),
      ui.ButtonSet.OK
    );
  }
}

function onOpen() {
  setup();
}
