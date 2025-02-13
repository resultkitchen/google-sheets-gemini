// Script Properties Keys
const PROPERTIES = {
  API_KEY: 'GEMINI_API_KEY',
  RATE_LIMIT_TOKENS: 'RATE_LIMIT_TOKENS',
  RATE_LIMIT_LAST_REFILL: 'RATE_LIMIT_LAST_REFILL',
  AUTO_CONVERT_TO_VALUES: 'AUTO_CONVERT_TO_VALUES',
  QUEUE_IN_PROGRESS: 'QUEUE_IN_PROGRESS',
  LAST_REQUEST_TIME: 'LAST_REQUEST_TIME'
};

// Cache duration in seconds (6 hours)
const CACHE_DURATION = 21600;

// Rate limiting configuration
const RATE_LIMIT = {
  MAX_TOKENS: 60,  // Maximum tokens
  REFILL_RATE: 60, // Tokens added per minute
  TOKENS_PER_REQUEST: 1 // Tokens used per request
};

function createTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    ScriptApp.newTrigger('onOpen')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onOpen()
      .create();
  }
}

function onOpen() {
  createTrigger(); // This ensures the trigger stays set
  SpreadsheetApp.getUi()
    .createMenu('Gemini')
    .addItem('Set API Key', 'setApiKey')
    .addItem('Toggle Auto-Convert to Values', 'toggleAutoConvert')
    .addItem('Convert Selected Formulas to Values', 'convertFormulasToValues')
    .addToUi();
}

function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Gemini API Key Setup',
    'Enter your Gemini API Key (it will be securely stored):',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    
    // Validate API key with a test request
    const testUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${apiKey}`;
    try {
      const testPayload = {
        contents: [{
          parts: [{
            text: "test"
          }]
        }]
      };
      
      UrlFetchApp.fetch(testUrl, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(testPayload)
      });

      PropertiesService.getScriptProperties().setProperty(PROPERTIES.API_KEY, apiKey);
      ui.alert('Success', 'API Key has been saved and verified!', ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('Error', 'Invalid API key. Please check and try again.', ui.ButtonSet.OK);
    }
  }
}

function toggleAutoConvert() {
  const props = PropertiesService.getScriptProperties();
  const currentSetting = props.getProperty(PROPERTIES.AUTO_CONVERT_TO_VALUES) === 'true';
  props.setProperty(PROPERTIES.AUTO_CONVERT_TO_VALUES, (!currentSetting).toString());
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(`Auto-conversion of formulas to values is now ${!currentSetting ? 'enabled' : 'disabled'}`);
}

function processQueue() {
  const props = PropertiesService.getScriptProperties();
  const sheet = SpreadsheetApp.getActiveSheet();
  const now = Date.now();
  const lastRequestTime = Number(props.getProperty(PROPERTIES.LAST_REQUEST_TIME) || 0);
  const timeToWait = Math.max(0, 1000 - (now - lastRequestTime)); // Ensure 1 second between requests

  if (timeToWait > 0) {
    Utilities.sleep(timeToWait);
  }

  // Update last request time
  props.setProperty(PROPERTIES.LAST_REQUEST_TIME, Date.now().toString());
}

function updateStatus(message) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange('A1'); // Or wherever you want the status
  range.setNote(message);
}

function GEMINI(prompt, systemPrompt = "", temperature = 0.7, autoConvert = null) {
  // Check if prompt is empty
  if (!prompt) return "Please enter a question or prompt in the formula";
  
  // Generate cache key
  const cacheKey = Utilities.base64Encode(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      prompt + systemPrompt + temperature
    )
  );
  
  // Check cache first
  const cache = CacheService.getScriptCache();
  const cachedResult = cache.get(cacheKey);
  if (cachedResult) return cachedResult;
  
  const apiKey = PropertiesService.getScriptProperties().getProperty(PROPERTIES.API_KEY);
  if (!apiKey) return "⚠️ API Key needed! Click the 'Gemini' menu above and select 'Set API Key'";

  // Process queue before making request
  processQueue();
  updateStatus('Processing request...');
  
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=' + apiKey;
  
  const payload = {
    contents: [{
      parts: [{
        text: systemPrompt ? `${systemPrompt}\n${prompt}` : prompt
      }]
    }],
    generationConfig: {
      temperature: temperature
    }
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    });
    
    const result = JSON.parse(response.getContentText());
    const generatedText = result.candidates[0].content.parts[0].text;
    
    // Cache the result
    cache.put(cacheKey, generatedText, CACHE_DURATION);
    
    // Handle auto-conversion if enabled
    const props = PropertiesService.getScriptProperties();
    const shouldAutoConvert = autoConvert ?? (props.getProperty(PROPERTIES.AUTO_CONVERT_TO_VALUES) === 'true');
    
    if (shouldAutoConvert) {
      const sheet = SpreadsheetApp.getActiveSheet();
      const activeRange = sheet.getActiveRange();
      activeRange.setValue(generatedText);
    }
    
    updateStatus('Done!');
    return generatedText;
  } catch (error) {
    if (error.toString().includes("API key")) {
      return "❌ API Key error. Please check if your key is valid in the Gemini menu.";
    }
    return "❌ Error: " + error.toString() + ". Try refreshing the page or checking your API key.";
  }
}

function convertFormulasToValues() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();
  range.setValues(values);
}

function install() {
  createTrigger();
  const ui = SpreadsheetApp.getUi();
  ui.alert('Installation Complete', 
    'The Gemini integration has been installed. You should now see a "Gemini" menu at the top. ' +
    'Please set your API key through that menu to get started.', 
    ui.ButtonSet.OK);
}
