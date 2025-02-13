// Script Properties Keys
const PROPERTIES = {
  API_KEY: 'GEMINI_API_KEY',
  RATE_LIMIT_TOKENS: 'RATE_LIMIT_TOKENS',
  RATE_LIMIT_LAST_REFILL: 'RATE_LIMIT_LAST_REFILL',
  AUTO_CONVERT_TO_VALUES: 'AUTO_CONVERT_TO_VALUES'
};

// Cache duration in seconds (6 hours)
const CACHE_DURATION = 21600;

// Rate limiting configuration
const RATE_LIMIT = {
  MAX_TOKENS: 60,  // Maximum tokens
  REFILL_RATE: 60, // Tokens added per minute
  TOKENS_PER_REQUEST: 1 // Tokens used per request
};

// Store API key in Script Properties
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Enter your Gemini API Key');
  if (response.getSelectedButton() === ui.Button.OK) {
    PropertiesService.getScriptProperties().setProperty(PROPERTIES.API_KEY, response.getResponseText());
    ui.alert('API Key saved!');
  }
}

// Toggle auto-conversion setting
function toggleAutoConvert() {
  const props = PropertiesService.getScriptProperties();
  const currentSetting = props.getProperty(PROPERTIES.AUTO_CONVERT_TO_VALUES) === 'true';
  props.setProperty(PROPERTIES.AUTO_CONVERT_TO_VALUES, (!currentSetting).toString());
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(`Auto-conversion of formulas to values is now ${!currentSetting ? 'enabled' : 'disabled'}`);
}

// Rate limiter implementation
function getRateLimitTokens() {
  const props = PropertiesService.getScriptProperties();
  let tokens = Number(props.getProperty(PROPERTIES.RATE_LIMIT_TOKENS) || RATE_LIMIT.MAX_TOKENS);
  const lastRefill = Number(props.getProperty(PROPERTIES.RATE_LIMIT_LAST_REFILL) || Date.now());
  
  // Calculate tokens to add based on time passed
  const now = Date.now();
  const minutesPassed = (now - lastRefill) / (1000 * 60);
  const tokensToAdd = Math.floor(minutesPassed * RATE_LIMIT.REFILL_RATE);
  
  if (tokensToAdd > 0) {
    tokens = Math.min(RATE_LIMIT.MAX_TOKENS, tokens + tokensToAdd);
    props.setProperty(PROPERTIES.RATE_LIMIT_TOKENS, tokens.toString());
    props.setProperty(PROPERTIES.RATE_LIMIT_LAST_REFILL, now.toString());
  }
  
  return tokens;
}

function useRateLimitTokens(count) {
  const props = PropertiesService.getScriptProperties();
  const currentTokens = getRateLimitTokens();
  
  if (currentTokens < count) {
    return false;
  }
  
  props.setProperty(PROPERTIES.RATE_LIMIT_TOKENS, (currentTokens - count).toString());
  return true;
}

// Main GEMINI function
function GEMINI(prompt, systemPrompt = "", temperature = 0.7, autoConvert = null) {
  // Check if prompt is empty
  if (!prompt) return "Error: Prompt is required";
  
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
  
  // Check rate limit
  if (!useRateLimitTokens(RATE_LIMIT.TOKENS_PER_REQUEST)) {
    return "Rate limit exceeded. Please try again later.";
  }
  
  const apiKey = PropertiesService.getScriptProperties().getProperty(PROPERTIES.API_KEY);
  if (!apiKey) return "Please set API key first using the Gemini menu";
  
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
    
    return generatedText;
  } catch (error) {
    return "Error: " + error.toString();
  }
}

// Convert formulas to values in selected range
function convertFormulasToValues() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();
  range.setValues(values);
}

// Add menu to spreadsheet
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Gemini')
    .addItem('Set API Key', 'setApiKey')
    .addItem('Toggle Auto-Convert to Values', 'toggleAutoConvert')
    .addItem('Convert Selected Formulas to Values', 'convertFormulasToValues')
    .addToUi();
}
