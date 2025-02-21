// Configuration
const API_CONFIG = {
  BASE_URL: 'https://generativelanguage.googleapis.com/v1',
  DEFAULT_MODEL: 'models/gemini-2.0-flash',
  TIMEOUT: {
    FREE: 30000,    // 30 seconds for free tier
    PAID: 60000     // 60 seconds for paid tier
  },
  RATE_LIMIT: {
    FREE: {
      REQUESTS_PER_MINUTE: 15,
      REQUESTS_PER_DAY: 60
    },
    PAID: {
      REQUESTS_PER_MINUTE: 60,
      REQUESTS_PER_DAY: 1000
    }
  },
  RETRY: {
    MAX_ATTEMPTS: 3,
    BASE_DELAY: 1000
  },
  MODELS: {
    'models/gemini-2.0-flash': {
      name: 'Gemini 2.0 Flash',
      experimental: false,
      default: true
    },
    'models/gemini-2.0-flash-lite': {
      name: 'Gemini 2.0 Flash Lite',
      experimental: true
    },
    'models/gemini-1.5-pro': {
      name: 'Gemini 1.5 Pro',
      experimental: false
    }
  },
  LEGACY_MODEL_MAPPING: {
    'gemini-pro': 'models/gemini-2.0-flash',
    'gemini-pro-flash': 'models/gemini-2.0-flash'
  }
};

const CACHE_KEYS = {
  STATE: 'geminiState',
  RESPONSE_PREFIX: 'geminiResponse_'
};

// ---- State Management ----

/**
 * Manages the state of the Gemini add-on
 */
class GeminiState {
  constructor() {
    this.state = {
      apiKey: null,
      apiKeySet: false,
      lastKeyValidation: null,
      tier: 'FREE',
      defaultModel: API_CONFIG.DEFAULT_MODEL,
      experimental: false,
      showLegacy: false,
      history: [],
      availableModels: null,
      modelsLastUpdated: null
    };
    this.processing = new Map();
    this.loadState();
  }

  loadState() {
    try {
      const savedState = PropertiesService.getScriptProperties().getProperty(CACHE_KEYS.STATE);
      if (savedState) {
        const parsed = JSON.parse(savedState);
        this.state = { ...this.state, ...parsed };
      }
    } catch (error) {
      console.error('Failed to load state:', error);
    }
  }

  saveState() {
    try {
      PropertiesService.getScriptProperties().setProperty(
        CACHE_KEYS.STATE,
        JSON.stringify(this.state)
      );
    } catch (error) {
      console.error('Failed to save state:', error);
    }
  }

  getApiKey() {
    return this.state.apiKey;
  }

  setApiKey(key) {
    if (!key) return false;
    this.state.apiKey = key;
    this.state.apiKeySet = true;
    this.state.lastKeyValidation = Date.now();
    this.saveState();
    return true;
  }

  validateApiKey(apiKey) {
    if (!apiKey) return false;
    
    try {
      const response = UrlFetchApp.fetch(`${API_CONFIG.BASE_URL}/models`, {
        headers: { 'x-goog-api-key': apiKey },
        muteHttpExceptions: true,
        timeout: this.state.tier === 'PAID' ? API_CONFIG.TIMEOUT.PAID : API_CONFIG.TIMEOUT.FREE
      });
      
      if (response.getResponseCode() === 200) {
        // Update tier based on response headers or response data
        this.state.tier = 'PAID'; // You might want to check specific headers/data
        this.state.lastKeyValidation = Date.now();
        this.saveState();
        return true;
      }
      
      return false;
    } catch (error) {
      console.error('API key validation failed:', error);
      return false;
    }
  }

  updateProcessing(id, details) {
    this.processing.set(id, { ...this.processing.get(id), ...details });
  }

  isCompleted(id) {
    const request = this.processing.get(id);
    return request && request.status === 'complete';
  }

  addToHistory(entry) {
    this.state.history.push({
      ...entry,
      timestamp: Date.now()
    });
    if (this.state.history.length > 1000) {
      this.state.history = this.state.history.slice(-1000);
    }
    this.saveState();
  }

  getStats() {
    return {
      totalRequests: this.state.history.length,
      completedRequests: this.state.history.filter(h => h.status === 'complete').length,
      errorRequests: this.state.history.filter(h => h.status === 'error').length,
      averageTime: this.calculateAverageTime()
    };
  }

  calculateAverageTime() {
    const completed = this.state.history.filter(h => h.status === 'complete' && h.startTime && h.endTime);
    if (completed.length === 0) return 0;
    const total = completed.reduce((sum, h) => sum + (h.endTime - h.startTime), 0);
    return total / completed.length;
  }

  migrateModel(model) {
    if (!model) return API_CONFIG.DEFAULT_MODEL;
    
    // Check if it's a legacy model that needs migration
    if (API_CONFIG.LEGACY_MODEL_MAPPING[model]) {
      return API_CONFIG.LEGACY_MODEL_MAPPING[model];
    }
    
    // Check if it's a valid current model
    if (API_CONFIG.MODELS[model]) {
      return model;
    }
    
    // Default to the recommended model
    return API_CONFIG.DEFAULT_MODEL;
  }
}

// Initialize state after API_CONFIG is defined
var geminiState = new GeminiState();

// Add automatic trigger for onOpen
function createTriggers() {
  // Delete any existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create new trigger for onOpen
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
}

// Runs when the add-on is installed
function onInstall(e) {
  createTriggers();
  onOpen(e);
}

// Runs when the spreadsheet is opened
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Gemini')
    .addItem('🚀 Get Started', 'showSetupWizard')
    .addSeparator()
    .addItem('📊 Dashboard', 'createUnifiedSidebar')
    .addItem('⚙️ Settings', 'createSettingsSidebar')
    .addItem('❓ Help', 'showHelp')
    .addItem('🔄 Reset All', 'resetAllSettings')
    .addToUi();
  
  // Show setup wizard if not configured
  if (!PropertiesService.getUserProperties().getProperty('setupComplete')) {
    showSetupWizard();
  }
}

// Single RequestQueue declaration
const RequestQueue = {
  queue: [],
  processing: false,
  
  add: function(request) {
    // Initialize request status
    geminiState.updateProcessing(request.id, { 
      status: 'queued', 
      prompt: request.prompt,
      model: request.model,
      systemPrompt: request.systemPrompt,
      temperature: request.temperature,
      startTime: Date.now(),
      queuePosition: this.queue.length + 1
    });
    
    this.queue.push(request);
    
    // Start processing if not already running
    if (!this.processing) {
      this.processQueue();
    }
  },
  
  processQueue: function() {
    if (this.queue.length === 0) {
      this.processing = false;
      return;
    }
    
    this.processing = true;
    const config = geminiState.state.tier === 'PAID' ? 
      API_CONFIG.RATE_LIMIT.PAID : 
      API_CONFIG.RATE_LIMIT.FREE;
    
    const batch = this.queue.splice(0, config.REQUESTS_PER_MINUTE);
    
    batch.forEach(req => {
      try {
        // Update status to processing
        geminiState.updateProcessing(req.id, { 
          status: 'processing',
          startProcessing: Date.now()
        });
        
        // Use retryWithBackoffSync to retry API request if necessary
        const result = retryWithBackoffSync(() => req.process(), API_CONFIG.RETRY.MAX_ATTEMPTS);
        
        // Update status to complete
        geminiState.updateProcessing(req.id, { 
          status: 'complete', 
          response: result,
          endTime: Date.now()
        });
        
      } catch (error) {
        // Update status to error
        geminiState.updateProcessing(req.id, { 
          status: 'error',
          error: error.message,
          endTime: Date.now()
        });
      }
    });
    
    // Process next batch
    if (this.queue.length > 0) {
      Utilities.sleep(config.BASE_DELAY);
      this.processQueue();
    } else {
      this.processing = false;
    }
  }
};

/**
 * Generates text using the Gemini API.
 * 
 * This is a synchronous custom function that can be used directly in Google Sheets.
 * It handles API key validation, caching, queuing, and error propagation.
 * 
 * @param {string} prompt - The text prompt to send to Gemini
 * @param {string} [model=models/gemini-2.0-flash] - The model to use (e.g., gemini-2.0-flash, gemini-2.0-flash-lite, gemini-1.5-pro)
 * @param {string} [systemPrompt=""] - Optional system prompt to guide the model's behavior
 * @param {number} [temperature=0.7] - Controls randomness (0.0 = focused, 1.0 = creative)
 * @return {string} Generated text or an error message that can be displayed in the cell
 * @customfunction
 */
function GEMINI(prompt, model = API_CONFIG.DEFAULT_MODEL, systemPrompt = "", temperature = 0.7) {
  if (!prompt || !prompt.trim()) {
    return "⚠️ Error: Prompt is required";
  }
  
  try {
    // Check API key
    const apiKey = geminiState.getApiKey();
    if (!apiKey) {
      return "⚠️ API key not set. Please use the Gemini menu to configure your API key.";
    }
    
    // Fallback for generateFormulaId if not defined
    const id = (typeof generateFormulaId === 'function') ? generateFormulaId(prompt, model, systemPrompt, temperature) : (prompt + '|' + model + '|' + systemPrompt + '|' + temperature);
    
    // Check processing status
    if (geminiState.processing.has(id)) {
      const status = geminiState.processing.get(id).status;
      if (status === 'error') {
        const error = geminiState.processing.get(id).error;
        return `❌ Error: ${error}`;
      }
      return `⏳ Request in progress... (Status: ${status})`;
    }
    
    // Check if completed
    if (geminiState.isCompleted(id)) {
      const result = geminiState.processing.get(id).response;
      if (!result) {
        return "❌ Error: No response found";
      }
      return result;
    }
    
    // Check cache
    const formulaKey = getCacheKey(prompt, model, systemPrompt, temperature);
    const cachedResponse = getCachedResponse(formulaKey);
    if (cachedResponse) {
      return cachedResponse;
    }
    
    // Create new request
    const request = {
      id: id,
      prompt: prompt,
      model: model,
      systemPrompt: systemPrompt,
      temperature: temperature,
      process: function() {
        try {
          const response = makeGeminiRequest(this.prompt, this.model, this.systemPrompt, this.temperature);
          if (!response) {
            throw new Error('Empty response from API');
          }
          // Cache the successful response
          setCachedResponse(formulaKey, response);
          return response;
        } catch (error) {
          throw error;
        }
      }
    };
    
    // Add request to queue
    RequestQueue.add(request);
    
    // Return loading message with estimated wait
    const queuePosition = RequestQueue.queue.findIndex(req => req.id === id) + 1;
    const config = geminiState.state.tier === 'PAID' ? API_CONFIG.RATE_LIMIT.PAID : API_CONFIG.RATE_LIMIT.FREE;
    const estimatedWait = Math.ceil(queuePosition / config.REQUESTS_PER_MINUTE) * (config.BASE_DELAY / 1000);
    
    return `⏳ Loading... (#${queuePosition} in queue, ~${estimatedWait}s)`;
    
  } catch (error) {
    // Log error details for debugging
    logError(error, { function: 'GEMINI', prompt: prompt });
    return `❌ Error: ${error.message}`;
  }
}

// Utility Functions
function logError(error, context = {}) {
  console.error(JSON.stringify({
    timestamp: new Date().toISOString(),
    error: error.message || error,
    stack: error.stack,
    ...context
  }));
}

function retryWithBackoffSync(operation, maxAttempts = API_CONFIG.RETRY.MAX_ATTEMPTS) {
  let attempt = 1;
  let delay = API_CONFIG.RETRY.BASE_DELAY;
  while (attempt <= maxAttempts) {
    try {
      return operation();
    } catch (error) {
      if (attempt === maxAttempts || !isRetryableError(error)) {
        throw error;
      }
      Utilities.sleep(delay);
      delay = Math.min(delay * 2, API_CONFIG.RETRY.MAX_DELAY);
      attempt++;
    }
  }
}

function isRetryableError(error) {
  const retryableCodes = [408, 429, 500, 502, 503, 504];
  return (error.code && retryableCodes.includes(error.code)) || (error.message && error.message.includes('timeout'));
}

function getAvailableModelsSync() {
  const cache = CacheService.getUserCache();
  const cachedModels = cache.get(CACHE_KEYS.MODELS);
  const timestamp = cache.get(CACHE_KEYS.MODEL_TIMESTAMP);
  if (cachedModels && timestamp) {
    const age = Date.now() - parseInt(timestamp, 10);
    if (age < 360000) { // 1 hour
      return JSON.parse(cachedModels);
    }
  }
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) { throw new Error('API key not set'); }
    const response = UrlFetchApp.fetch(`${API_CONFIG.BASE_URL}?key=${apiKey}`, {
      method: 'GET',
      muteHttpExceptions: true,
      timeout: API_CONFIG.TIMEOUT
    });
    if (response.getResponseCode() !== 200) {
      throw new Error(`Failed to fetch models: ${response.getContentText()}`);
    }
    const result = JSON.parse(response.getContentText());
    const models = result.models || [];
    const filteredModels = models.filter(model => model.name.includes('gemini')).map(model => ({
      name: model.name.split('/').pop(),
      displayName: model.displayName,
      description: model.description,
      inputTokenLimit: model.inputTokenLimit,
      outputTokenLimit: model.outputTokenLimit,
      isExperimental: model.version === 'experimental',
      isRecommended: model.name.includes('2.0-flash')
    })).sort((a, b) => {
      if (a.isRecommended !== b.isRecommended) return b.isRecommended ? 1 : -1;
      if (a.isExperimental !== b.isExperimental) return a.isExperimental ? 1 : -1;
      return b.name.localeCompare(a.name);
    });
    cache.put(CACHE_KEYS.MODELS, JSON.stringify(filteredModels), 3600);
    cache.put(CACHE_KEYS.MODEL_TIMESTAMP, Date.now().toString(), 3600);
    return filteredModels;
  } catch (error) {
    logError(error, { function: 'getAvailableModelsSync' });
    return [API_CONFIG.DEFAULT_MODEL]; // Return default model if fetching fails
  }
}

// Add backward compatibility function
function fetchAvailableModels() {
  return getAvailableModelsSync();
}

// Add cache management functions
function getCacheKey(prompt, model, systemPrompt, temperature) {
  // Create a shorter hash by taking first 32 chars of input
  const input = [
    prompt.substring(0, 100),
    model,
    systemPrompt ? systemPrompt.substring(0, 50) : '',
    temperature
  ].join('|');
  
  // Use a shorter hash
  return Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      input
    )
  ).substring(0, 32);
}

function getCachedResponse(cacheKey) {
  const cache = CacheService.getScriptCache();
  const value = cache.get(cacheKey);
  if (!value) return null;
  
  const parsed = JSON.parse(value);
  if (parsed.type === 'chunked') {
    const chunks = cache.getAll(parsed.keys);
    if (!chunks || Object.keys(chunks).length !== parsed.length) return null;
    return parsed.keys.map(key => chunks[key]).join('');
  }
  
  return parsed.data;
}

function setCachedResponse(cacheKey, response) {
  const cache = CacheService.getScriptCache();
  // Split data into chunks if it's too large
  if (response.length > 100000) {
    const chunks = [];
    for (let i = 0; i < response.length; i += 100000) {
      chunks.push(response.slice(i, i + 100000));
    }
    const chunkKeys = chunks.map((_, i) => `${cacheKey}_${i}`);
    cache.put(cacheKey, JSON.stringify({
      type: 'chunked',
      keys: chunkKeys,
      length: chunks.length
    }), 21600);
    
    chunks.forEach((chunk, i) => {
      cache.put(chunkKeys[i], chunk, 21600);
    });
  } else {
    cache.put(cacheKey, JSON.stringify({
      type: 'single',
      data: response
    }), 21600);
  }
}

// UI Functions
function createSettingsSidebar() {
  const template = HtmlService.createTemplate(SETTINGS_HTML);
  const currentSettings = {
    apiKey: geminiState.getApiKey(),
    apiKeySet: !!geminiState.getApiKey(),
    defaultModel: geminiState.state.defaultModel,
    experimental: geminiState.state.experimental,
    showLegacy: geminiState.state.showLegacy,
    models: (geminiState.state.availableModels || []).map(model => ({
      name: model,
      info: API_CONFIG.MODELS[model]
    }))
  };
  template.settings = currentSettings;
  const output = template.evaluate().setTitle('Gemini Settings').setWidth(400);
  SpreadsheetApp.getUi().showSidebar(output);
}

function createUnifiedSidebar() {
  const template = HtmlService.createTemplate(DASHBOARD_HTML);
  template.page = 'dashboard';
  template.settings = {
    apiKeySet: !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
    experimental: geminiState.state.experimental,
    showLegacy: geminiState.state.showLegacy,
    defaultModel: geminiState.state.defaultModel,
    models: Object.entries(API_CONFIG.MODELS)
      .filter(([, info]) => {
        if (info.isExperimental && !geminiState.state.experimental) return false;
        if (info.isLegacy && !geminiState.state.showLegacy) return false;
        return true;
      })
      .map(([id, info]) => ({ name: id, info: info }))
  };
  const htmlOutput = template.evaluate().setTitle('Gemini Dashboard').setWidth(300);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function createWelcomeFlow() {
  const template = HtmlService.createTemplate(WELCOME_HTML);
  template.hasApiKey = !!geminiState.state.apiKey;
  const ui = template.evaluate().setWidth(600).setHeight(400).setTitle('Welcome to Gemini for Sheets');
  SpreadsheetApp.getUi().showModalDialog(ui, 'Welcome to Gemini for Sheets');
}

function onSetupApiKey(apiKey) {
  try {
    // First try to set the API key
    setApiKey(apiKey);
    
    // Test the API key by listing models
    const testResponse = UrlFetchApp.fetch(`${API_CONFIG.BASE_URL}/models`, {
      method: 'GET',
      headers: { 
        'Content-Type': 'application/json',
        'x-goog-api-key': apiKey
      },
      muteHttpExceptions: true
    });
    
    if (testResponse.getResponseCode() !== 200) {
      throw new Error('Invalid API key');
    }
    
    // NEW: Fetch available models and update state
    geminiState.state.availableModels = getAvailableModelsSync();
    geminiState.saveState();
    
    // If we got here, the API key is valid
    return { success: true };
  } catch (error) {
    // Clean up if validation failed
    PropertiesService.getScriptProperties().deleteProperty('GEMINI_API_KEY');
    geminiState.state.apiKey = null;
    geminiState.state.apiKeySet = false;
    geminiState.saveState();
    return { success: false, error: error.message };
  }
}

function setApiKey(key) {
  if (!key) { throw new Error('API key is required'); }
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  geminiState.state.apiKey = key;
  geminiState.state.apiKeySet = true;
  geminiState.state.lastKeyValidation = Date.now();
  geminiState.saveState();
  return true;
}

// Additional UI and Helper Functions
function showSetupWizard() {
  const template = HtmlService.createTemplate(WELCOME_HTML);
  const html = template.evaluate()
    .setTitle('Welcome to Gemini')
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Welcome to Gemini');
}

function showHelp() {
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(
    HtmlService.createHtmlOutput(`
    <style>
    body { font-family: Arial, sans-serif; margin: 16px; color: #202124; }
    h2 { margin-top: 0; }
    .section { margin-bottom: 24px; }
    .code { background: #f8f9fa; padding: 8px; border-radius: 4px; font-family: monospace; }
    .tip { color: #137333; }
    </style>
    <h2>Gemini for Sheets Help</h2>
    <div class="section">
    <h3>Quick Start</h3>
    <ol>
    <li>Get an API key from <a href="https://aistudio.google.com/app/apikey" class="api-link">Google AI Studio</a></li>
    <li>Click the button below to set up your API key</li>
    <li>Start using the GEMINI() formula in your spreadsheet!</li>
    </ol>
    </div>
    <div class="section">
    <h3>Formula Usage</h3>
    <p>The GEMINI formula has these optional parameters:</p>
    <ul>
    <li><strong>prompt</strong>: Your main prompt (required)</li>
    <li><strong>model</strong>: Specific model to use (optional)</li>
    <li><strong>systemPrompt</strong>: Additional context/instructions (optional)</li>
    <li><strong>temperature</strong>: Creativity level (0-1, default 0.7)</li>
    </ul>
    <div class="code">=GEMINI("Write a story", "models/gemini-2.0-flash", "Make it funny", 0.8)</div>
    </div>
    <div class="section">
    <h3>Tips</h3>
    <ul>
    <li class="tip">Use "models/gemini-2.0-flash" for fastest responses</li>
    <li class="tip">Convert formulas to text to preserve responses</li>
    <li class="tip">Monitor processing status in the dashboard</li>
    <li class="tip">Check cell notes for detailed error messages</li>
    </ul>
    </div>
    <div class="section">
    <h3>Rate Limits</h3>
    <ul>
    <li>Free Tier: 30 requests/minute</li>
    <li>Paid Tier: 2,000 requests/minute</li>
    </ul>
    <p>Change your tier in Settings if you have a paid API key.</p>
    `).setWidth(500).setHeight(600),
    'Help & Support'
  );
}

function resetAllSettings() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Reset All Settings',
    'This will clear all settings including your API key. Are you sure?',
    ui.ButtonSet.YES_NO
  );
  if (response === ui.Button.YES) {
    geminiState.resetState();
    ui.alert('Settings Reset', 'All settings have been reset to default. Please use "Get Started" to set up your API key.', ui.ButtonSet.OK);
    showSetupWizard();
  }
}

function runTests() {
  const results = { passed: 0, failed: 0, errors: [] };
  function assert(condition, message) {
    if (!condition) { results.failed++; results.errors.push(message); throw new Error(message); }
    results.passed++;
  }
  function runTestSuite() {
    try {
      // Test GeminiState initialization
      assert(geminiState instanceof GeminiState, 'GeminiState not initialized properly');
      assert(typeof geminiState.state === 'object', 'State object not initialized');
      
      // Test default model setting
      const expectedModel = API_CONFIG.DEFAULT_MODEL;
      const actualModel = geminiState.state.defaultModel;
      assert(actualModel === expectedModel, `Default model not set correctly. Expected: ${expectedModel}, Got: ${actualModel}`);
      
      // Test model configuration
      assert(API_CONFIG.MODELS['models/gemini-2.0-flash'], 'Default model not in API config');
      assert(API_CONFIG.LEGACY_MODEL_MAPPING['gemini-pro'], 'Legacy model mapping missing');
      
      // Test model migration
      const migratedModel = geminiState.migrateModel('gemini-pro');
      assert(migratedModel === 'models/gemini-2.0-flash', 'Model migration failed');
      
      // Test API key validation
      const mockApiKey = 'test_api_key';
      const mockResponse = {
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({
          quotaInfo: { isPaid: false }
        })
      };
      
      // Mock UrlFetchApp
      const originalUrlFetchApp = UrlFetchApp;
      UrlFetchApp = {
        fetch: (url, options) => mockResponse
      };
      
      const validationResult = geminiState.validateApiKey(mockApiKey);
      assert(validationResult === true, 'API key validation failed');
      assert(geminiState.state.tier === 'FREE', 'Account tier not set correctly');
      
      // Restore original UrlFetchApp
      UrlFetchApp = originalUrlFetchApp;
      
      // Test formula parsing
      const testFormula = '=GEMINI("test", "models/gemini-2.0-flash", "", 0.7)';
      const [prompt, model, systemPrompt, temp] = (function(formula) {
        const params = formula.substring(8, formula.length - 1).split(',');
        return [
          params[0]?.trim().replace(/^"|"$/g, '') || '',
          params[1]?.trim().replace(/^"|"$/g, '') || '',
          params[2]?.trim().replace(/^"|"$/g, '') || '',
          Number(params[3]?.trim()) || 0.7
        ];
      })(testFormula);
      
      assert(prompt === 'test', 'Formula parsing failed - prompt');
      assert(model === 'models/gemini-2.0-flash', 'Formula parsing failed - model');
      assert(systemPrompt === '', 'Formula parsing failed - systemPrompt');
      assert(temp === 0.7, 'Formula parsing failed - temperature');
      
      return { 
        ...results, 
        message: `All tests passed! (${results.passed} tests)` 
      };
    } catch (error) {
      return { 
        ...results, 
        message: `Tests failed: ${error.message}\nPassed: ${results.passed}, Failed: ${results.failed}`,
        error: error.toString(),
        stack: error.stack 
      };
    }
  }
  const testResults = runTestSuite();
  const ui = SpreadsheetApp.getUi();
  if (testResults.failed === 0) {
    ui.alert('✅ Tests Passed', `All ${testResults.passed} tests passed successfully!\n\nVerified:\n- State Management\n- API Configuration\n- Model Migration\n- API Integration\n- Formula Parsing`, ui.ButtonSet.OK);
  } else {
    ui.alert('❌ Tests Failed', `${testResults.failed} test(s) failed, ${testResults.passed} passed.\n\nErrors:\n${testResults.errors.join('\n')}`, ui.ButtonSet.OK);
  }
  console.log('Test Results:', testResults);
  return testResults;
}

function getSettings() {
  return {
    apiKeySet: !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'),
    experimental: geminiState.state.experimental,
    showLegacy: geminiState.state.showLegacy,
    defaultModel: geminiState.state.defaultModel,
    availableModels: (geminiState.state.availableModels || []).map(model => ({
      name: model,
      displayName: API_CONFIG.MODELS[model]?.name || model
    }))
  };
}

function updateSettings(settings) {
  geminiState.state.experimental = settings.experimental;
  geminiState.state.showLegacy = settings.showLegacy;
  geminiState.state.defaultModel = settings.defaultModel;
  geminiState.saveState();
  return true;
}

const SETTINGS_HTML = `
<!DOCTYPE html>
<html>
  <head>
    <base target="_blank">
    <style>
      body { font-family: Arial, sans-serif; margin: 10px; }
      .section { margin-bottom: 20px; }
      .header { font-weight: bold; margin-bottom: 10px; }
      .field { margin-bottom: 15px; }
      .status { margin-top: 5px; color: #666; }
      .error { color: #d32f2f; }
      .success { color: #388e3c; }
      button { 
        background-color: #1a73e8;
        color: white;
        border: none;
        padding: 8px 16px;
        border-radius: 4px;
        cursor: pointer;
      }
      button:hover { background-color: #1557b0; }
      button:disabled {
        background-color: #ccc;
        cursor: not-allowed;
      }
      select { 
        padding: 6px;
        border-radius: 4px;
        border: 1px solid #ccc;
      }
      .checkbox-label {
        display: flex;
        align-items: center;
        gap: 8px;
      }
      #apiKeySection { margin-bottom: 30px; }
      #modelSection { 
        display: none;
        opacity: 0;
        transition: opacity 0.3s ease-in-out;
      }
      #modelSection.visible {
        display: block;
        opacity: 1;
      }
    </style>
  </head>
  <body>
    <div id="apiKeySection" class="section">
      <div class="header">🔑 API Key Configuration</div>
      <div class="field">
        <input type="password" id="apiKey" placeholder="Enter your API key" style="width: 100%; padding: 6px;">
        <div id="apiKeyStatus" class="status"></div>
      </div>
      <button id="saveKeyBtn" onclick="saveApiKey()">Save API Key</button>
    </div>
    
    <div id="modelSection" class="section">
      <div class="header">⚙️ Model Settings</div>
      <div class="field">
        <label>Default Model:</label><br>
        <select id="defaultModel" style="width: 100%"></select>
      </div>
      <div class="field">
        <label class="checkbox-label">
          <input type="checkbox" id="experimental">
          Enable Experimental Models
        </label>
      </div>
      <div class="field">
        <label class="checkbox-label">
          <input type="checkbox" id="showLegacy">
          Show Legacy Models
        </label>
      </div>
      <button onclick="saveSettings()">Save Settings</button>
    </div>

    <script>
      let isProcessing = false;

      document.getElementById('apiKey').addEventListener('input', function() {
        const keyLength = this.value.trim().length;
        document.getElementById('saveKeyBtn').disabled = keyLength === 0;
        document.getElementById('apiKeyStatus').textContent = '';
      });

      function saveApiKey() {
        if (isProcessing) return;
        isProcessing = true;

        const apiKey = document.getElementById('apiKey').value.trim();
        const statusDiv = document.getElementById('apiKeyStatus');
        const saveBtn = document.getElementById('saveKeyBtn');
        
        statusDiv.className = 'status';
        statusDiv.textContent = 'Validating API key...';
        saveBtn.disabled = true;
        
        google.script.run
          .withSuccessHandler(function(result) {
            isProcessing = false;
            saveBtn.disabled = false;
            
            if (result.success) {
              statusDiv.className = 'status success';
              statusDiv.textContent = '✓ API key validated successfully';
              
              // Show model settings
              const modelSection = document.getElementById('modelSection');
              modelSection.style.display = 'block';
              setTimeout(() => modelSection.classList.add('visible'), 50);
              
              // Load available models
              loadSettings();
            } else {
              statusDiv.className = 'status error';
              statusDiv.textContent = '✗ ' + (result.error || 'Failed to validate API key');
              document.getElementById('modelSection').style.display = 'none';
            }
          })
          .withFailureHandler(function(error) {
            isProcessing = false;
            saveBtn.disabled = false;
            statusDiv.className = 'status error';
            statusDiv.textContent = '✗ ' + error.message;
            document.getElementById('modelSection').style.display = 'none';
          })
          .onSetupApiKey(apiKey);
      }

      function saveSettings() {
        const settings = {
          defaultModel: document.getElementById('defaultModel').value,
          experimental: document.getElementById('experimental').checked,
          showLegacy: document.getElementById('showLegacy').checked
        };
        
        google.script.run
          .withSuccessHandler(function() {
            const statusDiv = document.createElement('div');
            statusDiv.className = 'status success';
            statusDiv.textContent = '✓ Settings saved';
            document.querySelector('#modelSection button').after(statusDiv);
            setTimeout(() => statusDiv.remove(), 3000);
          })
          .updateSettings(settings);
      }

      function loadSettings() {
        google.script.run
          .withSuccessHandler(function(settings) {
            if (settings.apiKeySet) {
              document.getElementById('apiKeyStatus').className = 'status success';
              document.getElementById('apiKeyStatus').textContent = '✓ API key is set';
              document.getElementById('modelSection').style.display = 'block';
              setTimeout(() => document.getElementById('modelSection').classList.add('visible'), 50);
            }
            
            const modelSelect = document.getElementById('defaultModel');
            modelSelect.innerHTML = '';
            settings.availableModels.forEach(function(model) {
              const option = document.createElement('option');
              option.value = model.name;
              option.text = model.displayName;
              option.selected = model.name === settings.defaultModel;
              modelSelect.appendChild(option);
            });

            document.getElementById('experimental').checked = settings.experimental;
            document.getElementById('showLegacy').checked = settings.showLegacy;
          })
          .getSettings();
      }

      // Initial load
      document.getElementById('saveKeyBtn').disabled = true;
      if (document.getElementById('apiKey').value.trim()) {
        loadSettings();
      }
    </script>
  </body>
</html>
`;

const DASHBOARD_HTML = 
'<!DOCTYPE html>' +
'<html>' +
'  <head>' +
'    <base target="_blank">' +
'    <style>' +
'      body { font-family: Arial, sans-serif; margin: 10px; }' +
'      .section { margin-bottom: 20px; }' +
'      .header { font-weight: bold; margin-bottom: 10px; }' +
'      .stats { ' +
'        display: grid;' +
'        grid-template-columns: repeat(2, 1fr);' +
'        gap: 10px;' +
'        margin-bottom: 15px;' +
'      }' +
'      .stat-card {' +
'        background: #f8f9fa;' +
'        padding: 10px;' +
'        border-radius: 4px;' +
'      }' +
'      .stat-value { ' +
'        font-size: 24px;' +
'        color: #1a73e8;' +
'      }' +
'      .stat-label { ' +
'        font-size: 12px;' +
'        color: #666;' +
'      }' +
'      .processing-list {' +
'        max-height: 200px;' +
'        overflow-y: auto;' +
'        border: 1px solid #eee;' +
'        padding: 10px;' +
'        border-radius: 4px;' +
'      }' +
'      .processing-item {' +
'        padding: 8px;' +
'        border-bottom: 1px solid #eee;' +
'      }' +
'      .processing-item:last-child {' +
'        border-bottom: none;' +
'      }' +
'      .status-running { color: #fb8c00; }' +
'      .status-complete { color: #388e3c; }' +
'      .status-error { color: #d32f2f; }' +
'    </style>' +
'  </head>' +
'  <body>' +
'    <div class="section">' +
'      <div class="header">📊 Usage Statistics</div>' +
'      <div class="stats">' +
'        <div class="stat-card">' +
'          <div class="stat-value" id="totalRequests">-</div>' +
'          <div class="stat-label">Total Requests</div>' +
'        </div>' +
'        <div class="stat-card">' +
'          <div class="stat-value" id="avgTime">-</div>' +
'          <div class="stat-label">Avg Response Time</div>' +
'        </div>' +
'        <div class="stat-card">' +
'          <div class="stat-value" id="errorRate">-</div>' +
'          <div class="stat-label">Error Rate</div>' +
'        </div>' +
'        <div class="stat-card">' +
'          <div class="stat-value" id="activeModel">-</div>' +
'          <div class="stat-label">Active Model</div>' +
'        </div>' +
'      </div>' +
'    </div>' +
'    <div class="section">' +
'      <div class="header">🔄 Processing Queue</div>' +
'      <div id="processingList" class="processing-list">' +
'        <div class="processing-item">No active requests</div>' +
'      </div>' +
'    </div>' +
'    <script>' +
'      function updateStats() {' +
'        google.script.run' +
'          .withSuccessHandler(function(stats) {' +
'            document.getElementById("totalRequests").textContent = stats.totalRequests;' +
'            document.getElementById("avgTime").textContent = stats.avgResponseTime + "ms";' +
'            document.getElementById("errorRate").textContent = stats.errorRate + "%";' +
'            document.getElementById("activeModel").textContent = stats.activeModel;' +
'            const list = document.getElementById("processingList");' +
'            list.innerHTML = "";' +
'            if (stats.processing.length === 0) {' +
'              list.innerHTML = "<div class=\\"processing-item\\">No active requests</div>";' +
'              return;' +
'            }' +
'            stats.processing.forEach(function(item) {' +
'              const div = document.createElement("div");' +
'              div.className = "processing-item";' +
'              const status = item.status === "running" ? "⏳" : ' +
'                           item.status === "complete" ? "✓" : "✗";' +
'              div.innerHTML = "<div class=\\"status-" + item.status + "\\">" +' +
'                            status + " " + item.prompt.substring(0, 50) + "..." +' +
'                            "<small>(" + item.duration + "ms)</small></div>";' +
'              list.appendChild(div);' +
'            });' +
'          })' +
'          .getStats();' +
'      }' +
'      setInterval(updateStats, 2000);' +
'      updateStats();' +
'    </script>' +
'  </body>' +
'</html>';

const WELCOME_HTML = `
<!DOCTYPE html>
<html>
  <head>
    <base target="_blank">
    <style>
      body { 
        font-family: Arial, sans-serif; 
        margin: 20px;
        line-height: 1.6;
      }
      .welcome-card {
        background: #f8f9fa;
        padding: 20px;
        border-radius: 8px;
        margin-bottom: 20px;
      }
      .title {
        font-size: 24px;
        color: #1a73e8;
        margin-bottom: 10px;
      }
      .subtitle {
        color: #666;
        margin-bottom: 20px;
      }
      .step {
        margin-bottom: 15px;
        padding-left: 24px;
        position: relative;
      }
      .step:before {
        content: "→";
        position: absolute;
        left: 0;
        color: #1a73e8;
      }
      button { 
        background-color: #1a73e8;
        color: white;
        border: none;
        padding: 10px 20px;
        border-radius: 4px;
        cursor: pointer;
        font-size: 14px;
      }
      button:hover { background-color: #1557b0; }
      .api-link {
        color: #1a73e8;
        text-decoration: none;
      }
      .api-link:hover { text-decoration: underline; }
    </style>
  </head>
  <body>
    <div class="welcome-card">
      <div class="title">Welcome to Gemini for Google Sheets! 🚀</div>
      <div class="subtitle">Let's get you set up with Google's most advanced AI models.</div>
      
      <div class="step">
        Visit <a href="https://aistudio.google.com/app/apikey" class="api-link">Google AI Studio</a> to get your API key
      </div>
      <div class="step">
        Click the button below to set up your API key
      </div>
      <div class="step">
        Start using the GEMINI() formula in your spreadsheet!
      </div>
      
      <button onclick="showSettings()">Set Up API Key</button>
    </div>

    <script>
      function showSettings() {
        google.script.host.close();
        google.script.run.createSettingsSidebar();
      }
    </script>
  </body>
</html>
`;
