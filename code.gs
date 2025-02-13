// STEP 1: COPY ALL THIS CODE
// STEP 2: PASTE INTO APPS SCRIPT
// STEP 3: CLICK SAVE AND RUN setup()

//=============== SETUP AND MENU ===============
function setup() {
  // Create menu
  SpreadsheetApp.getUi()
    .createMenu('🤖 Gemini')
    .addItem('✨ First Time Setup - Click Here', 'firstTimeSetup')
    .addItem('🔑 Set/Update API Key', 'setApiKey')
    .addItem('📝 Convert Formulas to Text', 'convertFormulasToValues')
    .addToUi();
    
  // Show welcome message
  SpreadsheetApp.getUi().alert(
    '✨ Almost there! ✨\n\n' +
    '1. Refresh your sheet\n' +
    '2. Click the new "🤖 Gemini" menu above\n' +
    '3. Click "✨ First Time Setup"\n' +
    '4. Get your free API key from aistudio.google.com'
  );
}

//=============== FIRST TIME SETUP ===============
function firstTimeSetup() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    '🔑 Get Your Free API Key\n\n' +
    '1. Go to aistudio.google.com\n' +
    '2. Click "Get API Key" (top right)\n' +
    '3. Create a new key (it\'s free!)\n' +
    '4. Copy the key\n' +
    '5. Come back here and click "🔑 Set/Update API Key"'
  );
}

//=============== API KEY MANAGEMENT ===============
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '🔑 API Key Setup',
    'Paste your API Key from aistudio.google.com:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    
    // Test the API key
    const testUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${apiKey}`;
    try {
      UrlFetchApp.fetch(testUrl, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{ parts: [{ text: "test" }] }]
        })
      });

      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', apiKey);
      ui.alert('✅ Success!', 
        'Your API key is working!\n\nTry typing this in any cell:\n=GEMINI("Write a thank you email")', 
        ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('❌ Error', 
        'That API key didn\'t work.\nPlease check you copied it correctly from aistudio.google.com', 
        ui.ButtonSet.OK);
    }
  }
}

//=============== MAIN GEMINI FUNCTION ===============
/**
 * Makes an AI request to Gemini
 * @param {string} prompt The prompt to send to Gemini
 * @param {string=} systemPrompt Optional additional instructions
 * @param {number=} temperature Optional creativity level (0.0 to 1.0)
 * @customfunction
 */
function GEMINI(prompt, systemPrompt = "", temperature = 0.7) {
  // Input validation
  if (!prompt) return "⚠️ Please enter a question or prompt";
  
  // Get API key
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return "⚠️ Click '🤖 Gemini' menu above → '✨ First Time Setup'";

  try {
    // Combine prompts if system prompt is provided
    const finalPrompt = systemPrompt ? `${systemPrompt}\n\nUser request: ${prompt}` : prompt;

    // Make API request
    const response = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=' + apiKey,
      {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [{ parts: [{ text: finalPrompt }] }],
          generationConfig: { temperature: temperature }
        })
      }
    );
    
    // Parse and return response
    const result = JSON.parse(response.getContentText());
    return result.candidates[0].content.parts[0].text;
  } catch (error) {
    return "❌ Error: " + error.toString();
  }
}

//=============== UTILITY FUNCTIONS ===============
function convertFormulasToValues() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getValues();
  range.setValues(values);
  
  SpreadsheetApp.getUi().alert(
    '✅ Done!\n\nThe selected cells have been converted from formulas to plain text.'
  );
}

// Add this to make sure menu appears on open
function onOpen() {
  setup();
}
