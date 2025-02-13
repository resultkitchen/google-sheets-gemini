# google-sheets-gemini
Want a Google Gemini formula in your sheet? This is a SIMPLE, forever free script that you can paste in and run with it.

# Gemini Sheets Integration

A Google Apps Script that integrates Google's Gemini AI with Google Sheets, allowing you to use Gemini's capabilities directly in your spreadsheet formulas.

## Features

* Custom GEMINI() formula for direct AI queries in cells
* Rate limiting to prevent API exhaustion (60 requests per minute)
* Response caching (6 hours) for efficiency
* Auto-conversion of formulas to values (optional)
* Simple menu interface for configuration

## Setup

1. Open your Google Sheet
2. Go to Extensions > Apps Script
3. Copy the code into Code.gs
4. Save the script
5. Refresh your spreadsheet
6. Get your Gemini API key from Google AI Studio
7. Use the new "Gemini" menu to set your API key

## Usage

Basic formula:
```
=GEMINI("Your prompt here")
```

Advanced usage:
```
=GEMINI("Your prompt", "Optional system prompt", 0.7, true)
```

### Parameters

* prompt: Your main prompt (required)
* systemPrompt: Additional context or instructions (optional)
* temperature: Controls randomness (0.0-1.0, default 0.7)
* autoConvert: Convert formula to value automatically (true/false)

## Example Use Cases

* Content Generation: =GEMINI("Write a product description for: " & A1)
* Translation: =GEMINI("Translate to Spanish: " & B1)
* Data Analysis: =GEMINI("Analyze this data trend: " & C1)
* Categorization: =GEMINI("Categorize this item: " & D1)

## Features Explained

### Rate Limiting
* Uses token bucket algorithm
* 60 requests per minute limit
* Automatic token refill

### Response Caching
* 6-hour cache duration
* Prevents duplicate API calls
* Improves performance for repeated queries

### Auto-Convert Feature
* Option to convert formulas to static values
* Can be toggled globally or per-formula
* Useful for preserving results
