# Google Sheets Gemini Pro Integration
🚀 A lightning-fast, free integration of Google's Gemini Pro AI into Google Sheets. No subscriptions, no complexity - just powerful AI in your spreadsheet.

## What's New in v2.0

* ⚡ **Optimized Performance**: Dramatically faster response times
* 🔄 **Smart Caching**: Improved API key handling for faster execution
* 📊 **Response Time Analytics**: Built-in latency testing and logging
* 🎯 **Enhanced Rate Limiting**: More efficient request management
* 🛠️ **Developer Tools**: Added testing and debugging features

## Why Choose This Integration?

* 🆓 **Forever Free**: Uses Google's free Gemini Pro API
* ⚡ **High Performance**: Optimized for speed and reliability
* 🔌 **Easy Setup**: Copy, paste, and you're running
* 🛡️ **Secure**: Your API key stays in your Google Sheet
* 📈 **Production Ready**: Built for reliability at scale

## Quick Setup

1. Open your Google Sheet
2. Go to "Extensions" > "Apps Script"
3. Replace any code with our script (from Code.gs)
4. Save and close
5. Get your API key:
   * Visit https://aistudio.google.com
   * Click "Get API Key" (top right)
   * Create a new key (free)
   * Copy it
6. Refresh your sheet
7. Use the "Gemini" menu to enter your key

## Basic Usage

=GEMINI("Your prompt here")


Example prompts:
* `=GEMINI("Write a professional email about: " & A1)`
* `=GEMINI("Translate to Spanish: " & B1)`
* `=GEMINI("Analyze these numbers: " & C1:C10)`

## Performance Features

### Speed Optimization
* Cached API key handling
* Efficient rate limiting
* Response time logging
* Built-in latency testing

### Rate Limiting
* 300 requests per minute capacity
* Smart request queuing
* Automatic retry on 429 errors

### Monitoring
* Test response times via menu
* Detailed execution logs
* Performance breakdown by operation

## Advanced Features

### Formula Options
=GEMINI("prompt", "system_prompt", temperature)


Parameters:
* `prompt`: Your main query
* `system_prompt`: Additional instructions (optional)
* `temperature`: Creativity level 0-1 (default: 0.7)

### Utility Functions
* Test response times
* Convert formulas to values
* Monitor API usage

## Debugging

Added "Test Response Time" in the Gemini menu to check:
* API response speed
* Total execution time
* Component-level timing

## Common Issues

If you experience slow responses:
* Check your network connection
* Run the latency test
* Review execution logs
* Verify API key validity

## Contributing

Found a way to make it even faster? We love contributions! Please:
1. Fork the repository
2. Create a feature branch
3. Submit a pull request

## Support

* Create an issue for bugs/features
* Check execution logs for troubleshooting
* Run latency tests for performance issues

## License

MIT License - Free for commercial and personal use

---

Made with ⚡ by developers who believe AI should be fast, free, and accessible to everyone.
