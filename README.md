# google-sheets-gemini
🤔 Tired of paying monthly subscriptions just to use AI in Google Sheets? Here's a SIMPLE, forever-free script that gives you Gemini AI right in your spreadsheet. No coding knowledge needed!

# What is This?

This is a simple script that adds a `=GEMINI()` formula to your Google Sheet. Just like you use `=SUM()` or `=AVERAGE()`, you can now use `=GEMINI()` to ask AI questions directly in your cells. And the best part? It's completely free - you just need Google's free API key.

## Why This is Awesome

* 🆓 **Forever Free**: Just get a free API key from Google - no subscriptions, no hidden fees
* 🔌 **Super Simple Setup**: Copy, paste, done. Seriously.
* ⚡ **Fast**: Built-in caching means repeated queries are instant
* 🛡️ **Safe**: Your API key stays private in your own Google Sheet
* 🎯 **Just Works**: No coding knowledge required

## 5-Minute Setup

1. Open your Google Sheet
2. Click "Extensions" > "Apps Script" at the top
3. Delete any code you see and paste in our code (from the Code.gs file above)
4. Save and close that tab
5. Get your free API key:
   * Go to https://aistudio.google.com
   * Click "Get API key" in the top right
   * Create a new API key (it's free!)
   * Copy the key
6. Back in your sheet, refresh the page
7. You'll see a new "Gemini" menu - click it and paste in your API key

That's it! You're ready to use AI in your spreadsheet! 🎉

## How to Use It

Just type this in any cell:
```
=GEMINI("Your question here")
```

For example:
```
=GEMINI("Write a professional email to thank someone for their time")
```

### Real-World Examples

* **Content Writing**: `=GEMINI("Write a product description for: " & A1)`
* **Translation**: `=GEMINI("Translate to Spanish: " & B1)`
* **Analysis**: `=GEMINI("Analyze this sales trend: " & C1)`
* **Summaries**: `=GEMINI("Summarize this text: " & D1)`
* **Email Writing**: `=GEMINI("Write a professional email about: " & E1)`

## Power User Features

Want to get fancy? You can:
* Cache responses for 6 hours (saves on API usage)
* Limit to 60 requests per minute (prevents API overuse)
* Auto-convert formulas to static text (great for saving important responses)
* Add system prompts for more specific instructions

### Advanced Formula (Optional)
```
=GEMINI("Your question", "Optional system prompt", 0.7, true)
```

## Need Help?

* If you get stuck, check that you:
  * Copied ALL the code
  * Got your API key from Google AI Studio
  * Pasted your API key using the Gemini menu
  * Refreshed your sheet after setup

Still have questions? Open an issue in this repo! 

## Why I Made This

I got tired of seeing simple AI features locked behind expensive subscriptions. Google provides free access to Gemini - we should all be able to use it easily in our spreadsheets! 

Enjoy your free AI-powered spreadsheet! 🚀

# Advanced Features Guide

Don't worry - all these features work automatically! But here's how to control them if you want to:

## Caching (Save API Costs)

The script automatically saves responses for 6 hours. This means:
* If you ask "What is 2+2?" in cell A1
* Then copy that formula to A2
* The second call won't use your API quota - it's free!

You don't need to do anything - this happens automatically to save you money and make responses faster.

## Rate Limiting (Prevent Overuse)

The script automatically limits itself to 60 requests per minute to keep you within Google's free limits. If you hit the limit:
* You'll see "Rate limit exceeded. Please try again later."
* Just wait a minute and try again
* The limit resets automatically

This protects your API key from accidental overuse (like if you drag a formula down 1000 cells).

## Auto-Converting Formulas (Save Important Responses)

Sometimes you want to keep an AI response exactly as it is, without it changing if you modify the formula. Here's how:

1. Click the "Gemini" menu at the top
2. Select "Toggle Auto-Convert to Values"
3. When enabled:
   * Your GEMINI formulas will automatically convert to plain text
   * This "locks in" the response
   * Great for important content you want to keep

You can also do this for individual formulas using the advanced formula.

## Advanced Formula Explained

The basic formula is:
```
=GEMINI("Write a thank you email")
```

The advanced formula has more options:
```
=GEMINI("Write a thank you email", "Make it formal and professional", 0.7, true)
```

Let's break this down in simple terms:

1. First part: Your question or prompt
   * `"Write a thank you email"`
   * This is what you want Gemini to do

2. Second part: System prompt (optional)
   * `"Make it formal and professional"`
   * This gives Gemini extra instructions
   * Like telling a person "Oh, and make sure it's..."

3. Third part: Temperature (optional)
   * `0.7` is the default
   * Lower (like 0.2) = more focused, repetitive responses
   * Higher (like 0.8) = more creative, varied responses
   * Think of it like a "creativity knob"

4. Fourth part: Auto-convert (optional)
   * `true` or `false`
   * `true` = convert formula to text immediately
   * `false` = keep as formula (default)
   * Use `true` when you want to "lock in" a good response

Examples:

More focused response:
```
=GEMINI("Write a thank you email", "Keep it under 3 sentences", 0.2, false)
```

Creative response that saves automatically:
```
=GEMINI("Write a thank you email", "Make it witty and fun", 0.9, true)
```

Remember: You only need these advanced options if you want them. The basic `=GEMINI("...")` works great for most uses!
