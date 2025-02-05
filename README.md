# JSON-like SpreadSheet

Turn Excel hell into something LLMs can actually understand. Because spreadsheets shouldn't be stuck in the past.

## Why?

Let's be real - spreadsheets are everywhere and they're a pain. This tool:
- Converts Excel mess into clean JSON
- Makes it actually readable for AI
- Keeps the important stuff intact
- Handles massive files without breaking

✨ **What you get:**
- Smart JSON that AI can work with
- No more token limit headaches
- All the formulas and relationships preserved
- Works with any Excel file

## Quick Start

1. Get it:
```bash
git clone https://github.com/FrancyJGLisboa/JSON-like-SpreadSheet.git
cd JSON-like-SpreadSheet
pip install -r requirements.txt
```

2. Use it:
```bash
python spreadsheet_converter.py your_file.xlsx
```

Need to handle a huge file? No problem:
```bash
python spreadsheet_converter.py your_file.xlsx 100  # Process 100 rows per sheet
```

## What It Captures

Everything that matters:
- Formulas & calculations
- Styles & formatting
- Data rules & validation
- Protection settings
- Charts & pivot tables
- Cell connections

## Output

Gets you a clean JSON file with:
- Auto token counting
- Smart file naming
- AI model suggestions
- Full context preserved

Example structure:
```json
{
  "file_name": "example.xlsx",
  "metadata": {
    "token_count": 1234,
    "conversion_timestamp": "20240205_155518"
  },
  "sheets": {
    "Sheet1": {
      "cells": {
        "A1": {
          "value": "Data",
          "formula": "=SUM(B1:B10)",
          "style": {...}
        }
      }
    }
  }
}
```

## Using with AI

The JSON output works great with:
- GPT-3.5 (< 4K tokens)
- GPT-4 (< 32K tokens)
- Claude & other LLMs

Quick prompt template:
```
Check out this spreadsheet and tell me:
1. What's it trying to do?
2. Any formulas that could break?
3. How to make it better?
```

## Pro Tips

💡 For big files:
- Start with a few rows
- Keep the headers
- Split if needed

🔥 For best results:
- Let the AI see the structure
- Keep the formulas connected
- Use the token count

## Want to Help?

Open issues or PRs if you've got ideas to make this better.

## License

MIT - do what you want with it. 