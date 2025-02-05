# JSON-like SpreadSheet Converter

Transform Excel spreadsheets into LLM-friendly JSON format, making it easy for Large Language Models to understand and work with spreadsheet data.

## Why This Tool?

### The Spreadsheet-to-LLM Challenge
Spreadsheets are everywhere in business, but they're not easily consumable by Large Language Models (LLMs). This tool bridges that gap by:
- Converting complex Excel structures into clean, structured JSON
- Making spreadsheet data and formulas "readable" for LLMs
- Preserving essential context and relationships
- Managing token limits through smart sampling

### Key Benefits
- **LLM Integration**: Perfect for applications where LLMs need to analyze, explain, or transform spreadsheet data
- **Token Management**: Built-in token counting and sampling to stay within LLM context limits
- **Complete Context**: Captures formulas, styles, and relationships that help LLMs understand spreadsheet logic
- **Universal Format**: Converts widespread Excel files into JSON that any LLM can process

## Features

- Converts Excel files (.xlsx, .xls, .xlsm) to detailed JSON format
- Captures comprehensive spreadsheet metadata:
  - Named ranges
  - Cell values and formulas
  - Styles and formatting
  - Conditional formatting rules
  - Data validation
  - Protection settings
  - Sheet view settings
  - Pivot tables
  - Charts
  - Form controls
  - Auto-filters
  - Cell dependencies
- Intelligent token counting for LLM compatibility
- Optional row sampling for large spreadsheets
- Colored terminal output for better visibility
- Comprehensive error handling

## Requirements

```
openpyxl
termcolor
```

## Installation

1. Clone the repository:
```bash
git clone https://github.com/FrancyJGLisboa/JSON-like-SpreadSheet.git
cd JSON-like-SpreadSheet
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Basic usage:
```bash
python spreadsheet_converter.py <path_to_spreadsheet>
```

With row sampling for large spreadsheets (helps stay within LLM token limits):
```bash
python spreadsheet_converter.py <path_to_spreadsheet> <sample_size>
```

### Example
```bash
# Convert spreadsheet with 100 rows per sheet (good for GPT-3.5's 4K token limit)
python spreadsheet_converter.py example.xlsx 100
```

## Output

The script generates an LLM-friendly JSON file in the `converted_json` directory with:
- Automatic token counting for LLM compatibility
- Smart filename format including token count
- LLM-specific recommendations based on file size
- Complete metadata for context preservation

### Output Format
```json
{
  "file_name": "example.xlsx",
  "metadata": {
    "token_count": 1234,
    "conversion_timestamp": "20240205_155518",
    "original_filename": "example.xlsx"
  },
  "named_ranges": {
    "MyRange": {
      "value": "Sheet1!A1:B10",
      "scope": "workbook"
    }
  },
  "sheets": {
    "Sheet1": {
      "metadata": {
        "title": "Sheet1",
        "dimensions": "A1:Z100",
        "max_row": 100,
        "max_column": 26
      },
      "cells": {
        "A1": {
          "value": "Example",
          "style": {
            "font": {
              "bold": true,
              "italic": false,
              "color": "FF0000"
            }
          }
        }
      },
      "merged_cells": ["A1:B1"],
      "conditional_formatting": [...],
      "protection": {...},
      "view_settings": {...}
    }
  }
}
```

## Error Handling

The script includes comprehensive error handling:
- Validates file existence and format
- Handles missing attributes gracefully
- Provides detailed error messages
- Continues processing when encountering non-critical errors

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## LLM Integration Guide

### Token Management
The tool helps manage LLM token limits by:
- Counting tokens in the output JSON
- Providing sampling options for large spreadsheets
- Suggesting appropriate LLM models based on token count

### Model Recommendations
Based on the token count, the tool suggests suitable LLM models:
- < 4,000 tokens: Ideal for GPT-3.5-turbo
- < 8,000 tokens: Suitable for GPT-3.5-turbo-8K
- < 16,000 tokens: Works with GPT-3.5-turbo-16K
- < 32,000 tokens: Compatible with GPT-4
- \> 32,000 tokens: Consider using sampling or splitting the data

### Example LLM Prompts
Here are some ways to use the JSON output with LLMs:

1. **Analysis Prompt**:
```
Analyze this spreadsheet JSON and explain:
1. The overall structure and purpose
2. Key formulas and their relationships
3. Data validation rules and their meaning
4. Conditional formatting logic
```

2. **Transformation Prompt**:
```
Based on this spreadsheet JSON:
1. Suggest improvements to the formula structure
2. Identify potential errors or inconsistencies
3. Recommend optimization opportunities
4. Explain the business logic encoded in the formulas
```

3. **Documentation Prompt**:
```
Create documentation for this spreadsheet by explaining:
1. The purpose of each sheet
2. How the named ranges are used
3. The protection scheme and its business logic
4. The relationship between different data elements
```

## Best Practices

1. **Token Optimization**:
   - Start with a small sample (3-4 rows) to test LLM understanding
   - Include header rows for context
   - Use sampling strategically for large sheets

2. **Context Preservation**:
   - Keep metadata and structure information
   - Preserve formula relationships
   - Maintain styling that indicates importance

3. **LLM Integration**:
   - Use the token count to choose appropriate models
   - Structure prompts to leverage the JSON format
   - Consider chunking for large spreadsheets 