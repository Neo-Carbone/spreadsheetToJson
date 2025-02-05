# JSON-like SpreadSheet Converter

A powerful Python tool that converts Excel spreadsheets into a comprehensive JSON format, capturing all essential spreadsheet elements including formulas, styles, conditional formatting, and more.

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

With row sampling (process only N rows per sheet):
```bash
python spreadsheet_converter.py <path_to_spreadsheet> <sample_size>
```

### Example
```bash
python spreadsheet_converter.py example.xlsx 100
```

## Output

The script generates a JSON file in the `converted_json` directory with:
- Filename including token count and timestamp
- Complete spreadsheet structure and metadata
- LLM compatibility recommendations based on token count

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