# Token-Efficient Excel to JSON Converter

A tool for converting Excel workbooks to a token-efficient JSON format, suitable for use with large language models.

## Features

- **Token Efficiency**: Abbreviates property names and only includes essential data.
- **Formula Extraction**: Preserves formulas and their structure for AI interpretation.
- **Enriched Context**: Adds context about tables, column types, and formula patterns.
- **Named Range Support**: Includes information about named ranges used in formulas.
- **Pivot Table Support**: Extracts pivot table structures and fields in a token-efficient format.
- **Command-line Options**: Flexible options for customizing the conversion process.

## Usage

```bash
python spreadsheet_converter.py <excel_file> [options]
```

### Command-line Options

- `--rows N`: Process only the first N rows of each sheet (default: all rows)
- `--formulas-only`: Only include cells with formulas and their dependencies
- `--keep-formatting`: Include formatting information (uses more tokens)
- `--no-minify`: Output formatted JSON instead of minified (uses more tokens)
- `--no-context`: Don't add enriched context information

## Enriched Context

The converter provides enriched context in the resulting JSON to help AI understand and implement Excel functionality:

1. **Table Structures**: Identifies Excel tables, their columns, and ranges
2. **Column Types**: Detects data types (string, number, date, currency) for columns
3. **Formula Patterns**: Extracts frequently used formula patterns
4. **Sample Calculated Values**: Provides samples of calculated values for testing
5. **Implementation Notes**: Guidance for JavaScript implementation of formulas
6. **Pivot Tables**: Information about pivot tables, their fields, and aggregation functions

## Example

Input Excel file:
```
| Project | Amount | Tax   | Total  |
|---------|--------|-------|--------|
| A       | 100    | 15    | =B2+C2 |
| B       | 200    | 30    | =B3+C3 |
```

Output JSON (minified for brevity):
```json
{
  "fn": "example.xlsx",
  "sh": {
    "Sheet1": {
      "cl": {
        "A1": {"v": "Project"},
        "B1": {"v": "Amount"},
        "C1": {"v": "Tax"},
        "D1": {"v": "Total"},
        "A2": {"v": "A"},
        "B2": {"v": 100},
        "C2": {"v": 15},
        "D2": {"v": {"f": "=B2+C2", "cv": 115}, "d": {"cr": ["B2", "C2"]}},
        "A3": {"v": "B"},
        "B3": {"v": 200},
        "C3": {"v": 30},
        "D3": {"v": {"f": "=B3+C3", "cv": 230}, "d": {"cr": ["B3", "C3"]}}
      }
    }
  },
  "ec": {
    "ct": {"A": "string", "B": "decimal", "C": "decimal", "D": "decimal"},
    "fp": {"pattern_1": "=B#+C#"},
    "in": {"js": "This can be implemented in JavaScript as: row => row['Amount'] + row['Tax']"}
  }
}
```

## Comparison with Original Tool

| Feature | Original | Token-Efficient |
|---------|----------|-----------------|
| Property names | Full names | Abbreviated (1-2 chars) |
| Empty cells | Included | Excluded |
| Formatting | Included | Optional |
| Context | None | Tables, formulas, types |
| JSON size | Large | ~70-95% smaller |
| Pivot tables | Limited | Comprehensive |

## License

MIT 