#!/usr/bin/env python3

import argparse
import json
import sys
from typing import Dict, Any, Optional
from termcolor import colored
from excel_converter import excel_to_json, print_status

def process_excel_file(file_path: str, sample_size: Optional[int] = None) -> bool:
    """
    Process local Excel file with optional sample size limit.
    Returns True if successful, False otherwise.
    """
    try:
        print_status(f"Processing Excel file: {file_path}", 'info')
        if sample_size:
            print_status(f"Using sample size of {sample_size} rows per sheet", 'info')

        data = excel_to_json(file_path, sample_size)
        if data is None:
            print_status("Failed to convert Excel file", 'error')
            return False

        print(json.dumps(data, indent=2, ensure_ascii=False))
        print_status("Conversion completed successfully!", 'success')
        return True

    except Exception as e:
        print_status(f"Error processing Excel file: {str(e)}", 'error')
        return False

def main() -> int:
    """Main function. Returns exit code (0 for success, 1 for failure)."""
    parser = argparse.ArgumentParser(
        description="Convert an Excel spreadsheet to JSON. Optionally sample a limited number of rows."
    )
    parser.add_argument("--file", type=str, help="Path to a local Excel file.")
    parser.add_argument(
        "--sample-size",
        type=int,
        default=None,
        help="Number of rows per sheet to capture (optional)."
    )

    args = parser.parse_args()

    # If file is not provided, print help
    if not args.file:
        parser.print_help()
        return 1

    success = process_excel_file(args.file, args.sample_size)
    return 0 if success else 1

if __name__ == "__main__":
    sys.exit(main()) 