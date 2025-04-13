#!/usr/bin/env python3

import os
import sys
import glob
import json
import argparse
import re
import time
import hashlib
from typing import Dict, List, Any, Set, Tuple
from termcolor import colored

# Import from our existing converter
from spreadsheet_converter import (
    SUPPORTED_EXTENSIONS,
    TOKEN_EFFICIENT,
    PROPERTY_MAP,
    print_status,
    ExcelJSONEncoder
)

# Import but don't run the batch converter main function
from batch_converter import find_excel_files, process_workbook

# Regex patterns to detect cross-workbook references
# Matches patterns like '[Workbook.xlsx]Sheet1'!A1 or '[Workbook.xlsx]Sheet1!A1'
CROSS_WB_REF_PATTERN = r"\[([^\]]+)\]([^!]+)!([A-Z]+[0-9]+)"

def map_key(key: str) -> str:
    """Map a key to its token-efficient version if TOKEN_EFFICIENT is True."""
    if TOKEN_EFFICIENT and key in PROPERTY_MAP:
        return PROPERTY_MAP[key]
    return key

def generate_workbook_id(filename: str) -> str:
    """Generate a short, unique identifier for a workbook based on its filename."""
    # Remove extension and create a hash-based ID
    base_name = os.path.splitext(os.path.basename(filename))[0]
    hash_id = hashlib.md5(base_name.encode()).hexdigest()[:6]  # Short 6-character hash
    
    # Create a clean version of the name (alphanumeric only) + hash
    clean_name = re.sub(r'[^a-zA-Z0-9]', '', base_name)[:8]  # Take first 8 chars of name
    return f"{clean_name}_{hash_id}"

def extract_cross_workbook_references(json_data: Dict, workbook_id: str) -> Dict[str, List[str]]:
    """
    Extract cross-workbook references from formulas in the JSON data.
    Returns a dictionary mapping source cells to target cells.
    """
    cross_references = {}
    
    # Alias key names based on token efficiency
    sheets_key = map_key('sheets')
    cells_key = map_key('cells')
    formula_key = map_key('formula')
    
    # Scan for formulas
    for sheet_name, sheet_data in json_data.get(sheets_key, {}).items():
        for cell_ref, cell_data in sheet_data.get(cells_key, {}).items():
            # Check if cell has a formula
            if isinstance(cell_data, dict) and formula_key in cell_data:
                formula = cell_data[formula_key]
                
                # Look for cross-workbook references
                matches = re.finditer(CROSS_WB_REF_PATTERN, formula)
                
                for match in matches:
                    target_workbook = match.group(1)  # The referenced workbook name
                    target_sheet = match.group(2)  # The referenced sheet name
                    target_cell = match.group(3)  # The referenced cell
                    
                    # Create source and target cell identifiers
                    source_cell_id = f"{workbook_id}_{sheet_name}_{cell_ref}"
                    
                    # We'll replace this with the actual workbook ID later
                    target_cell_id = f"{target_workbook}_{target_sheet}_{target_cell}"
                    
                    # Add to cross-references
                    if source_cell_id not in cross_references:
                        cross_references[source_cell_id] = []
                    
                    cross_references[source_cell_id].append(target_cell_id)
    
    return cross_references

def resolve_cross_references(cross_refs: Dict[str, List[str]], workbook_id_map: Dict[str, str]) -> Dict[str, List[str]]:
    """
    Resolve cross-workbook references by replacing workbook filenames with workbook IDs.
    """
    resolved_refs = {}
    
    for source_cell, target_cells in cross_refs.items():
        resolved_refs[source_cell] = []
        
        for target_cell in target_cells:
            # Extract workbook name, sheet name, and cell reference
            parts = target_cell.split('_', 2)
            if len(parts) == 3:
                wb_name, sheet_name, cell_ref = parts
                
                # Replace workbook name with workbook ID if it exists in our map
                if wb_name in workbook_id_map:
                    resolved_target = f"{workbook_id_map[wb_name]}_{sheet_name}_{cell_ref}"
                    resolved_refs[source_cell].append(resolved_target)
                else:
                    # Keep the reference as is if we can't resolve it
                    resolved_refs[source_cell].append(target_cell)
    
    return resolved_refs

def merge_workbooks(json_files: List[str]) -> Dict[str, Any]:
    """
    Merge multiple workbook JSON files into a single consolidated structure.
    """
    # Initialize the consolidated structure
    consolidated = {
        map_key('workbooks'): {},
        map_key('cross_references'): {},
        map_key('enriched_context'): {
            map_key('tables'): {},
            map_key('column_types'): {},
            map_key('formula_patterns'): {},
            map_key('implementation_notes'): {}
        },
        map_key('metadata'): {
            map_key('workbook_count'): len(json_files),
            map_key('conversion_timestamp'): time.strftime('%Y%m%d_%H%M%S')
        }
    }
    
    # Map workbook filenames to IDs
    workbook_id_map = {}
    all_cross_refs = {}
    
    # First pass: extract workbook data and collect cross-references
    for json_file in json_files:
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Get the original filename
            metadata_key = map_key('metadata')
            original_filename_key = map_key('original_filename')
            filename = data.get(metadata_key, {}).get(original_filename_key, os.path.basename(json_file))
            
            # Generate workbook ID
            workbook_id = generate_workbook_id(filename)
            workbook_id_map[filename] = workbook_id
            
            # Extract workbook data
            consolidated[map_key('workbooks')][workbook_id] = {
                map_key('file_name'): filename,
                map_key('sheets'): data.get(map_key('sheets'), {}),
                map_key('named_ranges'): data.get(map_key('named_ranges'), {})
            }
            
            # Extract cross-workbook references
            cross_refs = extract_cross_workbook_references(data, workbook_id)
            all_cross_refs.update(cross_refs)
            
            # Merge enriched context
            ec_key = map_key('enriched_context')
            if ec_key in data:
                # Merge tables
                tb_key = map_key('tables')
                if tb_key in data[ec_key]:
                    for table_name, table_data in data[ec_key][tb_key].items():
                        # Prefix table names with workbook ID to avoid collisions
                        consolidated[ec_key][tb_key][f"{workbook_id}_{table_name}"] = table_data
                
                # Merge column types
                ct_key = map_key('column_types')
                if ct_key in data[ec_key]:
                    consolidated[ec_key][ct_key].update(data[ec_key][ct_key])
                
                # Merge formula patterns
                fp_key = map_key('formula_patterns')
                if fp_key in data[ec_key]:
                    for pattern_name, pattern in data[ec_key][fp_key].items():
                        # Prefix pattern names with workbook ID
                        consolidated[ec_key][fp_key][f"{workbook_id}_{pattern_name}"] = pattern
                
                # Merge implementation notes
                in_key = map_key('implementation_notes')
                if in_key in data[ec_key]:
                    consolidated[ec_key][in_key].update(data[ec_key][in_key])
        
        except Exception as e:
            print_status(f"Error processing {json_file}: {str(e)}", 'error')
    
    # Second pass: resolve cross-references
    resolved_refs = resolve_cross_references(all_cross_refs, workbook_id_map)
    consolidated[map_key('cross_references')] = resolved_refs
    
    # Update metadata
    token_count = estimate_token_count(consolidated)
    consolidated[map_key('metadata')][map_key('token_count')] = token_count
    
    return consolidated

def estimate_token_count(data: Dict[str, Any]) -> int:
    """
    Estimate the token count of the consolidated JSON.
    This is a simplified estimation - for more accuracy, use a proper tokenizer.
    """
    # Convert to JSON string
    json_str = json.dumps(data, cls=ExcelJSONEncoder)
    
    # Simple estimation: ~4 characters per token
    return len(json_str) // 4

def save_consolidated_json(data: Dict[str, Any], output_dir: str, minify: bool = False) -> str:
    """
    Save the consolidated JSON data to a file.
    """
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Generate filename
    timestamp = time.strftime('%Y%m%d_%H%M%S')
    workbook_count = data.get(map_key('metadata'), {}).get(map_key('workbook_count'), 0)
    token_count = data.get(map_key('metadata'), {}).get(map_key('token_count'), 0)
    
    filename = f"consolidated_{workbook_count}workbooks_{token_count}tokens_{timestamp}.json"
    output_path = os.path.join(output_dir, filename)
    
    # Save the file
    with open(output_path, 'w', encoding='utf-8') as f:
        if minify:
            json.dump(data, f, ensure_ascii=False, cls=ExcelJSONEncoder, separators=(',', ':'))
        else:
            json.dump(data, f, indent=2, ensure_ascii=False, cls=ExcelJSONEncoder)
    
    print_status(f"Successfully saved consolidated JSON to: {output_path}", 'success')
    return output_path

def process_directory(directory: str, args) -> List[str]:
    """
    Process all Excel files in a directory using the batch processing logic.
    Returns a list of generated JSON files.
    """
    # Find all Excel files in the directory
    excel_files = find_excel_files([directory])
    
    if not excel_files:
        print_status(f"No Excel files found in {directory}", 'error')
        return []
    
    print_status(f"Found {len(excel_files)} Excel files to process", 'info')
    
    # Prepare arguments for individual workbook processing
    process_args = {
        'sample_size': args.rows,
        'formulas_only': args.formulas_only,
        'keep_formatting': args.keep_formatting,
        'minify': args.minify,
        'no_context': args.no_context,
        'intelligent_sampling': args.intelligent_sampling
    }
    
    # Process each workbook
    json_files = []
    
    for file_path in excel_files:
        try:
            print_status(f"Processing: {os.path.basename(file_path)}", 'info')
            result = process_workbook(file_path, process_args)
            
            if result['success']:
                json_files.append(result['output_path'])
                print_status(
                    f"✓ {os.path.basename(file_path)} → {os.path.basename(result['output_path'])}",
                    'success'
                )
            else:
                print_status(
                    f"✗ {os.path.basename(file_path)} - Error: {result['error']}",
                    'error'
                )
        except Exception as e:
            print_status(f"✗ {os.path.basename(file_path)} - Exception: {str(e)}", 'error')
    
    return json_files

def parse_args():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description='Combine multiple Excel workbooks into a single consolidated JSON',
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    
    parser.add_argument(
        'directory',
        help='Directory containing Excel workbooks to process'
    )
    
    parser.add_argument(
        '-o', '--output-dir',
        default='consolidated_json',
        help='Directory to save the consolidated JSON'
    )
    
    parser.add_argument(
        '-r', '--rows',
        type=int,
        help='Number of rows to process per sheet'
    )
    
    parser.add_argument(
        '-f', '--formulas-only',
        action='store_true',
        help='Only include cells with formulas and their dependencies'
    )
    
    parser.add_argument(
        '-k', '--keep-formatting',
        action='store_true',
        help='Include formatting information (uses more tokens)'
    )
    
    parser.add_argument(
        '-m', '--minify',
        action='store_true',
        help='Remove whitespace from JSON output'
    )
    
    parser.add_argument(
        '-n', '--full-names',
        action='store_true',
        help='Use full property names instead of abbreviated ones'
    )
    
    parser.add_argument(
        '-c', '--no-context',
        action='store_true',
        help='Skip adding enriched context information'
    )
    
    parser.add_argument(
        '-s', '--skip-processing',
        action='store_true',
        help='Skip individual workbook processing, use existing JSON files'
    )
    
    parser.add_argument(
        '-i', '--intelligent-sampling',
        action='store_true',
        help='Use formula-preserving intelligent sampling to reduce token count'
    )
    
    return parser.parse_args()

def main():
    # Parse arguments
    args = parse_args()
    
    # Set token efficiency based on arguments
    global TOKEN_EFFICIENT
    TOKEN_EFFICIENT = not args.full_names
    
    start_time = time.time()
    
    # Check if directory exists
    if not os.path.isdir(args.directory):
        print_status(f"Directory not found: {args.directory}", 'error')
        sys.exit(1)
    
    # Process directory to get JSON files
    json_files = []
    
    if args.skip_processing:
        # Use existing JSON files in the converted_json directory
        json_dir = os.path.join(os.path.dirname(args.directory), 'converted_json')
        if os.path.isdir(json_dir):
            json_files = glob.glob(os.path.join(json_dir, '*.json'))
            print_status(f"Using {len(json_files)} existing JSON files from {json_dir}", 'info')
        else:
            print_status(f"No existing JSON files found in {json_dir}", 'error')
            sys.exit(1)
    else:
        # Process all workbooks in the directory
        json_files = process_directory(args.directory, args)
        
        if not json_files:
            print_status("No JSON files generated. Cannot proceed with consolidation.", 'error')
            sys.exit(1)
    
    # Merge workbooks
    print_status(f"Merging {len(json_files)} workbooks into a consolidated JSON...", 'info')
    consolidated_data = merge_workbooks(json_files)
    
    # Save consolidated JSON
    save_consolidated_json(consolidated_data, args.output_dir, args.minify)
    
    # Print summary
    end_time = time.time()
    total_time = end_time - start_time
    
    print_status("\nConsolidation Summary:", 'info')
    print_status(f"  Total workbooks processed: {len(json_files)}", 'info')
    print_status(f"  Cross-workbook references: {len(consolidated_data.get(map_key('cross_references'), {}))}", 'info')
    print_status(f"  Total processing time: {total_time:.2f} seconds", 'info')
    
    # Token count recommendations
    token_count = consolidated_data.get(map_key('metadata'), {}).get(map_key('token_count'), 0)
    print_status(f"  Estimated token count: {token_count}", 'info')
    
    # Include info about sampling in the summary
    if args.intelligent_sampling:
        print_status("  Used intelligent sampling to reduce token count", 'info')
    
    print_status("\nLLM Recommendations based on token count:", 'info')
    if token_count < 4000:
        print_status("✓ Suitable for GPT-3.5-turbo (4K context)", 'success')
    elif token_count < 8000:
        print_status("✓ Suitable for GPT-3.5-turbo-8K", 'success')
    elif token_count < 16000:
        print_status("✓ Suitable for GPT-3.5-turbo-16K", 'success')
    elif token_count < 32000:
        print_status("✓ Suitable for GPT-4 (32K context)", 'success')
    else:
        print_status("⚠ Warning: Large token count. Consider splitting the data or using a model with larger context window", 'warning')
    
    print_status("\nConsolidation completed successfully!", 'success')

if __name__ == "__main__":
    main() 