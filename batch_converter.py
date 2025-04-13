#!/usr/bin/env python3

import os
import sys
import glob
import argparse
import concurrent.futures
from typing import List, Dict, Any, Optional
from termcolor import colored
import time

# Import from spreadsheet_converter
from spreadsheet_converter import (
    convert_spreadsheet_to_json, 
    save_json_output, 
    SUPPORTED_EXTENSIONS,
    print_status,
    TOKEN_EFFICIENT,
    convert_spreadsheet_to_json_with_sampling
)

def process_workbook(file_path: str, args: Dict[str, Any]) -> Dict[str, Any]:
    """Process a single workbook with the given arguments."""
    try:
        print_status(f"Processing: {os.path.basename(file_path)}", 'info')
        start_time = time.time()
        
        # Convert spreadsheet to JSON
        if args.get('intelligent_sampling', False):
            # Use the version with intelligent sampling
            json_data = convert_spreadsheet_to_json_with_sampling(
                file_path, 
                args['sample_size'], 
                args['formulas_only'], 
                args['keep_formatting'],
                not args['no_context'],
                True  # intelligent_sampling=True
            )
        else:
            # Use the regular version without intelligent sampling
            json_data = convert_spreadsheet_to_json(
                file_path, 
                args['sample_size'], 
                args['formulas_only'], 
                args['keep_formatting'],
                not args['no_context']
            )
        
        # Save with token count in filename
        output_path = save_json_output(json_data, file_path, args['minify'])
        
        # Get token count
        metadata_key = 'm' if TOKEN_EFFICIENT else 'metadata'
        token_count_key = 'tc' if TOKEN_EFFICIENT else 'token_count'
        token_count = json_data.get(metadata_key, {}).get(token_count_key, 0)
        
        end_time = time.time()
        processing_time = end_time - start_time
        
        return {
            'file_path': file_path,
            'output_path': output_path,
            'token_count': token_count,
            'success': True,
            'processing_time': processing_time,
            'error': None
        }
    except Exception as e:
        return {
            'file_path': file_path,
            'output_path': None,
            'token_count': 0,
            'success': False,
            'processing_time': 0,
            'error': str(e)
        }

def find_excel_files(path_patterns: List[str]) -> List[str]:
    """Find all Excel files matching the given patterns."""
    all_files = []
    
    for pattern in path_patterns:
        # If pattern is a directory, search for Excel files inside
        if os.path.isdir(pattern):
            for ext in SUPPORTED_EXTENSIONS:
                all_files.extend(glob.glob(os.path.join(pattern, f"*{ext}")))
                all_files.extend(glob.glob(os.path.join(pattern, f"**/*{ext}"), recursive=True))
        # If pattern is a file, check if it's an Excel file
        elif os.path.isfile(pattern):
            if any(pattern.lower().endswith(ext) for ext in SUPPORTED_EXTENSIONS):
                all_files.append(pattern)
        # Handle glob patterns
        else:
            matching_files = glob.glob(pattern, recursive=True)
            excel_files = [f for f in matching_files if any(f.lower().endswith(ext) for ext in SUPPORTED_EXTENSIONS)]
            all_files.extend(excel_files)
    
    # Remove duplicates while preserving order
    unique_files = []
    for file in all_files:
        if file not in unique_files:
            unique_files.append(file)
    
    return unique_files

def parse_args():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description='Batch convert Excel workbooks to JSON in parallel',
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    
    parser.add_argument(
        'paths', 
        nargs='+', 
        help='Path(s) to Excel files, directories, or glob patterns'
    )
    
    parser.add_argument(
        '-j', '--jobs',
        type=int,
        default=os.cpu_count(),
        help='Number of parallel jobs to run'
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
    
    return parser.parse_args()

def main():
    # Parse arguments
    args = parse_args()
    
    # Set token efficiency based on arguments
    global TOKEN_EFFICIENT
    TOKEN_EFFICIENT = not args.full_names
    
    # Find all Excel files
    excel_files = find_excel_files(args.paths)
    
    if not excel_files:
        print_status("No Excel files found matching the given paths", 'error')
        sys.exit(1)
    
    print_status(f"Found {len(excel_files)} Excel files to process", 'info')
    
    # Prepare arguments for processing
    process_args = {
        'sample_size': args.rows,
        'formulas_only': args.formulas_only,
        'keep_formatting': args.keep_formatting,
        'minify': args.minify,
        'no_context': args.no_context
    }
    
    # Print settings
    print_status("Batch conversion settings:", 'info')
    print_status(f"  Parallel jobs: {args.jobs}", 'info')
    if args.rows:
        print_status(f"  Processing {args.rows} rows per sheet", 'info')
    if args.formulas_only:
        print_status("  Including only cells with formulas and their dependencies", 'info')
    if not args.keep_formatting:
        print_status("  Excluding formatting information to save tokens", 'info')
    if TOKEN_EFFICIENT:
        print_status("  Using abbreviated property names to save tokens", 'info')
    if args.minify:
        print_status("  Minifying JSON output to save tokens", 'info')
    if not args.no_context:
        print_status("  Adding enriched context for formula interpretation", 'info')
    
    # Process files in parallel
    start_time = time.time()
    results = []
    
    with concurrent.futures.ProcessPoolExecutor(max_workers=args.jobs) as executor:
        # Submit all workbooks for processing
        future_to_file = {
            executor.submit(process_workbook, file, process_args): file
            for file in excel_files
        }
        
        # Process results as they complete
        for future in concurrent.futures.as_completed(future_to_file):
            file = future_to_file[future]
            try:
                result = future.result()
                results.append(result)
                
                if result['success']:
                    print_status(
                        f"✓ {os.path.basename(file)} → {os.path.basename(result['output_path'])} "
                        f"({result['token_count']} tokens, {result['processing_time']:.2f}s)",
                        'success'
                    )
                else:
                    print_status(
                        f"✗ {os.path.basename(file)} - Error: {result['error']}",
                        'error'
                    )
            except Exception as e:
                print_status(f"✗ {os.path.basename(file)} - Exception: {str(e)}", 'error')
                results.append({
                    'file_path': file,
                    'success': False,
                    'error': str(e)
                })
    
    # Print summary
    end_time = time.time()
    total_time = end_time - start_time
    successful = sum(1 for r in results if r['success'])
    failed = len(results) - successful
    
    print_status("\nBatch Conversion Summary:", 'info')
    print_status(f"  Total files processed: {len(results)}", 'info')
    print_status(f"  Successful conversions: {successful}", 'success')
    if failed > 0:
        print_status(f"  Failed conversions: {failed}", 'error')
    print_status(f"  Total processing time: {total_time:.2f} seconds", 'info')
    
    if successful > 0:
        # Calculate average tokens per workbook
        avg_tokens = sum(r['token_count'] for r in results if r['success']) / successful
        print_status(f"  Average tokens per workbook: {avg_tokens:.0f}", 'info')
        
        print_status("\nLLM Recommendations based on average token count:", 'info')
        if avg_tokens < 4000:
            print_status("✓ Suitable for GPT-3.5-turbo (4K context)", 'success')
        elif avg_tokens < 8000:
            print_status("✓ Suitable for GPT-3.5-turbo-8K", 'success')
        elif avg_tokens < 16000:
            print_status("✓ Suitable for GPT-3.5-turbo-16K", 'success')
        elif avg_tokens < 32000:
            print_status("✓ Suitable for GPT-4 (32K context)", 'success')
        else:
            print_status("⚠ Warning: Large token count. Consider splitting the data or using a model with larger context window", 'warning')
    
    print_status("\nBatch conversion completed!", 'success')

if __name__ == "__main__":
    main() 