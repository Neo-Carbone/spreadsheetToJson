#!/usr/bin/env python3

import json
import os
from typing import Dict, Any, List, Optional
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from termcolor import colored
import sys
import datetime
import re

# Constants
SUPPORTED_EXTENSIONS = ['.xlsx', '.xls', '.xlsm']
OUTPUT_DIR = 'converted_json'
DEFAULT_ENCODING = 'utf-8'

class ExcelJSONEncoder(json.JSONEncoder):
    """Custom JSON encoder to handle Excel-specific types."""
    def default(self, obj):
        try:
            if isinstance(obj, datetime.datetime):
                return obj.isoformat()
            elif isinstance(obj, datetime.date):
                return obj.isoformat()
            elif hasattr(obj, 'rgb'):
                # Handle RGB objects from openpyxl
                if obj.rgb:
                    return obj.rgb.decode() if isinstance(obj.rgb, bytes) else str(obj.rgb)
                return None
            elif hasattr(obj, '__dict__'):
                # Handle other objects by converting their attributes to dict
                return {k: v for k, v in obj.__dict__.items() 
                       if not k.startswith('_') and v is not None}
            return super().default(obj)
        except Exception as e:
            # If we can't serialize, return None instead of failing
            print_status(f"Warning: Could not serialize object {type(obj)}: {str(e)}", 'warning')
            return None

def print_status(message: str, status: str = 'info') -> None:
    """Print colored status messages."""
    color_map = {
        'info': 'cyan',
        'success': 'green',
        'error': 'red',
        'warning': 'yellow'
    }
    print(colored(message, color_map.get(status, 'white')))

def get_data_validation(cell: Any) -> Optional[Dict[str, Any]]:
    """Extract data validation information from a cell."""
    try:
        if hasattr(cell, 'data_validation') and cell.data_validation:
            validation_data = {
                'type': cell.data_validation.type,
                'operator': cell.data_validation.operator,
                'formula1': cell.data_validation.formula1,
                'formula2': cell.data_validation.formula2,
                'allow_blank': cell.data_validation.allow_blank,
                'show_error_message': cell.data_validation.showErrorMessage,
                'error_title': cell.data_validation.errorTitle,
                'error_message': cell.data_validation.error,
                'prompt_title': cell.data_validation.promptTitle,
                'prompt_message': cell.data_validation.prompt,
            }
            return validation_data
        return None
    except Exception as e:
        # Don't print error for missing data validation
        return None

def get_cell_value(cell: Any) -> Any:
    """Extract cell value and handle different data types."""
    try:
        if cell.value is None:
            return None
        
        # Handle formula
        if cell.data_type == 'f':
            return {
                'formula': str(cell.value),
                'calculated_value': cell.internal_value
            }
        
        # Handle dates
        if isinstance(cell.value, (datetime.date, datetime.datetime)):
            return cell.value.isoformat()
            
        return cell.value
    except Exception as e:
        print_status(f"Error processing cell: {str(e)}", 'error')
        return None

def get_sheet_metadata(sheet: Worksheet) -> Dict[str, Any]:
    """Extract sheet metadata."""
    try:
        tab_color = None
        if hasattr(sheet.sheet_properties, 'tabColor') and sheet.sheet_properties.tabColor:
            if hasattr(sheet.sheet_properties.tabColor, 'rgb'):
                rgb = sheet.sheet_properties.tabColor.rgb
                tab_color = rgb.decode() if isinstance(rgb, bytes) else str(rgb)

        return {
            'title': sheet.title,
            'dimensions': sheet.dimensions,
            'max_row': sheet.max_row,
            'max_column': sheet.max_column,
            'sheet_properties': {
                'tab_color': tab_color,
                'page_setup': {
                    'orientation': sheet.page_setup.orientation,
                    'paper_size': sheet.page_setup.paperSize
                } if hasattr(sheet, 'page_setup') else None
            }
        }
    except Exception as e:
        print_status(f"Error processing sheet metadata: {str(e)}", 'error')
        return {
            'title': sheet.title,
            'dimensions': sheet.dimensions,
            'max_row': sheet.max_row,
            'max_column': sheet.max_column
        }

def get_conditional_formatting(sheet: Worksheet) -> List[Dict[str, Any]]:
    """Extract conditional formatting rules from the sheet."""
    try:
        cf_rules = []
        for cf_range in sheet.conditional_formatting:
            for rule in cf_range.rules:
                cf_rules.append({
                    'range': str(cf_range),
                    'type': rule.type,
                    'priority': rule.priority,
                    'formula': rule.formula if hasattr(rule, 'formula') else None,
                    'dxf': {
                        'font': rule.dxf.font.__dict__ if rule.dxf and hasattr(rule.dxf, 'font') else None,
                        'fill': rule.dxf.fill.__dict__ if rule.dxf and hasattr(rule.dxf, 'fill') else None,
                    } if hasattr(rule, 'dxf') and rule.dxf else None
                })
        return cf_rules
    except Exception as e:
        print_status(f"Error processing conditional formatting: {str(e)}", 'error')
        return []

def get_named_ranges(workbook: Any) -> Dict[str, Any]:
    """Extract named ranges from the workbook."""
    try:
        named_ranges = {}
        if hasattr(workbook, 'defined_names'):
            # Handle different versions of openpyxl
            if hasattr(workbook.defined_names, 'items'):
                # New version of openpyxl
                for name, defn in workbook.defined_names.items():
                    try:
                        # Get the destinations for this named range
                        destinations = defn.destinations if hasattr(defn, 'destinations') else [(defn.attr_map.get('localSheetId'), defn.value)]
                        
                        # Process each destination
                        range_refs = []
                        for sheet_id, coord in destinations:
                            try:
                                # Get sheet name if sheet_id is provided
                                sheet_name = None
                                if sheet_id is not None:
                                    sheet_name = workbook.sheetnames[int(sheet_id)]
                                
                                # Clean up the reference
                                clean_ref = coord.strip('$')
                                if '!' not in clean_ref and sheet_name:
                                    clean_ref = f"{sheet_name}!{clean_ref}"
                                elif '!' not in clean_ref:
                                    # Use the first sheet if no sheet is specified
                                    clean_ref = f"{workbook.sheetnames[0]}!{clean_ref}"
                                
                                range_refs.append(clean_ref)
                            except Exception as e:
                                print_status(f"Warning: Could not process destination for named range '{name}': {str(e)}", 'warning')
                                continue
                        
                        if range_refs:
                            named_ranges[name] = {
                                'value': range_refs[0] if len(range_refs) == 1 else range_refs,
                                'scope': 'workbook' if defn.localSheetId is None else f"sheet_{defn.localSheetId}"
                            }
                    except Exception as e:
                        print_status(f"Warning: Could not process named range '{name}': {str(e)}", 'warning')
                        continue
            
            elif hasattr(workbook.defined_names, '_dict'):
                # Alternative version of openpyxl
                for name, defn in workbook.defined_names._dict.items():
                    try:
                        # Get the reference
                        ref = defn.value if hasattr(defn, 'value') else defn.attr_map.get('refersTo', '')
                        
                        # Clean up the reference
                        if ref.startswith('='):
                            ref = ref[1:]
                        ref = ref.strip('$')
                        
                        # Add sheet name if missing
                        if '!' not in ref:
                            sheet_id = defn.localSheetId if hasattr(defn, 'localSheetId') else None
                            sheet_name = workbook.sheetnames[int(sheet_id)] if sheet_id is not None else workbook.sheetnames[0]
                            ref = f"{sheet_name}!{ref}"
                        
                        named_ranges[name] = {
                            'value': ref,
                            'scope': 'workbook' if not hasattr(defn, 'localSheetId') or defn.localSheetId is None else f"sheet_{defn.localSheetId}"
                        }
                    except Exception as e:
                        print_status(f"Warning: Could not process named range '{name}': {str(e)}", 'warning')
                        continue
        
        return named_ranges
    except Exception as e:
        print_status(f"Warning: Could not process named ranges: {str(e)}", 'warning')
        return {}

def get_protection_settings(sheet: Worksheet) -> Dict[str, Any]:
    """Extract protection settings from the sheet."""
    try:
        protection_info = {}
        
        # Check if sheet has protection
        if hasattr(sheet, 'protection') and sheet.protection:
            protection = sheet.protection
            protection_info['sheet_protection'] = {
                'enabled': getattr(protection, 'sheet', False),
                'password': getattr(protection, 'password', None) is not None,
            }
            
            # Get available protection attributes
            for attr in dir(protection):
                if not attr.startswith('_') and attr not in ['sheet', 'password']:
                    protection_info['sheet_protection'][attr] = getattr(protection, attr, None)
        
        # Get protected ranges if available
        if hasattr(sheet, 'protected_ranges'):
            protection_info['protected_ranges'] = [
                str(protected_range) for protected_range in sheet.protected_ranges
            ]
        else:
            protection_info['protected_ranges'] = []
            
        return protection_info
    except Exception as e:
        print_status(f"Warning: Could not process protection settings: {str(e)}", 'warning')
        return {
            'sheet_protection': {'enabled': False},
            'protected_ranges': []
        }

def get_hyperlinks(cell: Any) -> Optional[Dict[str, Any]]:
    """Extract hyperlink information from a cell."""
    try:
        if cell.hyperlink:
            return {
                'target': cell.hyperlink.target,
                'tooltip': cell.hyperlink.tooltip if hasattr(cell.hyperlink, 'tooltip') else None,
                'location': cell.hyperlink.location if hasattr(cell.hyperlink, 'location') else None
            }
        return None
    except Exception as e:
        print_status(f"Error processing hyperlink: {str(e)}", 'error')
        return None

def get_comments(cell: Any) -> Optional[Dict[str, Any]]:
    """Extract comment information from a cell."""
    try:
        if cell.comment:
            return {
                'text': cell.comment.text,
                'author': cell.comment.author if hasattr(cell.comment, 'author') else None
            }
        return None
    except Exception as e:
        print_status(f"Error processing comment: {str(e)}", 'error')
        return None

def get_sheet_view_settings(sheet: Worksheet) -> Dict[str, Any]:
    """Extract sheet view settings."""
    try:
        return {
            'frozen_panes': {
                'rows': sheet.freeze_panes[1] if sheet.freeze_panes else None,
                'columns': sheet.freeze_panes[0] if sheet.freeze_panes else None
            },
            'zoom_scale': sheet.sheet_view.zoomScale if hasattr(sheet.sheet_view, 'zoomScale') else 100,
            'show_gridlines': sheet.sheet_view.showGridLines if hasattr(sheet.sheet_view, 'showGridLines') else True,
            'hidden_rows': [row for row in range(1, sheet.max_row + 1) if sheet.row_dimensions[row].hidden],
            'hidden_columns': [col for col in range(1, sheet.max_column + 1) if sheet.column_dimensions[get_column_letter(col)].hidden]
        }
    except Exception as e:
        print_status(f"Error processing sheet view settings: {str(e)}", 'error')
        return {}

def get_pivot_tables(sheet: Worksheet) -> List[Dict[str, Any]]:
    """Extract pivot table information from the sheet."""
    try:
        pivot_tables = []
        if hasattr(sheet, '_pivots'):
            for pivot in sheet._pivots:
                pivot_data = {
                    'location': str(pivot.location),
                    'name': pivot.name,
                    'cache_id': pivot.cache.cacheId if hasattr(pivot, 'cache') else None,
                    'fields': {
                        'row_fields': [field.name for field in pivot.row_fields] if hasattr(pivot, 'row_fields') else [],
                        'column_fields': [field.name for field in pivot.column_fields] if hasattr(pivot, 'column_fields') else [],
                        'data_fields': [{'name': field.name, 'function': field.function} for field in pivot.data_fields] if hasattr(pivot, 'data_fields') else [],
                        'page_fields': [field.name for field in pivot.page_fields] if hasattr(pivot, 'page_fields') else []
                    },
                    'options': {
                        'merge_labels': pivot.merge_labels if hasattr(pivot, 'merge_labels') else None,
                        'show_empty': pivot.show_empty if hasattr(pivot, 'show_empty') else None,
                        'indent': pivot.indent if hasattr(pivot, 'indent') else None
                    }
                }
                pivot_tables.append(pivot_data)
        return pivot_tables
    except Exception as e:
        print_status(f"Error processing pivot tables: {str(e)}", 'error')
        return []

def get_charts(sheet: Worksheet) -> List[Dict[str, Any]]:
    """Extract chart information from the sheet."""
    try:
        charts = []
        if hasattr(sheet, '_charts'):
            for chart in sheet._charts:
                chart_data = {
                    'type': type(chart).__name__,
                    'title': chart.title.text if hasattr(chart, 'title') and chart.title else None,
                    'anchor': str(chart.anchor) if hasattr(chart, 'anchor') else None,
                    'series': [{
                        'title': series.title if hasattr(series, 'title') else None,
                        'values': str(series.values) if hasattr(series, 'values') else None,
                        'categories': str(series.categories) if hasattr(series, 'categories') else None
                    } for series in chart.series] if hasattr(chart, 'series') else [],
                    'style': chart.style if hasattr(chart, 'style') else None,
                    'legend': {
                        'position': chart.legend.position if hasattr(chart, 'legend') and hasattr(chart.legend, 'position') else None,
                        'overlay': chart.legend.overlay if hasattr(chart, 'legend') and hasattr(chart.legend, 'overlay') else None
                    } if hasattr(chart, 'legend') else None
                }
                charts.append(chart_data)
        return charts
    except Exception as e:
        print_status(f"Error processing charts: {str(e)}", 'error')
        return []

def get_form_controls(sheet: Worksheet) -> List[Dict[str, Any]]:
    """Extract form control information from the sheet."""
    try:
        controls = []
        if hasattr(sheet, '_controls'):
            for control in sheet._controls:
                control_data = {
                    'type': type(control).__name__,
                    'name': control.name if hasattr(control, 'name') else None,
                    'location': str(control.anchor) if hasattr(control, 'anchor') else None,
                    'properties': {
                        'caption': control.caption if hasattr(control, 'caption') else None,
                        'value': control.value if hasattr(control, 'value') else None,
                        'linked_cell': str(control.linked_cell) if hasattr(control, 'linked_cell') else None,
                        'disabled': control.disabled if hasattr(control, 'disabled') else None,
                        'print_object': control.print_object if hasattr(control, 'print_object') else None,
                        'macro': control.macro if hasattr(control, 'macro') else None
                    }
                }
                controls.append(control_data)
        return controls
    except Exception as e:
        print_status(f"Error processing form controls: {str(e)}", 'error')
        return []

def get_auto_filters(sheet: Worksheet) -> Optional[Dict[str, Any]]:
    """Extract auto-filter information from the sheet."""
    try:
        if sheet.auto_filter.ref:
            filter_data = {
                'range': str(sheet.auto_filter.ref),
                'filters': {}
            }
            
            if hasattr(sheet.auto_filter, 'filterColumn'):
                for col_id, filter_column in sheet.auto_filter.filterColumn.items():
                    filter_data['filters'][col_id] = {
                        'type': filter_column.type if hasattr(filter_column, 'type') else None,
                        'values': filter_column.vals if hasattr(filter_column, 'vals') else None,
                        'custom_filters': filter_column.customFilters if hasattr(filter_column, 'customFilters') else None
                    }
            return filter_data
        return None
    except Exception as e:
        print_status(f"Error processing auto-filters: {str(e)}", 'error')
        return None

def get_cell_dependencies(cell: Any, sheet: Worksheet) -> Optional[Dict[str, Any]]:
    """Extract cell formula dependencies."""
    try:
        if cell.data_type == 'f':
            formula = str(cell.value)
            dependencies = {
                'precedents': [],  # Cells that this formula depends on
                'dependents': []   # Cells that depend on this cell
            }
            
            # Extract cell references from the formula
            cell_refs = re.findall(r'([A-Za-z]+[0-9]+(?::[A-Za-z]+[0-9]+)?)', formula)
            dependencies['precedents'] = [ref for ref in cell_refs if ':' not in ref]  # Exclude ranges
            
            # Find cells that reference this cell (dependents)
            current_cell_addr = cell.coordinate
            for row in sheet.iter_rows():
                for other_cell in row:
                    if (other_cell.data_type == 'f' and 
                        current_cell_addr in str(other_cell.value)):
                        dependencies['dependents'].append(other_cell.coordinate)
            
            return dependencies if dependencies['precedents'] or dependencies['dependents'] else None
        return None
    except Exception as e:
        print_status(f"Error processing cell dependencies: {str(e)}", 'error')
        return None

def convert_spreadsheet_to_json(file_path: str, sample_size: Optional[int] = None) -> Dict[str, Any]:
    """Convert spreadsheet to JSON structure with optional sample size limit."""
    print_status(f"Processing file: {file_path}", 'info')
    
    try:
        workbook = openpyxl.load_workbook(file_path, data_only=False)
        
        result = {
            'file_name': os.path.basename(file_path),
            'named_ranges': get_named_ranges(workbook),
            'sheets': {}
        }
        
        for sheet_name in workbook.sheetnames:
            print_status(f"Processing sheet: {sheet_name}", 'info')
            sheet = workbook[sheet_name]
            
            sheet_data = {
                'metadata': get_sheet_metadata(sheet),
                'cells': {},
                'merged_cells': [str(merged_range) for merged_range in sheet.merged_cells.ranges],
                'conditional_formatting': get_conditional_formatting(sheet),
                'protection': get_protection_settings(sheet),
                'view_settings': get_sheet_view_settings(sheet),
                'pivot_tables': get_pivot_tables(sheet),
                'charts': get_charts(sheet),
                'form_controls': get_form_controls(sheet),
                'auto_filter': get_auto_filters(sheet)
            }
            
            # Determine the number of rows to process
            max_rows = min(sheet.max_row, sample_size) if sample_size else sheet.max_row
            
            # Process cells up to sample_size if specified
            for row in range(1, max_rows + 1):
                for col in range(1, sheet.max_column + 1):
                    cell = sheet.cell(row=row, column=col)
                    cell_value = get_cell_value(cell)
                    
                    if cell_value is not None:
                        cell_address = f"{get_column_letter(col)}{row}"
                        sheet_data['cells'][cell_address] = {
                            'value': cell_value,
                            'style': {
                                'font': {
                                    'bold': cell.font.bold,
                                    'italic': cell.font.italic,
                                    'color': cell.font.color.rgb if cell.font.color else None
                                },
                                'fill': {
                                    'background': cell.fill.start_color.rgb if cell.fill and hasattr(cell.fill, 'start_color') else None
                                },
                                'alignment': {
                                    'horizontal': cell.alignment.horizontal,
                                    'vertical': cell.alignment.vertical
                                }
                            },
                            'data_validation': get_data_validation(cell),
                            'hyperlink': get_hyperlinks(cell),
                            'comment': get_comments(cell),
                            'dependencies': get_cell_dependencies(cell, sheet)
                        }
            
            result['sheets'][sheet_name] = sheet_data
        
        return result
    
    except Exception as e:
        print_status(f"Error converting spreadsheet: {str(e)}", 'error')
        raise

def count_json_tokens(data: Dict[str, Any]) -> int:
    """
    Count the number of tokens in JSON data.
    This is an approximation that tries to match typical LLM tokenization patterns.
    """
    try:
        # Convert to JSON string using the custom encoder
        json_str = json.dumps(data, cls=ExcelJSONEncoder)
        
        # Basic tokenization rules that approximate LLM tokenizers:
        # 1. Split on whitespace
        # 2. Split on punctuation
        # 3. Split on case changes (camelCase, PascalCase)
        # 4. Keep common JSON structural elements as single tokens
        
        # First, handle JSON structural elements
        json_str = re.sub(r'([\{\}\[\],:])', r' \1 ', json_str)
        
        # Split on camelCase and PascalCase
        json_str = re.sub(r'([a-z])([A-Z])', r'\1 \2', json_str)
        json_str = re.sub(r'([A-Z])([A-Z][a-z])', r'\1 \2', json_str)
        
        # Handle numbers and special characters
        json_str = re.sub(r'([0-9]+)', r' \1 ', json_str)
        json_str = re.sub(r'([^a-zA-Z0-9\s\{\}\[\],:])', r' \1 ', json_str)
        
        # Split and filter empty strings
        tokens = [token for token in json_str.split() if token.strip()]
        
        # Add a small overhead for JSON structure and formatting
        structural_overhead = len(re.findall(r'[\{\}\[\],:]', json_str)) * 0.5
        
        # Return total token count with overhead
        return int(len(tokens) + structural_overhead)
        
    except Exception as e:
        print_status(f"Error counting tokens: {str(e)}", 'error')
        return 0

def save_json_output(data: Dict[str, Any], original_file_path: str) -> str:
    """Save JSON output to file with token count in filename."""
    try:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        
        # Count tokens
        token_count = count_json_tokens(data)
        print_status(f"Total tokens in output: {token_count}", 'info')
        
        # Generate filename with token count and timestamp
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = os.path.splitext(os.path.basename(original_file_path))[0]
        output_filename = f"{base_name}_{token_count}tokens_{timestamp}.json"
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        
        # Add metadata including token count
        if 'metadata' not in data:
            data['metadata'] = {}
        
        data['metadata'].update({
            'token_count': token_count,
            'conversion_timestamp': timestamp,
            'original_filename': os.path.basename(original_file_path)
        })
        
        # Save the file using the custom encoder
        with open(output_path, 'w', encoding=DEFAULT_ENCODING) as f:
            json.dump(data, f, indent=2, ensure_ascii=False, cls=ExcelJSONEncoder)
        
        print_status(f"Successfully saved JSON to: {output_path} ({token_count} tokens)", 'success')
        return output_path
    
    except Exception as e:
        print_status(f"Error saving JSON output: {str(e)}", 'error')
        raise

def main():
    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print_status("Usage: python spreadsheet_converter.py <path_to_spreadsheet> [sample_size]", 'error')
        print_status("  sample_size: Optional. Number of rows to process per sheet", 'info')
        sys.exit(1)
    
    file_path = sys.argv[1]
    
    # Get sample size if provided
    sample_size = None
    if len(sys.argv) == 3:
        try:
            sample_size = int(sys.argv[2])
            if sample_size <= 0:
                print_status("Sample size must be a positive integer", 'error')
                sys.exit(1)
        except ValueError:
            print_status("Sample size must be a valid integer", 'error')
            sys.exit(1)
    
    if not os.path.exists(file_path):
        print_status(f"File not found: {file_path}", 'error')
        sys.exit(1)
    
    if not any(file_path.lower().endswith(ext) for ext in SUPPORTED_EXTENSIONS):
        print_status(f"Unsupported file format. Supported formats: {', '.join(SUPPORTED_EXTENSIONS)}", 'error')
        sys.exit(1)
    
    try:
        print_status("Converting spreadsheet to JSON...", 'info')
        if sample_size:
            print_status(f"Processing {sample_size} rows per sheet", 'info')
        
        json_data = convert_spreadsheet_to_json(file_path, sample_size)
        
        # Save with token count in filename
        output_path = save_json_output(json_data, file_path)
        
        # Get token count from saved data
        token_count = json_data['metadata']['token_count']
        
        # Print LLM recommendations
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
        
        print_status("\nConversion completed successfully!", 'success')
    
    except Exception as e:
        print_status(f"Conversion failed: {str(e)}", 'error')
        sys.exit(1)

if __name__ == "__main__":
    main() 