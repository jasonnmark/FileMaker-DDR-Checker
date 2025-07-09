import sys
import subprocess
import platform
import os
import json
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from collections import defaultdict
import re
from lxml import etree as ET
from openpyxl.styles import Font, PatternFill, Alignment
import pandas as pd
import importlib.util
import pickle

# Check for debug mode and cache mode early
DEBUG_MODE = '-debug' in sys.argv or '--debug' in sys.argv
CACHE_MODE = '-cache' in sys.argv or '--cache' in sys.argv

# Define standard colors and styles that check modules can use
STANDARD_COLORS = {
    'error': {'fill': 'FFEBEE', 'font': 'B71C1C'},  # Red
    'error_strong': {'fill': 'FFCDD2', 'font': 'B71C1C'},  # Darker red
    'warning': {'fill': 'FFF9C4', 'font': 'F57C00'},  # Yellow/Orange
    'success': {'fill': 'C8E6C9', 'font': '2E7D32'},  # Green
    'info': {'fill': 'E3F2FD', 'font': '1976D2'},  # Blue
    'muted': {'fill': 'F5F5F5', 'font': '888888'},  # Grey
    'header': {'fill': '2D3142', 'font': 'FFFFFF'},  # Dark header
    
    # Category colors for SQL sheets
    'category_script': {'fill': 'B7E4C7', 'font': '000000'},
    'category_custom_function': {'fill': 'A9DEF9', 'font': '000000'},
    'category_field_calc': {'fill': 'FFE066', 'font': '000000'},
    'category_layout_object': {'fill': 'FFB4A2', 'font': '000000'},
    'category_other': {'fill': 'D3D3D3', 'font': '000000'},
}

def get_cache_dir():
    """Get the Cache directory path"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    cache_dir = os.path.join(script_dir, "Cache")
    if not os.path.exists(cache_dir):
        os.makedirs(cache_dir)
    return cache_dir

def get_exports_dir():
    """Get the Exports directory path"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    exports_dir = os.path.join(script_dir, "Exports")
    if not os.path.exists(exports_dir):
        os.makedirs(exports_dir)
    return exports_dir

def get_cache_filename():
    """Get the single cache filename"""
    return "ddr_normalized.cache"

def save_to_cache(normalized_xml, input_file, base_output_name):
    """Save normalized XML to cache"""
    cache_dir = get_cache_dir()
    cache_file = os.path.join(cache_dir, get_cache_filename())
    
    # Save both the normalized XML and metadata
    cache_data = {
        'normalized_xml': normalized_xml,
        'input_file': input_file,
        'base_output_name': base_output_name,
        'file_modified_time': os.path.getmtime(input_file)
    }
    
    try:
        with open(cache_file, 'wb') as f:
            pickle.dump(cache_data, f)
        print(f"✓ Cached normalized data")
        return True
    except Exception as e:
        print(f"Warning: Could not save cache: {e}")
        return False

def load_from_cache():
    """Load normalized XML from cache if it exists"""
    cache_dir = get_cache_dir()
    cache_file = os.path.join(cache_dir, get_cache_filename())
    
    if not os.path.exists(cache_file):
        return None
    
    try:
        with open(cache_file, 'rb') as f:
            cache_data = pickle.load(f)
        
        print(f"✓ Found cache")
        return cache_data
    except Exception as e:
        print(f"✗ Could not load cache: {e}")
        return None

def get_color_fill(color_name):
    """Get a PatternFill object for a standard color"""
    if color_name in STANDARD_COLORS:
        color = STANDARD_COLORS[color_name]['fill']
        return PatternFill(start_color=color, end_color=color, fill_type="solid")
    return None

def get_color_font(color_name, bold=False, strike=False):
    """Get a Font object for a standard color"""
    if color_name in STANDARD_COLORS:
        color = STANDARD_COLORS[color_name]['font']
        return Font(color=color, bold=bold, strike=strike)
    return Font(bold=bold, strike=strike)

def check_and_install(package):
    # Handle special case where package name differs from import name
    import_name = package
    if package == 'pyahocorasick':
        import_name = 'ahocorasick'
    
    try:
        __import__(import_name)
        print(f"✓ {package} found")
    except ImportError:
        print(f"❌ {package} not found")
        response = input(f"Install {package} now? (y/n): ").lower().strip()
        if response in ['y', 'yes']:
            print(f"Installing {package}...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                print(f"✓ {package} installed successfully!")
            except Exception as e:
                print(f"❌ Installation failed: {e}")
                sys.exit(1)
        else:
            print(f"❌ Cannot continue without {package}")
            sys.exit(1)

print("=== System Check ===")

def check_python_version():
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 7):
        print(f"❌ Python {version.major}.{version.minor} detected.")
        print("This script requires Python 3.7 or higher.")
        print("Please upgrade your Python installation.")
        return False
    else:
        print(f"✓ Python {version.major}.{version.minor}.{version.micro} - Compatible")
        return True

if not check_python_version():
    sys.exit(1)

for dep in ['pandas', 'openpyxl', 'lxml', 'pyahocorasick']:
    check_and_install(dep)

parser = ET.XMLParser(remove_blank_text=True, recover=True)
print("✓ Ready to process DDR files\n")

def replace_emojis_with_plus(text):
    """Replace all emojis in text with + character"""
    import unicodedata
    
    # Count emojis before replacement (for debug mode)
    emoji_count = 0
    
    result = []
    i = 0
    while i < len(text):
        char = text[i]
        
        # Check if this character is an emoji or part of an emoji sequence
        is_emoji = False
        
        # Method 1: Check Unicode category
        category = unicodedata.category(char)
        if category in ['So', 'Cn']:  # Symbol, Other or Unassigned
            # Additional check for emoji properties
            code_point = ord(char)
            
            # Check various emoji ranges
            if (0x1F300 <= code_point <= 0x1F9FF or  # Emoticons, pictographs, etc.
                0x1F000 <= code_point <= 0x1F2FF or  # Mahjong, dominoes, cards
                0x1FA00 <= code_point <= 0x1FAFF or  # Chess symbols, new emojis
                0x2600 <= code_point <= 0x27BF or    # Misc symbols, dingbats  
                0x2300 <= code_point <= 0x23FF or    # Misc technical
                0x2B00 <= code_point <= 0x2BFF or    # Misc symbols and arrows
                0x1F1E6 <= code_point <= 0x1F1FF or  # Regional indicators (flags)
                code_point in [0x200D, 0xFE0F] or     # ZWJ and variation selector
                0x1F3FB <= code_point <= 0x1F3FF or  # Skin tone modifiers
                0xE0020 <= code_point <= 0xE007F or  # Tag characters
                0x2000 <= code_point <= 0x206F or    # General punctuation (includes some emoji)
                0x20D0 <= code_point <= 0x20FF or    # Combining marks
                0x3000 <= code_point <= 0x303F or    # CJK symbols (includes some emoji)
                0xFE00 <= code_point <= 0xFE0F or    # Variation selectors
                0x1F000 <= code_point <= 0x1FFFF or  # All SMP emoji blocks
                0xE0000 <= code_point <= 0xE01EF):   # Tags
                is_emoji = True
        
        # Method 2: Check for specific emoji characters that might be missed
        elif ord(char) > 127:  # Non-ASCII
            # Check if it's a standalone emoji that might not be in 'So' category
            code_point = ord(char)
            if (code_point >= 0x231A or  # From watch emoji onwards
                char in '©®™℗⚠♻☠☢☣⚡☎☏✉✈🎯🎪🎨🎬🎤🎧🎼🎵🎶🎹🎻🎺🎷🎸🎮🃏🎴🀄🎲🎰🎳'):
                is_emoji = True
        
        # Check for emoji sequences (multi-codepoint emojis)
        sequence_length = 1
        if i + 1 < len(text):
            next_char = text[i + 1]
            next_code = ord(next_char)
            
            # Check for variation selector or zero-width joiner
            if next_code in [0xFE0F, 0xFE0E, 0x200D]:
                is_emoji = True
                sequence_length = 2
                
                # Check for longer sequences (e.g., family emojis, profession emojis)
                j = i + 2
                while j < len(text) and j < i + 10:  # Limit sequence check
                    following_code = ord(text[j])
                    if following_code in [0xFE0F, 0xFE0E, 0x200D] or 0x1F000 <= following_code <= 0x1FFFF:
                        sequence_length = j - i + 1
                        j += 1
                    else:
                        break
            
            # Check for flag sequences (regional indicators)
            elif 0x1F1E6 <= ord(char) <= 0x1F1FF and 0x1F1E6 <= next_code <= 0x1F1FF:
                is_emoji = True
                sequence_length = 2
            
            # Check for keycap sequences (e.g., 1️⃣)
            elif char in '0123456789*#' and next_code == 0xFE0F and i + 2 < len(text) and ord(text[i + 2]) == 0x20E3:
                is_emoji = True
                sequence_length = 3
            elif char in '0123456789*#' and next_code == 0x20E3:
                is_emoji = True
                sequence_length = 2
        
        if is_emoji:
            result.append('+')
            i += sequence_length
            emoji_count += 1
        else:
            result.append(char)
            i += 1
    
    # Final pass: catch any remaining suspicious unicode that might be emoji
    final_text = ''.join(result)
    
    # Pattern for any remaining high unicode characters that are likely emojis
    remaining_emoji_pattern = re.compile(r'[\U00010000-\U0010FFFF]+', flags=re.UNICODE)
    
    # Check each match to see if it's likely an emoji
    def check_and_replace(match):
        text = match.group(0)
        # If it's in the emoji ranges, replace it
        for char in text:
            code = ord(char)
            if 0x1F000 <= code <= 0x1FFFF:
                return '+'
        return text
    
    final_text = remaining_emoji_pattern.sub(check_and_replace, final_text)
    
    if DEBUG_MODE and emoji_count > 0:
        print(f"[DEBUG] Replaced {emoji_count} emoji sequences with +")
    
    return final_text

def load_check_module(module_name, file_path):
    """Dynamically load a check module from file path"""
    try:
        spec = importlib.util.spec_from_file_location(module_name, file_path)
        module = importlib.util.module_from_spec(spec)
        
        # Inject the color helpers into the module before executing
        module.STANDARD_COLORS = STANDARD_COLORS
        module.get_color_fill = get_color_fill
        module.get_color_font = get_color_font
        
        spec.loader.exec_module(module)
        return module
    except Exception as e:
        print(f"Error loading check module {module_name}: {e}")
        return None

def normalize_xml(raw_xml):
    """Normalize XML content - extracted from parse_ddr for reuse"""
    # Replace all emojis with + BEFORE any other processing
    raw_xml = replace_emojis_with_plus(raw_xml)
    
    count_nwea  = raw_xml.count("??")
    print(f'Found {count_nwea} occurrences of "??" in raw_xml')

    # Replace all double question marks (was "🤯."), this is because Filemaker DDR doesn't always export emoji's correctly and it causes parsing problems.
    # Now this should be redundant since we're replacing all emojis, but keeping for backwards compatibility
    raw_xml = raw_xml.replace("??", "+")
    
    # Check if the problematic pattern exists BEFORE any modifications
    if DEBUG_MODE:
        if "</DisplayCalculation>UpdateSingleStudentOnServer" in raw_xml:
            print("[DEBUG] Found </DisplayCalculation>UpdateSingleStudentOnServer BEFORE normalization")
        else:
            print("[DEBUG] Did NOT find </DisplayCalculation>UpdateSingleStudentOnServer BEFORE normalization")
    
    # Replace smart quotes with regular quotes to normalize script references
    # Using Unicode escapes to avoid encoding issues
    raw_xml = raw_xml.replace("\u201C", '"')  # U+201C " Left double quotation mark
    raw_xml = raw_xml.replace("\u201D", '"')  # U+201D " Right double quotation mark
    raw_xml = raw_xml.replace("\u201E", '"')  # U+201E „ Double low-9 quotation mark
    raw_xml = raw_xml.replace("\u201F", '"')  # U+201F ‟ Double high-reversed-9 quotation mark
    raw_xml = raw_xml.replace("\u2018", "'")  # U+2018 ' Left single quotation mark
    raw_xml = raw_xml.replace("\u2019", "'")  # U+2019 ' Right single quotation mark
    
    # Fix self-closing tags to ensure proper parsing
    # Add space before /> to prevent issues with script name detection
    raw_xml = raw_xml.replace('"/>', '" />')
    raw_xml = raw_xml.replace("'/>", "' />")
    
    # Debug: Check if the problematic pattern exists
    if DEBUG_MODE:
        import re
        problem_pattern = re.findall(r'</DisplayCalculation>[A-Za-z]+', raw_xml)
        if problem_pattern:
            print(f"[DEBUG] Found {len(problem_pattern)} instances of content directly after </DisplayCalculation>")
            for pattern in problem_pattern[:3]:
                print(f"[DEBUG]   Example: {pattern}")
    
    # Note: Not fixing the </DisplayCalculation>ScriptName pattern
    # This unusual XML structure needs to be preserved for accurate counting
    
    # Also add space after semicolons in quotes for better parsing
    raw_xml = raw_xml.replace('";', '" ;')
    raw_xml = raw_xml.replace("';", "' ;")
    
    count_plus = raw_xml.count("+")
    print(f"raw contains '+': {count_plus}")
    
    return raw_xml

def run_checks(raw_xml, base_output_name):
    """Run all checks - extracted from parse_ddr for reuse"""
    # Load and run all checks from Checks folder
    all_sheets = {}
    sheet_orders = {}  # Track sheet orders
    checks_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Checks")
    
    if os.path.exists(checks_dir):
        print("\nRunning checks...")
        # First, collect all check modules
        check_modules = []
        
        # Scan for all .py files in Checks folder
        for filename in sorted(os.listdir(checks_dir)):
            if filename.endswith('.py') and not filename.startswith('__'):
                check_path = os.path.join(checks_dir, filename)
                module_name = filename[:-3]  # Remove .py extension
                
                check_module = load_check_module(module_name, check_path)
                if check_module and hasattr(check_module, 'run_check'):
                    # Get sheet order if defined
                    sheet_order = None
                    if hasattr(check_module, 'get_sheet_order'):
                        try:
                            sheet_order = check_module.get_sheet_order()
                        except:
                            sheet_order = None
                    
                    check_modules.append({
                        'module': check_module,
                        'name': module_name,
                        'path': check_path,
                        'order': float(sheet_order) if sheet_order is not None else float('inf')
                    })
        
        # Sort modules by order, then by name
        check_modules.sort(key=lambda x: (x['order'], x['name']))
        
        # Run checks in sorted order
        total_results = 0
        for check_info in check_modules:
            check_module = check_info['module']
            module_name = check_info['name']
            
            print(f"Running {module_name}...")
            
            try:
                # Pass raw_xml to the check (with normalized quotes)
                check_results = check_module.run_check(raw_xml)
                if check_results and hasattr(check_module, 'get_sheet_name'):
                    sheet_name = check_module.get_sheet_name()
                else:
                    # Default sheet name based on module name
                    sheet_name = module_name.replace('_', ' ').replace('Check', '').strip()
                
                if check_results:
                    all_sheets[sheet_name] = {
                        'data': check_results,
                        'module': check_module
                    }
                    sheet_orders[sheet_name] = check_info['order']
                    total_results += len(check_results)
                    print(f"✓ {module_name} completed with {len(check_results)} results")
                else:
                    print(f"  {module_name} returned no results")
            except Exception as e:
                print(f"Error running {module_name}: {e}")
                import traceback
                traceback.print_exc()
        
        # Clear terminal before final output if not in debug mode
        if DEBUG_MODE:
            print("\n[DEBUG] Skipping terminal clear due to debug mode")
            print("=== FileMaker DDR Analysis Complete (Debug Mode) ===")
            print(f"Total issues found: {total_results}")
            print("Debug mode: All output preserved above")
            print()
        else:
            os.system('cls' if os.name == 'nt' else 'clear')
            print("=== FileMaker DDR Analysis Complete ===")
            print(f"Total issues found: {total_results}")
            print()
        
        # Generate output with all sheets
        if all_sheets:
            generate_output_files(all_sheets, base_output_name, total_results, sheet_orders)
        else:
            print("\nNo results found from any checks.")
    else:
        print(f"\nChecks folder not found at: {checks_dir}")
        print("Please create a 'Checks' folder and add check modules.")

def parse_ddr(xml_file, base_output_name):
    try:
        # Check if we should use cache
        cache_data = None
        if CACHE_MODE:
            cache_data = load_from_cache()
            # Verify it's for the same file
            if cache_data and cache_data.get('input_file') != xml_file:
                print("✗ Cache is for a different file, ignoring cache")
                cache_data = None
            elif cache_data and os.path.exists(xml_file):
                # Check if the file has been modified since caching
                if os.path.getmtime(xml_file) != cache_data.get('file_modified_time'):
                    print("✗ Cache invalid: source file has been modified")
                    cache_data = None
        
        if cache_data:
            # Use cached data
            raw_xml = cache_data['normalized_xml']
            print("✓ Using cached normalized XML")
        else:
            # --- Read & normalize XML with better encoding handling ---
            encodings_to_try = ['utf-8', 'utf-8-sig', 'utf-16', 'utf-16le', 'utf-16be']
            # Detect BOM-based UTF-16 and read accordingly
            with open(xml_file, 'rb') as __f:
                __bom = __f.read(2)
            if __bom in (b'\xff\xfe', b'\xfe\xff'):
                # File is UTF-16, read in one shot and skip further encoding tries
                with open(xml_file, 'r', encoding='utf-16', errors='replace') as __f:
                    raw_xml = __f.read()
                print("Successfully read file with utf-16 BOM detection")
            else:
                raw_xml = None

            if raw_xml is None:
                for encoding in encodings_to_try:
                    try:
                        with open(xml_file, 'r', encoding=encoding, errors='replace') as f:
                            raw_xml = f.read()
                        print(f"Successfully read file with {encoding} encoding")
                        break
                    except UnicodeDecodeError:
                        continue
            
            if raw_xml is None:
                print(f"❌ Failed to read '{xml_file}' with any of the tried encodings.")
                sys.exit(1)

            # Normalize the XML
            raw_xml = normalize_xml(raw_xml)
            
            # Save to cache for future runs
            save_to_cache(raw_xml, xml_file, base_output_name)

        # Run checks
        run_checks(raw_xml, base_output_name)

    except ET.ParseError as e:
        print(f"Error parsing XML: {e}")
    except FileNotFoundError:
        print(f"File not found: {xml_file}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

def generate_output_files(all_sheets, base_output_name, total_found, sheet_orders=None):
    """Write results to Excel with formatting."""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    # Output path in Exports folder
    exports_dir = get_exports_dir()
    output_path = os.path.join(exports_dir, f"{os.path.basename(base_output_name)}_full_report.xlsx")

    # Fix duplicate sheet orders by adjusting them
    if sheet_orders:
        # Find duplicate order values and resolve them
        order_values = list(sheet_orders.values())
        seen_orders = set()
        duplicate_orders = [order for order in order_values if order in seen_orders or seen_orders.add(order)]

        if duplicate_orders and DEBUG_MODE:
            print(f"[DEBUG] Found {len(duplicate_orders)} duplicate sheet order values: {duplicate_orders}")
        
        # Fix duplicate orders by adding increments of 10
        if duplicate_orders:
            # Keep track of orders we've seen and adjusted
            fixed_orders = {}
            increment = 10
            
            for sheet_name, order in sheet_orders.items():
                if order in fixed_orders:
                    # This order has already been seen and adjusted for another sheet
                    # Use the next increment (10, 20, 30, etc.)
                    new_order = order + increment
                    sheet_orders[sheet_name] = new_order
                    increment += 10
                    
                    if DEBUG_MODE:
                        print(f"[DEBUG] Adjusted order for '{sheet_name}' from {order} to {new_order}")
                else:
                    # First time seeing this order, record it
                    fixed_orders[order] = True
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Sort sheets by order, then by name
        if sheet_orders:
            sorted_sheets = sorted(all_sheets.items(), 
                                 key=lambda x: (sheet_orders.get(x[0], float('inf')), x[0]))
        else:
            sorted_sheets = sorted(all_sheets.items())
        
        # Verify that each sheet is unique (debug only)
        if DEBUG_MODE:
            sheet_names = [name for name, _ in sorted_sheets]
            if len(sheet_names) != len(set(sheet_names)):
                print(f"[DEBUG] WARNING: Duplicate sheet names found: {sheet_names}")
            print(f"[DEBUG] Writing {len(sorted_sheets)} sheets in order:")
            for idx, (name, _) in enumerate(sorted_sheets):
                order = sheet_orders.get(name, float('inf'))
                print(f"[DEBUG]   {idx+1}. '{name}' (order: {order})")
        
        for sheet_name, sheet_info in sorted_sheets:
            sheet_data = sheet_info['data']
            if isinstance(sheet_data, pd.DataFrame):
                sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                # Convert list of dicts to DataFrame
                pd.DataFrame(sheet_data).to_excel(writer, sheet_name=sheet_name, index=False)

    # Styling
    wb = load_workbook(output_path)
    
    # Verify all sheets are present (debug only)
    if DEBUG_MODE:
        expected_sheets = [name for name, _ in sorted_sheets]
        actual_sheets = wb.sheetnames
        print(f"[DEBUG] Expected {len(expected_sheets)} sheets, found {len(actual_sheets)} sheets")
        if set(expected_sheets) != set(actual_sheets):
            missing = set(expected_sheets) - set(actual_sheets)
            extra = set(actual_sheets) - set(expected_sheets)
            if missing:
                print(f"[DEBUG] Missing sheets: {missing}")
            if extra:
                print(f"[DEBUG] Unexpected sheets: {extra}")
    
    header_fill = get_color_fill('header')
    header_font = get_color_font('header', bold=True)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.freeze_panes = ws['A2']
        
        # Row height
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            ws.row_dimensions[row[0].row].height = 30
            
        # Wrap & alignment
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                
        # Column widths - get from check module if available
        column_widths = {}
        check_module = None
        
        # Find the check module for this sheet
        if sheet_name in all_sheets:
            check_module = all_sheets[sheet_name]['module']
            if hasattr(check_module, 'get_column_widths'):
                column_widths = check_module.get_column_widths()
        
        # Apply column widths
        for idx, cell in enumerate(ws[1], 1):
            val = cell.value
            col = get_column_letter(idx)
            
            if val in column_widths:
                ws.column_dimensions[col].width = column_widths[val]/7
            elif val == "SQL Text":
                ws.column_dimensions[col].width = 40
            elif val == "":
                ws.column_dimensions[col].width = 3
            else:
                ws.column_dimensions[col].width = 20
                
        # Header style
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            
        # Apply sheet-specific styling from check module
        if check_module and hasattr(check_module, 'apply_styling'):
            check_module.apply_styling(ws)

    wb.save(output_path)

    # Auto-open the file after saving
    auto_open_file(output_path)

    print(f"Full report saved to: {os.path.abspath(output_path)}")
    
    print(f"\nRun complete: {total_found} results found.")

def auto_open_file(filename):
    try:
        if DEBUG_MODE:
            print(f"[DEBUG] Opening file: {filename}")
        
        if platform.system() == "Darwin":
            subprocess.run(["open", filename])
        elif platform.system() == "Windows":
            os.startfile(filename)
        elif platform.system() == "Linux":
            subprocess.run(["xdg-open", filename])
    except Exception as e:
        if DEBUG_MODE:
            print(f"[DEBUG] Error opening file: {e}")
        pass

def get_input_file():
    """Handle file selection"""
    # If cache mode, try to use cached data
    if CACHE_MODE:
        cache_data = load_from_cache()
        if cache_data:
            input_file = cache_data.get('input_file')
            # Check if original file still exists
            if input_file and os.path.exists(input_file):
                print(f"Using cached data for: {os.path.basename(input_file)}")
                return input_file
            else:
                print("✗ Original file from cache not found")
                print("Please select the DDR file to process")
    
    # Show file picker
    Tk().withdraw()
    input_file = askopenfilename(
        title="Select DDR XML File",
        filetypes=[("XML Files", "*.xml")]
    )
    
    if not input_file:
        print("No file selected. Exiting.")
        sys.exit(0)
    
    return input_file

if __name__ == "__main__":
    # Display mode information
    if CACHE_MODE:
        print("Cache mode enabled - will use cached data if available")
    if DEBUG_MODE:
        print("Debug mode enabled - additional output will be shown")
    
    # Get input file
    input_file = get_input_file()
    
    print(f"\nProcessing: {os.path.basename(input_file)}")

    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_filename = os.path.splitext(os.path.basename(input_file))[0]
    base_name = os.path.join(script_dir, f"{input_filename}_parsed")

    print("Starting enhanced DDR parse...")
    parse_ddr(input_file, base_name)