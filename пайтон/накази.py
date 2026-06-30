import os
import sys
import subprocess
import importlib
import re
from datetime import datetime, timedelta
from pathlib import Path
import time  # Import time for potential delays if needed
import argparse  # Add this import
import traceback
from contextlib import nullcontext
import io
import winsound

# Auto-install missing dependencies
def install_and_import(module_name, package_name=None):
    if package_name is None:
        package_name = module_name
    try:
        importlib.import_module(module_name)
    except ImportError:
        print(f"Missing required dependency '{package_name}'. Attempting to install automatically...", flush=True)
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
            importlib.import_module(module_name)
            print(f"Successfully installed {package_name}!", flush=True)
        except Exception as e:
            print(f"Error installing {package_name} automatically: {e}", file=sys.stderr, flush=True)
            print(f"Please install it manually using: pip install {package_name}", file=sys.stderr, flush=True)
            sys.exit(1)

install_and_import('pandas')
install_and_import('openpyxl')
install_and_import('win32com.client', 'pywin32')
install_and_import('PyPDF2')
install_and_import('pdfplumber')
install_and_import('reportlab')

# Import external/installed dependencies
import pandas as pd
import win32com.client
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import pdfplumber
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- Make console output crash-proof against characters the active code page
#     (e.g. cp1251 on Windows) cannot encode, such as '→'. Keep the existing
#     encoding but replace un-encodable characters instead of raising. ---
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors='replace')
    except Exception:
        pass

# --- Configuration ---
SOUND_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sound")
COMPLETE_SOUND = os.path.join(SOUND_DIR, "complete.wav")

# --- Dynamic Root & Paths Configuration ---
_SCRIPT_DIR = Path(__file__).resolve().parent
if _SCRIPT_DIR.name.lower() == 'пайтон':
    ROOT_DIR = _SCRIPT_DIR.parent
else:
    ROOT_DIR = _SCRIPT_DIR

EXCEL_FILE_PATH = None  # Resolved dynamically in get_excel_file_path() / main()
SHEET_NAME = 'Дебетовий'  # Default sheet name, resolved dynamically upon loading
TABLE_NAME = 'debitlist'
OUTPUT_ROOT_PATH = ROOT_DIR / "накази"
TEMPLATE_DIR = ROOT_DIR / "шаблони"

def get_template_path(filename):
    for path in ROOT_DIR.rglob(filename):
        parts = [p.lower() for p in path.parts]
        if path.is_file() and not any(p.startswith('.') or p in ('venv', 'env', 'build', 'dist', 'node_modules', '__pycache__') for p in parts):
            return path
    return TEMPLATE_DIR / filename

# --- Column Name Mapping (from Excel Table to internal variable names) ---
# Updated to match headers in new target file
COLUMN_MAP = {
    'date_start': 'навантаження дата', # Used for 'Date', 'start', and filename date
    'driver_full': 'водій',          # Used for 'driver' placeholder and initials
    'plate_number': 'авто',           # Used for 'plate' placeholder
    'date_end': 'розвантаження дата', # Used for 'end' placeholder
    'route_desc': 'маршрут згідно товаро-транспортній накладній', # Used for 'route'
    'company_code': 'товариство',    # ASSUMING this column exists in your table for 'BB'
    'comment': 'накази та цпх',      # Parameter column to check for service agreements (цпх)
    'freight': 'вантаж (родовий відмінок)', # Freight/cargo column for service agreements
    'customer': 'замовник',          # Customer column to detect АВТОСТРАДА rows (per-day service docs)
}

# --- Word Placeholders ---
PLACEHOLDERS = {
    "Date": "date_start",
    "driver": "driver_full",
    "truck": "truck_model",
    "plate": "plate_number",
    "start": "date_start",
    "end": "date_end",
    "route": "route_desc"
    # "Sequence" and "shortdrname" are handled separately
}

# --- Service Template Placeholders (different format with ${} syntax) ---
SERVICE_PLACEHOLDERS = {
    "${date}": "date_start",
    "${driver}": "driver_full", 
    "${route-contract}": "route_contract",
    "${start}": "date_start",
    "${end}": "date_end",
    "${freight}": "freight_contract",
    "${money}": "money_formatted"
    # "${sequence}", "${shortdname}", "${individualnumber}" are handled separately
    # Certificate-specific placeholders are handled in duplicate_certificate_pages():
    # "${certificate-sequence}", "${certificate-start}", "${certificate-end}", 
    # "${certificate-freight}", "${certificate-route}", "${certificate-money}"
}

# --- Template File Names ---
TEMPLATES = {
    "ЗЕТТРА": "шаблон-наказу-зет.docx",
    "ЗІАВТОТРАНС": "шаблон-наказу-зіа.docx",
    # Add more if needed, ensure keys match CompanyCode values EXACTLY
}

# --- Service Agreement Templates ---
SERVICE_AGREEMENT_TEMPLATES = {
    "ЗЕТТРА": "шаблон-договору-зет.docx",
    "ЗІАВТОТРАНС": "шаблон-договору-зіа.docx",
    # Add more if needed, ensure keys match CompanyCode values EXACTLY
}

# --- Company Short Code Mapping ---
COMPANY_SHORT_CODES = {
    "zet": "ЗЕТТРА",
    "zia": "ЗІАВТОТРАНС",
}

# Word constants
WD_REPLACE_ALL = 2
WD_FORMAT_PDF = 17
WD_DO_NOT_SAVE_CHANGES = 0

# --- PDF Processing Configuration ---
REMOVE_BLANK_PAGES = True  # Set to False to disable blank page removal from PDFs

# --- Logging and Error Handling Support ---
class Tee(object):
    def __init__(self, *files):
        self.files = files
    def write(self, obj):
        for f in self.files:
            try:
                f.write(obj)
                f.flush()
            except Exception:
                pass
    def flush(self):
        for f in self.files:
            try:
                f.flush()
            except Exception:
                pass
    def __getattr__(self, attr):
        return getattr(self.files[0], attr)

log_file = None

def setup_logging():
    global log_file
    try:
        logs_dir = ROOT_DIR / "logs"
        os.makedirs(logs_dir, exist_ok=True)
        log_filename = logs_dir / f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        log_file = open(log_filename, "w", encoding="utf-8", errors="replace")
        sys.stdout = Tee(sys.__stdout__ or sys.stdout, log_file)
        sys.stderr = Tee(sys.__stderr__ or sys.stderr, log_file)
    except Exception as e:
        print(f"Warning: Could not set up file logging: {e}")

import atexit
def close_logging():
    global log_file
    if log_file:
        try:
            sys.stdout = sys.__stdout__ or sys.stdout
            sys.stderr = sys.__stderr__ or sys.stderr
            log_file.close()
        except:
            pass

atexit.register(close_logging)

def custom_excepthook(exctype, value, tb):
    import traceback
    if issubclass(exctype, KeyboardInterrupt):
        print("\nВиконання перервано користувачем.")
        close_logging()
        os._exit(0)
    traceback.print_exception(exctype, value, tb, file=sys.stderr)
    print("\nВиникла критична помилка. Деталі збережено в лог-файлі (папка logs).")
    try:
        input("Натисніть Enter для виходу...")
    except (KeyboardInterrupt, EOFError):
        pass
    close_logging()
    os._exit(1)

original_exit = sys.exit
def custom_exit(code=0):
    if code != 0:
        print(f"\nСкрипт завершився з кодом помилки: {code}. Деталі збережено в лог-файлі (папка logs).")
        try:
            input("Натисніть Enter для виходу...")
        except (KeyboardInterrupt, EOFError):
            pass
    close_logging()
    original_exit(code)

sys.excepthook = custom_excepthook
sys.exit = custom_exit

def parse_month_arg(month_str):
    """
    Parse month argument in format MMYY (e.g., '0525' for May 2025)
    Returns tuple (year, month) or None if invalid
    """
    if not month_str or len(month_str) != 4:
        return None
    
    try:
        month = int(month_str[:2])
        year = int(month_str[2:]) + 2000  # Convert YY to 20YY
        
        if month < 1 or month > 12:
            return None
            
        return (year, month)
    except ValueError:
        return None

def get_sheet_name(excel_path):
    """
    Dynamically detect the appropriate sheet name from the Excel workbook.
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(excel_path, read_only=True)
        names = wb.sheetnames
        wb.close()
        for preferred in ['Дебетовий', 'Дебетовий список']:
            if preferred in names:
                return preferred
        return names[0] if names else 'Дебетовий'
    except Exception:
        return 'Дебетовий'

def get_excel_file_path(month_str=None):
    """
    Dynamically find the target Excel spreadsheet in ROOT_DIR.
    Excludes temporary files starting with '~$'.
    """
    excel_extensions = ('.xlsx', '.xlsm', '.xls')
    files = [p for p in ROOT_DIR.iterdir() if p.is_file() and p.suffix.lower() in excel_extensions and not p.name.startswith('~$')]
    
    if not files:
        raise FileNotFoundError(f"No Excel files found in root directory: {ROOT_DIR}")
        
    if month_str:
        for f in files:
            if month_str in f.name:
                return str(f)
                
    return str(files[0])

def get_previous_month():
    """
    Get the previous month's year and month
    Returns tuple (year, month)
    """
    today = datetime.now()
    first_day_this_month = today.replace(day=1)
    last_day_prev_month = first_day_this_month - timedelta(days=1)
    return (last_day_prev_month.year, last_day_prev_month.month)

def format_date_ukrainian(date_str):
    """
    Format date from DD.MM.YYYY to 'DD month_name YYYY' with Ukrainian month name.
    Example: '15.11.2025' -> '15 листопада 2025'
    """
    ukrainian_months = {
        1: 'січня',
        2: 'лютого',
        3: 'березня',
        4: 'квітня',
        5: 'травня',
        6: 'червня',
        7: 'липня',
        8: 'серпня',
        9: 'вересня',
        10: 'жовтня',
        11: 'листопада',
        12: 'грудня'
    }
    
    try:
        # Parse the date string
        if isinstance(date_str, str):
            # Assuming format DD.MM.YYYY
            parts = date_str.split('.')
            if len(parts) == 3:
                day = int(parts[0])
                month = int(parts[1])
                year = parts[2]
                
                month_name = ukrainian_months.get(month, '')
                return f"{day} {month_name} {year}"
        
        return date_str  # Return as-is if parsing fails
    except Exception as e:
        print(f"Warning: Could not format date '{date_str}': {e}")
        return date_str

def format_driver_initials(full_name):
    """Formats driver name cleanly to 'Lastname I. O.' or handles full names to 'Lastname I. O.'."""
    if not isinstance(full_name, str) or not full_name.strip():
        return "NoName" # Handle empty or non-string names
    
    cleaned = full_name.strip(" \t\r\n.,")
    # Match standard initials format: e.g., 'Боіштян О.О.', 'Гладир А. М.', 'СоколенкоВ.С.', 'Бабак Р.М'
    # Group 1: Lastname (ends with a lowercase letter)
    # Group 2: First initial (uppercase)
    # Group 3: Patronymic initial (uppercase, optional)
    regex = re.compile(r'^([A-ZА-ЯІЇЄҐЁa-zа-яіїєґё\'\-’`]+?[a-zа-яіїєґё])\s*([A-ZА-ЯІЇЄҐЁ])\.?\s*([A-ZА-ЯІЇЄҐЁ]?)\.?$')
    m = regex.match(cleaned)
    if m:
        last = m.group(1)
        i1 = m.group(2)
        i2 = m.group(3)
        if i2:
            return f"{last} {i1}. {i2}."
        return f"{last} {i1}."
    
    # Fallback to original/general parsing for full names (e.g. 'Боіштян Олександр Олександрович')
    parts = cleaned.split()
    if len(parts) < 2:
        return cleaned  # Return as is if less than two parts
    
    initials = parts[0]  # Last name
    for part in parts[1:]:
        if part:  # Ensure part is not empty
            initials += f" {part[0]}."
    return initials

def is_service_agreement(param_text):
    """Check if the parameter contains service agreement pattern, e.g., 'цпх зіа 2000 0608251535' or old '${sa;...}'"""
    if not param_text or pd.isna(param_text):
        return False
    
    param_str = str(param_text).strip().lower()
    
    # Old format check
    if '${sa;' in param_str:
        pattern = r'\$\{sa;[^;]+;[^;]+;[^}]+\}'
        return bool(re.search(pattern, param_str))
        
    # New format check: starts with 'цпх' and has at least 3 space-separated parts following it
    parts = param_str.split()
    if len(parts) >= 4 and parts[0] == 'цпх':
        company = parts[1]
        if company in ('зіа', 'зет', 'zia', 'zet', 'зіавтотранс', 'зеттра'):
            return True
            
    return False

def is_exclude_param(param_text):
    """Check whether the parameter cell contains ${exclude} or 'виключити'."""
    if not param_text or pd.isna(param_text):
        return False
    s = str(param_text).strip().lower()
    return '${exclude}' in s or 'виключити' in s

def is_avtostrada_row(customer_value):
    """Check if row belongs to АВТОСТРАДА customer.
    These rows are grouped by day and create per-day service docs instead of assignments."""
    if not customer_value or pd.isna(customer_value):
        return False
    return str(customer_value).strip().upper() == 'АВТОСТРАДА'

def extract_service_agreement_details(param_text):
    """
    Extract service agreement details from parameter.
    Supports old format: ${sa;company;money;timestamp}
    Supports new format: цпх company money timestamp (e.g. цпх зіа 2000 0608251535)
    Returns {'company': company, 'money': money, 'timestamp': timestamp} or None.
    """
    if not param_text or pd.isna(param_text):
        return None
    
    param_str = str(param_text).strip()
    
    # 1. Try old format
    pattern = r'\$\{sa;([^;]+);([^;]+);([^}]+)\}'
    match = re.search(pattern, param_str)
    if match:
        company = match.group(1).strip().upper()
        if company in ('ЗІА', 'ЗІАВТОТРАНС'):
            company = 'ZIA'
        elif company in ('ЗЕТ', 'ЗЕТТРА'):
            company = 'ZET'
        return {'company': company, 'money': match.group(2).strip(), 'timestamp': match.group(3).strip()}
        
    # 2. Try new format: цпх company money timestamp
    parts = param_str.split()
    if len(parts) >= 4 and parts[0].lower() == 'цпх':
        company = parts[1].upper()
        # Normalize Ukrainian company codes to English equivalents (ZIA/ZET)
        if company in ('ЗІА', 'ЗІАВТОТРАНС'):
            company = 'ZIA'
        elif company in ('ЗЕТ', 'ЗЕТТРА'):
            company = 'ZET'
        
        money = parts[2]
        timestamp = parts[3]
        return {'company': company, 'money': money, 'timestamp': timestamp}
        
    return None

def join_day_items(items):
    """Collapse the routes/freights of a single certificate (one АВТОСТРАДА day,
    which may contain several trips) into one comma-separated string.
    Duplicates within the day are removed, order preserved.
    Used so each certificate maps to exactly one numbered contract item."""
    cleaned = []
    seen = set()
    for x in items:
        s = str(x).strip()
        if s and s not in seen:
            cleaned.append(s)
            seen.add(s)
    return ", ".join(cleaned)

def format_freight_contract(freights_list):
    """
    Format a list of freights into numbered format for service agreements.
    Example: ['Grain', 'Wheat', 'Corn'] -> 'Grain; Wheat; Corn'
    Always shows actual freight names, including duplicates.
    """
    if not freights_list:
        return ""
    
    # Include all freights, even duplicates
    all_freights = []
    for freight in freights_list:
        if freight and str(freight).strip():  # Only add non-empty freights
            all_freights.append(str(freight).strip())
    
    if len(all_freights) == 1:
        result = all_freights[0]
    else:
        # Number the freight names: 1. Freight1; 2. Freight2; etc.
        numbered_freights = []
        for i, freight in enumerate(all_freights, 1):
            numbered_freights.append(f"{i}. {freight}")
        result = "; ".join(numbered_freights)
    
    # No length cap: find_replace() chunks long text (see find_replace_chunked),
    # so the full list is inserted even when it exceeds Word's ~255-char limit.
    return result

def format_routes_contract(routes_list):
    """
    Format a list of routes into numbered format for service agreements.
    Example: ['Route A', 'Route B', 'Route C'] -> 'Route A; Route B; Route C'
    Always shows actual route names, including duplicates.
    """
    if not routes_list:
        return ""
    
    # Include all routes, even duplicates
    all_routes = []
    for route in routes_list:
        if route and str(route).strip():  # Only add non-empty routes
            all_routes.append(str(route).strip())
    
    if len(all_routes) == 1:
        result = all_routes[0]
    else:
        # Number the route names: 1. Route1; 2. Route2; etc.
        numbered_routes = []
        for i, route in enumerate(all_routes, 1):
            numbered_routes.append(f"{i}. {route}")
        result = "; ".join(numbered_routes)
    
    # No length cap: find_replace() chunks long text (see find_replace_chunked),
    # so the full list is inserted even when it exceeds Word's ~255-char limit.
    return result

def format_money_ukrainian(money_str):
    """
    Convert money amount to Ukrainian format: X гривень Y коп.
    Example: '8000' -> '8000 гривень 00 коп.'
    Example: '8000.50' -> '8000 гривень 50 коп.'
    """
    try:
        # Convert to float to handle decimal values
        amount = float(money_str)
        
        # Split into hryvnias and kopecks
        hryvnias = int(amount)
        kopecks = int((amount - hryvnias) * 100)
        
        return f"{hryvnias} гривень {kopecks:02d} коп."
        
    except (ValueError, TypeError):
        print(f"Warning: Invalid money format '{money_str}', using as-is")
        return str(money_str)

def number_to_ukrainian_words(number):
    """
    Convert number to Ukrainian words.
    Example: 2000 -> 'Дві тисячі'
    """
    ones = ['', 'один', 'два', 'три', 'чотири', "п'ять", 'шість', 'сім', 'вісім', "дев'ять"]
    ones_feminine = ['', 'одна', 'дві', 'три', 'чотири', "п'ять", 'шість', 'сім', 'вісім', "дев'ять"]
    teens = ['десять', 'одинадцять', 'дванадцять', 'тринадцять', 'чотирнадцять', "п'ятнадцять", 
             'шістнадцять', 'сімнадцять', 'вісімнадцять', "дев'ятнадцять"]
    tens = ['', '', 'двадцять', 'тридцять', 'сорок', "п'ятдесят", 'шістдесят', 'сімдесят', 'вісімдесят', "дев'яносто"]
    hundreds = ['', 'сто', 'двісті', 'триста', 'чотириста', "п'ятсот", 'шістсот', 'сімсот', 'вісімсот', "дев'ятсот"]
    
    if number == 0:
        return 'нуль'
    
    if number >= 1000000:
        return str(number)  # For very large numbers, just return digits
    
    result = []
    
    # Millions (not implemented for now, return digits)
    if number >= 1000000:
        return str(number)
    
    # Thousands
    if number >= 1000:
        thousands = number // 1000
        if thousands >= 100:
            result.append(hundreds[thousands // 100])
        
        remainder = thousands % 100
        if remainder >= 20:
            result.append(tens[remainder // 10])
            if remainder % 10 > 0:
                if remainder % 10 == 1 or remainder % 10 == 2:
                    result.append(ones_feminine[remainder % 10])  # тисяча/тисячі - feminine
                else:
                    result.append(ones[remainder % 10])
        elif remainder >= 10:
            result.append(teens[remainder - 10])
        elif remainder > 0:
            if remainder == 1 or remainder == 2:
                result.append(ones_feminine[remainder])  # тисяча/тисячі - feminine
            else:
                result.append(ones[remainder])
        
        # Add thousand word
        if thousands % 10 == 1 and thousands % 100 != 11:
            result.append('тисяча')
        elif thousands % 10 in [2, 3, 4] and thousands % 100 not in [12, 13, 14]:
            result.append('тисячі')
        else:
            result.append('тисяч')
    
    # Hundreds
    remainder = number % 1000
    if remainder >= 100:
        result.append(hundreds[remainder // 100])
    
    # Tens and ones
    remainder = remainder % 100
    if remainder >= 20:
        result.append(tens[remainder // 10])
        if remainder % 10 > 0:
            # Use feminine form for 1 and 2 when referring to гривні
            if remainder % 10 == 1 or remainder % 10 == 2:
                result.append(ones_feminine[remainder % 10])
            else:
                result.append(ones[remainder % 10])
    elif remainder >= 10:
        result.append(teens[remainder - 10])
    elif remainder > 0:
        # Use feminine form for 1 and 2 when referring to гривні
        if remainder == 1 or remainder == 2:
            result.append(ones_feminine[remainder])
        else:
            result.append(ones[remainder])
    
    return ' '.join(result).capitalize()

def format_certificate_money_ukrainian(total_money_str, num_rows):
    """
    Convert money amount divided by number of rows to Ukrainian format with words.
    Example: total_money='8000', num_rows=4 -> '2000 (Дві тисячі) гривні 00 коп.'
    """
    try:
        # Convert to float and divide by number of rows
        total_amount = float(total_money_str)
        amount_per_certificate = total_amount / num_rows
        
        # Split into hryvnias and kopecks
        hryvnias = int(amount_per_certificate)
        kopecks = int((amount_per_certificate - hryvnias) * 100)
        
        # Convert hryvnias to words
        hryvnias_words = number_to_ukrainian_words(hryvnias)
        
        # Format the result
        return f"{hryvnias} ({hryvnias_words}) гривні {kopecks:02d} коп."
        
    except (ValueError, TypeError):
        print(f"Warning: Invalid money format '{total_money_str}', using as-is")
        return str(total_money_str)

def extract_driver_from_parameter(param_text):
    """
    Extract driver name from parameter if it contains водій {Name} pattern (or legacy ${driver;Name}).
    Example: водій {Бондаренко Сергій Валерійович} -> returns 'Бондаренко Сергій Валерійович'
    Returns None if pattern not found.
    """
    if not param_text or pd.isna(param_text):
        return None
    
    param_str = str(param_text).strip()
    
    # Pattern to match водій {Name}
    pattern_new = r'водій\s*\{([^}]+)\}'
    match_new = re.search(pattern_new, param_str, re.IGNORECASE)
    if match_new:
        driver_name = match_new.group(1).strip()
        return driver_name if driver_name else None
        
    # Legacy pattern to match ${driver;Name}
    pattern_old = r'\$\{driver;([^}]+)\}'
    match_old = re.search(pattern_old, param_str)
    if match_old:
        driver_name = match_old.group(1).strip()
        return driver_name if driver_name else None
    
    return None

def get_driver_individual_number(excel_file_path, driver_name):
    """
    Look up driver's individual number (РНОКПП) from the "Водії" sheet.
    Returns the РНОКПП value for the given driver name, or empty string if not found.
    """
    try:
        # Read the drivers sheet
        drivers_df = pd.read_excel(excel_file_path, sheet_name='Водії', engine='openpyxl')
        
        # Find the driver row (case-insensitive comparison)
        driver_row = drivers_df[drivers_df['водій'].str.strip().str.upper() == driver_name.strip().upper()]
        
        if not driver_row.empty:
            # Get the РНОКПП value - note the space after РНОКПП
            individual_number = driver_row['РНОКПП '].iloc[0]
            if pd.isna(individual_number):
                print(f"Warning: РНОКПП is empty for driver '{driver_name}'")
                return ""
            
            # Convert to string and remove .0 if it's a whole number
            result = str(individual_number).strip()
            if result.endswith('.0'):
                result = result[:-2]
            return result
        else:
            print(f"Warning: Driver '{driver_name}' not found in 'Водії' sheet")
            return ""
            
    except Exception as e:
        print(f"Warning: Error looking up individual number for driver '{driver_name}': {e}")
        return ""

def _cert_replace_one(word_doc, placeholder, replacement_text):
    """Replace the FIRST occurrence of placeholder, handling text longer than Word's 255-char limit."""
    rep_str = str(replacement_text)
    if len(rep_str) <= 250:
        find_obj = word_doc.Content.Find
        find_obj.ClearFormatting()
        find_obj.Replacement.ClearFormatting()
        find_obj.Text = placeholder
        find_obj.Replacement.Text = rep_str
        find_obj.Forward = True
        find_obj.Wrap = 1
        find_obj.Format = False
        find_obj.MatchCase = False
        find_obj.MatchWholeWord = True
        find_obj.MatchWildcards = False
        find_obj.MatchSoundsLike = False
        find_obj.MatchAllWordForms = False
        return find_obj.Execute(Replace=1)
    # Long text: locate the placeholder first, then overwrite the range directly
    find_obj = word_doc.Content.Find
    find_obj.ClearFormatting()
    find_obj.Text = placeholder
    find_obj.Forward = True
    find_obj.Wrap = 1
    find_obj.Format = False
    find_obj.MatchCase = False
    find_obj.MatchWholeWord = True
    find_obj.MatchWildcards = False
    find_obj.MatchSoundsLike = False
    find_obj.MatchAllWordForms = False
    if not find_obj.Execute():
        return False
    found_range = find_obj.Parent
    chunks = [rep_str[i:i + 200] for i in range(0, len(rep_str), 200)]
    found_range.Text = chunks[0]
    ip = found_range.End
    for chunk in chunks[1:]:
        r = word_doc.Range(ip, ip)
        r.Text = chunk
        ip = r.End
    return True


def duplicate_certificate_pages(word_doc, num_rows, service_agreement_data):
    """
    Duplicate the 3rd page (certificate page) for service agreements with multiple rows.
    Each duplicated page gets a certificate sequence number (1, 2, 3, etc.) and row-specific data.
    
    Args:
        word_doc: Word document COM object
        num_rows: Number of rows in the service agreement group
        service_agreement_data: Dictionary containing service agreement data with row-specific info
        
    Returns:
        True if duplication was successful, False otherwise
    """
    try:
        if num_rows <= 1:
            # Single row - no duplication needed, just replace certificate placeholders with row data
            row_data = service_agreement_data.get('row_data', [])
            total_money = service_agreement_data.get('money', '0')
            
            # Get data for the single row
            current_row_data = row_data[0] if row_data else {}
            
            # Prepare replacement data for the single certificate
            certificate_replacements = {
                "${certificate-sequence}": "1",
                "${certificate-start}": current_row_data.get('date_start', ''),
                "${certificate-end}": current_row_data.get('date_end', ''),
                "${certificate-freight}": current_row_data.get('freight', ''),
                "${certificate-route}": current_row_data.get('route_desc', ''),
                "${certificate-money}": format_certificate_money_ukrainian(total_money, 1)
            }
            
            print("Single certificate replacements:")
            for placeholder, value in certificate_replacements.items():
                print(f"  {placeholder} -> {value}")
            
            # Replace all certificate placeholders
            for placeholder, replacement_value in certificate_replacements.items():
                find_replace(word_doc, placeholder, str(replacement_value))
            
            return True
            
        # Check if document has at least 3 pages (need certificate page to duplicate)
        total_pages = word_doc.ComputeStatistics(2)  # 2 = wdStatisticPages
        if total_pages < 3:
            print(f"Warning: Document has only {total_pages} pages, cannot duplicate certificate page")
            # Still replace certificate placeholders for the single certificate
            row_data = service_agreement_data.get('row_data', [])
            total_money = service_agreement_data.get('money', '0')
            current_row_data = row_data[0] if row_data else {}
            
            certificate_replacements = {
                "${certificate-sequence}": "1",
                "${certificate-start}": current_row_data.get('date_start', ''),
                "${certificate-end}": current_row_data.get('date_end', ''),
                "${certificate-freight}": current_row_data.get('freight', ''),
                "${certificate-route}": current_row_data.get('route_desc', ''),
                "${certificate-money}": format_certificate_money_ukrainian(total_money, 1)
            }
            
            for placeholder, replacement_value in certificate_replacements.items():
                find_replace(word_doc, placeholder, str(replacement_value))
            
            return False
            
        print(f"Duplicating certificate page {num_rows - 1} times for {num_rows} total certificate pages")
        
        # Get the Word application object from the document
        word_app = word_doc.Application
        
        # Force document to update/repaginate before working with pages
        word_doc.Repaginate()
        time.sleep(0.1)  # Small delay to let Word process
        
        # Recalculate page count after repagination
        total_pages_after_repaginate = word_doc.ComputeStatistics(2)
        print(f"Pages after repagination: {total_pages_after_repaginate}")
        
        # More robust approach to find page 3 content
        # First, go to the beginning of the document
        word_app.Selection.HomeKey(6)  # 6 = wdStory (beginning of document)
        
        # Navigate to start of page 3 more safely
        try:
            print("Attempting to navigate to page 3...")
            word_app.Selection.GoTo(1, 2, 3)  # 1=wdGoToPage, 2=wdGoToAbsolute, page 3
            start_of_page_3 = word_app.Selection.Start
            print(f"Start of page 3 position: {start_of_page_3}")
            
            # Move to end of page 3 by going to start of next page or end of document
            if total_pages_after_repaginate > 3:
                print("Document has more than 3 pages, navigating to page 4...")
                word_app.Selection.GoTo(1, 2, 4)  # Go to page 4
                end_of_page_3 = word_app.Selection.Start - 1
                print(f"End of page 3 position (before page 4): {end_of_page_3}")
            else:
                # If only 3 pages, go to end of document
                print("Document has exactly 3 pages, going to end of document...")
                word_app.Selection.EndKey(6)  # Go to end of document
                end_of_page_3 = word_app.Selection.End
                print(f"End of page 3 position (end of doc): {end_of_page_3}")
            
            # Validate the range values
            if start_of_page_3 < 0 or end_of_page_3 < 0 or start_of_page_3 >= end_of_page_3:
                raise ValueError(f"Invalid range: start={start_of_page_3}, end={end_of_page_3}")
            
            print(f"Creating range from {start_of_page_3} to {end_of_page_3}")
            
            # Create and validate the range
            page_3_range = word_doc.Range(start_of_page_3, end_of_page_3)
            
            # Select and copy the page content
            page_3_range.Select()
            page_3_range.Copy()  # Copy to clipboard with all formatting
            print("Successfully copied page 3 content")
            
        except Exception as range_error:
            print(f"Error creating page 3 range: {range_error}")
            # Fallback: try to select page 3 using a different method
            try:
                print("Trying fallback method...")
                # Alternative method: select from page 3 to end of page 3
                word_app.Selection.GoTo(1, 2, 3)  # Go to page 3
                word_app.Selection.MoveDown(5, 1000, True)  # Select many lines down
                page_3_range = word_app.Selection.Range
                page_3_range.Copy()
                print("Used fallback method to copy page 3 content")
            except Exception as fallback_error:
                print(f"Fallback method also failed: {fallback_error}")
                raise range_error  # Re-raise original error
        
        # Position cursor at the end of the document
        word_app.Selection.EndKey(6)  # Go to end of document
        
        # Duplicate the page for each additional row (rows 2, 3, 4, etc.)
        for cert_seq in range(2, num_rows + 1):
            # Insert page break
            word_app.Selection.InsertBreak(7)  # 7 = wdPageBreak
            
            # Paste the copied content with all formatting
            word_app.Selection.Paste()
            
        print(f"Created {num_rows - 1} additional certificate pages with preserved formatting")
        
        # Now replace certificate-specific placeholders for each certificate page
        # Get row-specific data from service agreement
        row_data = service_agreement_data.get('row_data', [])
        total_money = service_agreement_data.get('money', '0')
        
        for cert_seq in range(1, num_rows + 1):
            # Get data for this specific row (certificate)
            row_index = cert_seq - 1  # Convert to 0-based index
            if row_index < len(row_data):
                current_row_data = row_data[row_index]
            else:
                # Fallback to first row data if we don't have enough rows
                current_row_data = row_data[0] if row_data else {}
            
            # Prepare replacement data for this certificate
            certificate_replacements = {
                "${certificate-sequence}": str(cert_seq),
                "${certificate-start}": current_row_data.get('date_start', ''),
                "${certificate-end}": current_row_data.get('date_end', ''),
                "${certificate-freight}": current_row_data.get('freight', ''),
                "${certificate-route}": current_row_data.get('route_desc', ''),
                "${certificate-money}": format_certificate_money_ukrainian(total_money, num_rows)
            }
            
            print(f"Certificate {cert_seq} replacements:")
            for placeholder, value in certificate_replacements.items():
                print(f"  {placeholder} -> {value}")
            
            # Replace placeholders one by one for this certificate
            for placeholder, replacement_value in certificate_replacements.items():
                if _cert_replace_one(word_doc, placeholder, replacement_value):
                    print(f"Replaced {placeholder} with {str(replacement_value)[:80]}")
                else:
                    print(f"Warning: Could not find {placeholder} placeholder for sequence {cert_seq}")
        
        return True
        
    except Exception as e:
        print(f"Error duplicating certificate pages: {e}")
        import traceback
        print(traceback.format_exc())
        
        # Enhanced fallback - replace certificate placeholders with basic values
        try:
            print("Attempting fallback: replacing certificate placeholders without duplication...")
            row_data = service_agreement_data.get('row_data', []) if service_agreement_data else []
            total_money = service_agreement_data.get('money', '0') if service_agreement_data else '0'
            
            # For multiple rows, we'll use the first row data for all certificates
            # This is not ideal but better than failing completely
            if row_data:
                current_row_data = row_data[0]
                print(f"Using first row data for fallback: {current_row_data.get('date_start', 'N/A')}")
            else:
                current_row_data = {}
                print("No row data available for fallback")
            
            certificate_replacements = {
                "${certificate-sequence}": "1",
                "${certificate-start}": current_row_data.get('date_start', ''),
                "${certificate-end}": current_row_data.get('date_end', ''),
                "${certificate-freight}": current_row_data.get('freight', ''),
                "${certificate-route}": current_row_data.get('route_desc', ''),
                "${certificate-money}": format_certificate_money_ukrainian(total_money, num_rows)
            }
            
            print("Fallback certificate replacements:")
            for placeholder, value in certificate_replacements.items():
                print(f"  {placeholder} -> {value}")
            
            for placeholder, replacement_value in certificate_replacements.items():
                if find_replace(word_doc, placeholder, str(replacement_value)):
                    print(f"Fallback: Replaced {placeholder}")
                else:
                    print(f"Fallback: Could not replace {placeholder}")
                    
            print("Fallback certificate replacement completed")
            
        except Exception as fallback_error:
            print(f"Fallback also failed: {fallback_error}")
            print(traceback.format_exc())
            
        return False

def get_global_sequence_number(output_path, company_code, target_year, target_month, is_service_agreement=False, timestamp=None):
    """
    Finds the highest sequence number across ALL files for a given company/month
    in the specified folder and returns the next number.
    This ensures unique sequential numbering across all drivers for the target month.
    Filename format: CompanyCode DriverInitials від dd.mm.yyyy нSequenceNumber.docx (assignments)
                     CompanyCode DriverInitials від dd.mm.yyyy цпхSequenceNumber.pdf (service agreements)
    """
    max_seq_num = 0
    
    # Pattern to match files based on type
    if is_service_agreement:
        # Pattern for service agreements with цпх prefix
        pattern_str = rf"^{re.escape(company_code)}\s+.+\s+від\s+(\d{{2}})\.(\d{{2}})\.(\d{{4}})\s+цпх(\d+)\.pdf$"
    else:
        # Pattern for regular assignments with н prefix
        pattern_str = rf"^{re.escape(company_code)}\s+.+\s+від\s+(\d{{2}})\.(\d{{2}})\.(\d{{4}})\s+н(\d+)\.docx$"
    
    pattern = re.compile(pattern_str, re.IGNORECASE)

    if not output_path.exists():
        return 1 # Folder doesn't exist, start with 1

    try:
        for filename in os.listdir(output_path):
            match = pattern.match(filename)
            if match:
                file_day = int(match.group(1))
                file_month = int(match.group(2))
                file_year = int(match.group(3))
                seq_num = int(match.group(4))
                
                # Check if this file is from the target month/year
                if file_year == target_year and file_month == target_month:
                    if seq_num > max_seq_num:
                        max_seq_num = seq_num
                        
    except Exception as e:
        print(f"Warning: Error reading sequence numbers from {output_path}: {e}")

    return max_seq_num + 1

def get_company_service_sequence_number(company_code):
    """
    Finds the highest service agreement sequence number for a specific company
    in the company-specific ЦПХ-pdf folder and returns the next number.
    This ensures unique sequential numbering for service agreements per company.
    """
    max_seq_num = 0
    
    # Pattern for service agreements with цпх prefix for specific company
    pattern_str = rf"^{re.escape(company_code)}\s+.+\s+від\s+(\d{{2}})\.(\d{{2}})\.(\d{{4}})\s+цпх(\d+)\.pdf$"
    pattern = re.compile(pattern_str, re.IGNORECASE)

    # Check the company-specific ЦПХ-pdf folder
    company_service_pdf_path = OUTPUT_ROOT_PATH / f"{company_code}-ЦПХ-pdf"
    
    if not company_service_pdf_path.exists():
        return 1  # Folder doesn't exist, start with 1

    try:
        for filename in os.listdir(company_service_pdf_path):
            match = pattern.match(filename)
            if match:
                seq_num = int(match.group(4))
                if seq_num > max_seq_num:
                    max_seq_num = seq_num
                        
    except Exception as e:
        print(f"Warning: Error scanning for service sequence numbers in {company_service_pdf_path}: {e}")

    return max_seq_num + 1

def find_replace(doc, find_text, replace_text):
    """Performs a find and replace operation in the Word document's content, headers, and footers."""
    replace_text_str = str(replace_text)
    
    # Word's Find/Replace has a character limit (~255 chars for replacement text)
    # If replacement text is too long, break it into chunks
    if len(replace_text_str) > 250:
        return find_replace_chunked(doc, find_text, replace_text_str)
    
    # Use normal Find/Replace for shorter strings
    try:
        replacements_made = False
        
        # 1. Replace in main document content
        find_obj = doc.Content.Find
        find_obj.ClearFormatting()
        find_obj.Replacement.ClearFormatting()
        find_obj.Text = find_text
        find_obj.Replacement.Text = replace_text_str
        find_obj.Forward = True
        find_obj.Wrap = 1  # wdFindContinue
        find_obj.Format = False
        find_obj.MatchCase = False
        find_obj.MatchWholeWord = True
        find_obj.MatchWildcards = False
        find_obj.MatchSoundsLike = False
        find_obj.MatchAllWordForms = False
        find_obj.Execute(Replace=WD_REPLACE_ALL)
        replacements_made = True
        
        # 2. Replace in all headers (each section can have different headers)
        for section in doc.Sections:
            # Primary header
            if section.Headers(1).Range.Text:  # 1 = wdHeaderFooterPrimary
                header_find = section.Headers(1).Range.Find
                header_find.ClearFormatting()
                header_find.Replacement.ClearFormatting()
                header_find.Text = find_text
                header_find.Replacement.Text = replace_text_str
                header_find.Forward = True
                header_find.Wrap = 1
                header_find.Format = False
                header_find.MatchCase = False
                header_find.MatchWholeWord = True
                header_find.MatchWildcards = False
                header_find.Execute(Replace=WD_REPLACE_ALL)
                replacements_made = True
        
        # 3. Replace in all footers
        for section in doc.Sections:
            # Primary footer
            if section.Footers(1).Range.Text:  # 1 = wdHeaderFooterPrimary
                footer_find = section.Footers(1).Range.Find
                footer_find.ClearFormatting()
                footer_find.Replacement.ClearFormatting()
                footer_find.Text = find_text
                footer_find.Replacement.Text = replace_text_str
                footer_find.Forward = True
                footer_find.Wrap = 1
                footer_find.Format = False
                footer_find.MatchCase = False
                footer_find.MatchWholeWord = True
                footer_find.MatchWildcards = False
                footer_find.Execute(Replace=WD_REPLACE_ALL)
                replacements_made = True
        
        return replacements_made
        
    except Exception as e:
        print(f"Warning: Find/Replace failed for '{find_text}': {e}")
        import traceback
        print(traceback.format_exc())
        return find_replace_chunked(doc, find_text, replace_text_str)

def find_replace_chunked(doc, find_text, replace_text):
    """Break long text into chunks and insert them piece by piece."""
    try:
        # Find the placeholder first
        find_obj = doc.Content.Find
        find_obj.ClearFormatting()
        
        find_obj.Text = find_text
        find_obj.Forward = True
        find_obj.Wrap = 1  # wdFindContinue
        find_obj.Format = False
        find_obj.MatchCase = False
        find_obj.MatchWholeWord = True
        find_obj.MatchWildcards = False
        find_obj.MatchSoundsLike = False
        find_obj.MatchAllWordForms = False
        
        replacements_made = 0
        chunk_size = 200  # Safe chunk size
        
        # Process each occurrence of the placeholder
        while find_obj.Execute():
            # Get the found range
            found_range = find_obj.Parent
            
            # Break the replacement text into chunks
            chunks = []
            for i in range(0, len(replace_text), chunk_size):
                chunks.append(replace_text[i:i + chunk_size])
            
            if chunks:
                # Replace the placeholder with the first chunk
                found_range.Text = chunks[0]
                
                # Insert remaining chunks after the first one
                insertion_point = found_range.End
                for chunk in chunks[1:]:
                    # Create a range at the insertion point and insert the chunk
                    insert_range = doc.Range(insertion_point, insertion_point)
                    insert_range.Text = chunk
                    insertion_point = insert_range.End
                
                replacements_made += 1
                print(f"Chunked replacement: '{find_text}' -> {len(replace_text)} chars in {len(chunks)} chunks")
            
            # Continue searching from after the replacement
            remaining_start = insertion_point if chunks else found_range.End
            if remaining_start >= doc.Content.End:
                break
                
            # Create new range for remaining text and continue searching
            remaining_range = doc.Range(remaining_start, doc.Content.End)
            find_obj = remaining_range.Find
            find_obj.ClearFormatting()
            find_obj.Text = find_text
            find_obj.Forward = True
            find_obj.Wrap = 1
            find_obj.Format = False
            find_obj.MatchCase = False
            find_obj.MatchWholeWord = True
            find_obj.MatchWildcards = False
            find_obj.MatchSoundsLike = False
            find_obj.MatchAllWordForms = False
        
        return replacements_made > 0
        
    except Exception as e:
        print(f"Error in chunked replacement for '{find_text}': {e}")
        return False

def is_content_page(pdf_path, page_num):
    """Determine if a page contains meaningful content (text or images) in the body area.
    Header/footer text is excluded by checking only the middle portion of the page."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_num]
            page_height = page.height
            
            # Define header/footer margins (points from top/bottom)
            header_margin = 150  # Skip top ~150pt (company name, title, date line)
            footer_margin = 50   # Skip bottom ~50pt (page numbers)
            body_top = header_margin
            body_bottom = page_height - footer_margin
            
            # Check for text content in the body area only
            words = page.extract_words()
            body_words = [w for w in words if w['top'] >= body_top and w['bottom'] <= body_bottom]
            if len(body_words) > 0:
                return True
                
            # Check for images in the body area
            images = page.images
            body_images = [img for img in images if img['top'] >= body_top and img['bottom'] <= body_bottom]
            if len(body_images) > 0:
                return True
                
            return False  # Page is empty (no text or images)
    except Exception as e:
        print(f"Warning: Could not analyze page {page_num + 1}: {str(e)}")
        return True  # Default to True to be safe

def create_blank_pdf_page():
    """
    Create a blank PDF page with standard A4 dimensions.
    Returns a PdfReader object containing a single blank page.
    """
    from io import BytesIO
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    
    # Create a blank page using reportlab
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    c.showPage()  # Create a blank page
    c.save()
    
    # Reset buffer position
    buffer.seek(0)
    
    # Return as PdfReader
    return PdfReader(buffer)

def remove_blank_pages_from_pdf(pdf_path):
    """
    Remove blank pages from a PDF file and replace it with the cleaned version.
    Returns True if any blank pages were removed, False otherwise.
    """
    try:
        print(f"Checking for blank pages in: {os.path.basename(pdf_path)}")
        
        # Read the original PDF
        reader = PdfReader(pdf_path)
        total_pages = len(reader.pages)
        
        if total_pages == 0:
            print(f"Warning: PDF has no pages: {pdf_path}")
            return False
        
        # Identify non-empty pages
        non_empty_pages = []
        blank_pages = []
        
        for page_num in range(total_pages):
            if is_content_page(pdf_path, page_num):
                non_empty_pages.append(page_num)
            else:
                blank_pages.append(page_num + 1)  # Human-readable page numbers
        
        # If no blank pages found, nothing to do
        if not blank_pages:
            print(f"No blank pages found in {os.path.basename(pdf_path)}")
            return False
        
        print(f"Found {len(blank_pages)} blank page(s) in {os.path.basename(pdf_path)}: pages {', '.join(map(str, blank_pages))}")
        
        # Create new PDF with only non-empty pages
        writer = PdfWriter()
        for page_num in non_empty_pages:
            writer.add_page(reader.pages[page_num])
        
        # Create temporary file path
        temp_path = pdf_path + '.tmp'
        
        # Write the cleaned PDF to temporary file
        with open(temp_path, 'wb') as temp_file:
            writer.write(temp_file)
        
        # Replace original file with cleaned version
        os.replace(temp_path, pdf_path)
        
        print(f"Removed {len(blank_pages)} blank page(s) from {os.path.basename(pdf_path)}")
        return True
        
    except Exception as e:
        print(f"Error removing blank pages from {pdf_path}: {e}")
        print(traceback.format_exc())
        
        # Clean up temporary file if it exists
        temp_path = pdf_path + '.tmp'
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass
        
        return False

def find_excel_column_index(worksheet, column_name):
    """Find the column index for a given column name in Excel worksheet."""
    for col_idx in range(1, worksheet.UsedRange.Columns.Count + 1):
        cell_value = worksheet.Cells(1, col_idx).Value
        if cell_value and str(cell_value).strip() == column_name:
            return col_idx
    return None

def verify_date_order(df, column_map):
    """
    Verify that dates are in correct order for each driver.
    For each driver, the start date of a row must be >= the end date of the previous row,
    and the end date must be >= the start date of the current row.
    
    Returns:
        tuple: (is_valid, list of error messages)
    """
    errors = []
    
    if df.empty:
        return True, []
    
    # Sort by driver and start date
    df_sorted = df.sort_values(by=[column_map['driver_full'], column_map['date_start']])
    
    # Group by driver
    for driver_name, driver_group in df_sorted.groupby(column_map['driver_full']):
        prev_end_date = None
        prev_row_num = None
        
        for idx, row in driver_group.iterrows():
            # Get row number (Excel row = pandas index + 2, assuming header at row 1)
            row_num = idx + 2

            # Skip АВТОСТРАДА rows - multiple overlapping trips per day are expected
            customer_col = column_map.get('customer')
            if customer_col and customer_col in row.index:
                if is_avtostrada_row(row.get(customer_col)):
                    continue

            start_date = pd.to_datetime(row[column_map['date_start']], errors='coerce')
            end_date = pd.to_datetime(row[column_map['date_end']], errors='coerce')
            
            # Skip rows with invalid dates
            if pd.isna(start_date) or pd.isna(end_date):
                errors.append(f"Row {row_num}: Driver '{driver_name}' has invalid dates")
                continue
            
            # Normalize to date only (remove time component) for comparison
            start_date_only = start_date.normalize()
            end_date_only = end_date.normalize()
            
            # Check that end date is >= start date
            if end_date_only < start_date_only:
                errors.append(
                    f"Row {row_num}: Driver '{driver_name}' has end date "
                    f"({end_date.strftime('%d.%m.%Y')}) before start date "
                    f"({start_date.strftime('%d.%m.%Y')})"
                )
            
            # Check that current start date is >= previous end date
            # Allow start date to be equal to previous end date (same day is okay)
            if prev_end_date is not None:
                if start_date_only < prev_end_date:
                    errors.append(
                        f"Row {row_num}: Driver '{driver_name}' has start date "
                        f"({start_date.strftime('%d.%m.%Y')}) before the end date "
                        f"({prev_end_date.strftime('%d.%m.%Y')}) of previous row {prev_row_num}"
                    )
            
            prev_end_date = end_date_only
            prev_row_num = row_num
    
    is_valid = len(errors) == 0
    return is_valid, errors

def write_to_excel_cell(excel_app, file_path, sheet_name, row_num, col_idx, value, hyperlink_path=None):
    """Write value to Excel cell with optional hyperlink using COM."""
    try:
        # Open workbook
        workbook = excel_app.Workbooks.Open(file_path)
        worksheet = workbook.Worksheets(sheet_name)
        
        # Write value to cell
        cell = worksheet.Cells(row_num, col_idx)
        cell.Value = value
        
        # Add hyperlink if provided
        if hyperlink_path:
            worksheet.Hyperlinks.Add(
                Anchor=cell,
                Address=str(hyperlink_path),
                TextToDisplay=str(value)
            )
        
        # Save and close
        workbook.Save()
        workbook.Close()
        return True
        
    except Exception as e:
        print(f"Error writing to Excel: {e}")
        return False

def group_assignments_by_date_and_company(df, column_map):
    """
    Group assignments by company code and start date for consolidated processing.
    Returns a dictionary where key is (company_code, start_date_str) and value is list of row indices.
    Only groups rows with the same company and start date.
    """
    grouped = {}
    
    for index, row in df.iterrows():
        company_code = str(row.get(column_map['company_code'], '')).strip().upper()
        date_obj_start = pd.to_datetime(row.get(column_map['date_start']), errors='coerce')
        
        if not company_code or pd.isna(date_obj_start):
            continue
            
        date_str = date_obj_start.strftime('%d.%m.%Y')
        key = (company_code, date_str)
        
        if key not in grouped:
            grouped[key] = []
        grouped[key].append(index)
    
    return grouped

def insert_multiple_driver_lines(word_doc, driver_data_list, use_short_names=False):
    """
    Insert multiple driver assignment lines and signature lines in the Word document.
    
    Args:
        word_doc: Word document COM object
        driver_data_list: List of dictionaries, each containing:
            - driver_full: Full driver name
            - driver_initials: Driver initials
            - truck_model: Truck model
            - plate_number: Plate number
            - route_desc: Route description
            - date_start: Start date
            - date_end: End date
        use_short_names: If True, use initials; if False, use full names (default: False)
    
    Returns:
        True if successful, False otherwise
    """
    try:
        # Sort drivers alphabetically by full name for consistent ordering
        sorted_drivers = sorted(driver_data_list, key=lambda x: x['driver_full'])
        
        # === PART 1: Replace the assignment line with multiple numbered lines ===
        # Find the line that contains the driver assignment template
        # Search for unique text that appears in the template line
        search_text = "направити у відрядження за маршрутом"
        
        find_obj = word_doc.Content.Find
        find_obj.ClearFormatting()
        find_obj.Text = search_text
        find_obj.Forward = True
        find_obj.Wrap = 1  # wdFindContinue
        
        if not find_obj.Execute():
            print("Warning: Could not find assignment line template in document")
            return False
        
        # Get the paragraph containing this text
        found_range = find_obj.Parent
        para_range = found_range.Paragraphs(1).Range
        
        # Build the replacement text with all drivers (WITHOUT manual numbering - Word handles it)
        replacement_lines = []
        for driver_data in sorted_drivers:
            truck_str = f"{driver_data.get('truck_model', '')} ".lstrip()
            line = f"{driver_data['driver_full']}, водій автотранспортного засобу {truck_str}{driver_data['plate_number']}, направити у відрядження за маршрутом {driver_data['route_desc']}, терміном з {driver_data['date_start']} по {driver_data['date_end']}"
            replacement_lines.append(line)
        
        # Join all lines with paragraph breaks
        replacement_text = "\r\n".join(replacement_lines)
        
        # Replace the entire paragraph with the new text
        para_range.Text = replacement_text + "\r\n"
        
        # === PART 2: Replace signature lines with unique drivers only ===
        # Find the signature line template
        signature_search = "Водій ______________"
        
        find_obj2 = word_doc.Content.Find
        find_obj2.ClearFormatting()
        find_obj2.Text = signature_search
        find_obj2.Forward = True
        find_obj2.Wrap = 1
        
        if not find_obj2.Execute():
            print("Warning: Could not find signature line template in document")
            return False
        
        # Get the paragraph containing the signature
        sig_found_range = find_obj2.Parent
        sig_para_range = sig_found_range.Paragraphs(1).Range
        
        # Build unique driver signature lines (sorted alphabetically, no duplicates)
        # Add 3 blank lines between each signature line
        seen_drivers = set()
        signature_lines = []
        for driver_data in sorted_drivers:
            driver_name = driver_data['driver_full'] if not use_short_names else driver_data['driver_initials']
            driver_key = driver_data['driver_full']  # Use full name as unique key regardless
            
            if driver_key not in seen_drivers:
                seen_drivers.add(driver_key)
                signature_lines.append(f"Водій ______________ {driver_name}")
        
        # Join signature lines with 3 line breaks (empty lines) between them
        signature_text = "\r\n\r\n\r\n".join(signature_lines)
        
        # Replace the signature paragraph
        sig_para_range.Text = signature_text + "\r\n"
        
        print(f"Successfully inserted {len(sorted_drivers)} driver assignment lines and {len(signature_lines)} unique signature lines")
        return True
        
    except Exception as e:
        print(f"Error inserting multiple driver lines: {e}")
        print(traceback.format_exc())
        return False

def create_consolidated_filename(company_code, driver_initials_list, date_str, sequence_num):
    """
    Create filename for consolidated assignment.
    Format: "н#_consolidated COMPANY від DD.MM.YYYY Driver1, Driver2, Driver3"
    
    Args:
        company_code: Company code string
        driver_initials_list: List of driver initials (may contain duplicates)
        date_str: Date string in DD.MM.YYYY format
        sequence_num: Sequence number
    
    Returns:
        Base filename without extension
    """
    # Remove duplicates while preserving alphabetical order
    unique_drivers = []
    seen = set()
    for driver in sorted(driver_initials_list):
        if driver not in seen:
            unique_drivers.append(driver)
            seen.add(driver)
    
    # Join driver initials with commas
    drivers_str = ", ".join(unique_drivers)
    
    # Truncate if too long (Windows path limit is 260 chars, leave room for path and extension)
    max_drivers_length = 150
    if len(drivers_str) > max_drivers_length:
        # Truncate and add ellipsis
        drivers_str = drivers_str[:max_drivers_length] + "..."
    
    return f"н{sequence_num}_consolidated {company_code} від {date_str} {drivers_str}"

def truncate_path_to_limit(path, max_len=255):
    """Truncate a file path's stem so the full path string fits within max_len characters."""
    if len(str(path)) <= max_len:
        return path
    suffix = path.suffix
    parent_str = str(path.parent)
    available = max_len - len(parent_str) - 1 - len(suffix)  # 1 for the path separator
    if available <= 0:
        return path
    return path.parent / (path.stem[:available] + suffix)

class ExcelBatchWriter:
    """Helper class to batch Excel write operations for better performance."""
    
    def __init__(self, excel_app, file_path, sheet_name, assignments_col_idx):
        self.excel_app = excel_app
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.assignments_col_idx = assignments_col_idx
        self.workbook = None
        self.worksheet = None
        self.pending_writes = []
        
    def __enter__(self):
        """Open workbook for batch operations."""
        try:
            self.workbook = self.excel_app.Workbooks.Open(self.file_path)
            self.worksheet = self.workbook.Worksheets(self.sheet_name)
            return self
        except Exception as e:
            print(f"Error opening Excel workbook for batch writing: {e}")
            return None
            
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Save and close workbook."""
        if self.workbook:
            try:
                self.workbook.Save()
                self.workbook.Close()
            except Exception as e:
                print(f"Error closing Excel workbook: {e}")
                
    def add_sequence_number(self, row_num, sequence_num, hyperlink_path):
        """Add a sequence number with hyperlink to the batch."""
        try:
            if self.worksheet and self.assignments_col_idx:
                cell = self.worksheet.Cells(row_num, self.assignments_col_idx)
                cell.Value = sequence_num
                
                # Add hyperlink
                self.worksheet.Hyperlinks.Add(
                    Anchor=cell,
                    Address=str(hyperlink_path),
                    TextToDisplay=str(sequence_num)
                )
                return True
        except Exception as e:
            print(f"Error adding sequence number to Excel row {row_num}: {e}")
            return False
        return False

# --- Main Script ---
if __name__ == "__main__":
    setup_logging()
    # --- Argument Parsing ---
    parser = argparse.ArgumentParser(description="Generate assignments for specific company.")
    parser.add_argument("--zet", action="store_true", help="Process only ЗЕТТРА")
    parser.add_argument("--zia", action="store_true", help="Process only ЗІАВТОТРАНС")
    parser.add_argument("--month", type=str, help="Month to process in format MMYY (e.g., 0525 for May 2025). Default: previous month")
    parser.add_argument("--service-only", action="store_true", help="Process only service agreements (цпх) without regular assignments")
    parser.add_argument("--assignment-only", action="store_true", help="Process only regular assignments without service agreements (цпх)")
    parser.add_argument("--no-blank-removal", action="store_true", help="Disable automatic blank page removal from PDFs")
    parser.add_argument("--verify-only", nargs='?', const=True, default=False, help="Verify consolidated PDF. Without value: verify date order in Excel. With value (e.g., zia0326): verify page order in existing consolidated PDF")
    parser.add_argument("--no-excel-update", action="store_true", help="Do not update Excel 'накази' column with sequence numbers")
    parser.add_argument("--no-consolidate", action="store_true", help="Disable consolidation of assignments")
    parser.add_argument("--consolidate", nargs='?', const=True, default=True, help="Consolidate assignments by start date. Without value: combine during processing. With value (e.g., zet0326): rebuild combined PDF from existing files in the company folder")
    parser.add_argument("--short-names", action="store_true", help="Use short driver initials (e.g., 'Іваненко І. П.') instead of full names in signatures")
    parser.add_argument("--both-sides", action="store_true", help="Add blank pages between assignments in combined PDF so each starts on a sheet front (for duplex printing)")
    args = parser.parse_args()

    if args.no_consolidate:
        args.consolidate = False

    # Validate mutually exclusive arguments
    if args.service_only and args.assignment_only:
        print("Error: Cannot use both --service-only and --assignment-only at the same time")
        sys.exit(1)

    company_filter = None
    if args.zet:
        company_filter = "ЗЕТТРА"
    elif args.zia:
        company_filter = "ЗІАВТОТРАНС"    
    
    # --- Standalone verify mode: check existing consolidated PDF ---
    if isinstance(args.verify_only, str):
        verify_val = args.verify_only.lower().strip()
        
        match = re.match(r'^([a-z]+)(\d{4})$', verify_val)
        if not match:
            print(f"Error: Invalid --verify-only value '{args.verify_only}'. Use format like 'zia0326' (company + MMYY)")
            sys.exit(1)
        
        company_short = match.group(1)
        month_year_str = match.group(2)
        
        if company_short not in COMPANY_SHORT_CODES:
            print(f"Error: Unknown company code '{company_short}'. Valid codes: {', '.join(COMPANY_SHORT_CODES.keys())}")
            sys.exit(1)
        
        month_data = parse_month_arg(month_year_str)
        if not month_data:
            print(f"Error: Invalid month format '{month_year_str}'. Must be MMYY (e.g., 0326)")
            sys.exit(1)
        
        target_year, target_month = month_data
        company_name = COMPANY_SHORT_CODES[company_short]
        pdf_dir = OUTPUT_ROOT_PATH / f"{company_name}-pdf" / month_year_str
        
        if not pdf_dir.exists():
            print(f"Error: Folder not found: {pdf_dir}")
            sys.exit(1)
        
        # Find consolidated PDF
        consolidated_files = list(pdf_dir.glob(f"CONSOLIDATED_ALL*{company_name}*.pdf"))
        if not consolidated_files:
            print(f"Error: No consolidated PDF found in {pdf_dir}")
            sys.exit(1)
        
        consolidated_path = consolidated_files[0]
        print(f"Verifying: {consolidated_path.name}")
        print(f"Folder: {pdf_dir}")
        print()
        
        # Also list source assignment PDFs for comparison
        source_pdfs = list(pdf_dir.glob(f"\u043d*_consolidated {company_name} \u0432\u0456\u0434 *.pdf"))
        
        def _extract_seq(fp):
            m = re.match(r'\D*(\d+)', os.path.basename(str(fp)))
            return int(m.group(1)) if m else 0
        
        source_pdfs.sort(key=_extract_seq)
        print(f"Source assignment PDFs: {len(source_pdfs)}")
        
        # Analyze consolidated PDF page by page
        import pdfplumber
        errors = []
        prev_seq = 0
        prev_page = 0
        
        with pdfplumber.open(str(consolidated_path)) as pdf:
            total_pages = len(pdf.pages)
            print(f"Consolidated PDF pages: {total_pages}")
            print()
            
            for i, page in enumerate(pdf.pages):
                text = page.extract_text() or ''
                # Extract assignment number
                seq_match = re.search(r'[\u2116]\s*(\d+)', text)
                # Extract date
                date_match = re.search(r'(\d{1,2})\s+\w+\s+\d{4}\s+\u0440', text)
                
                if seq_match:
                    seq_num = int(seq_match.group(1))
                    date_str = date_match.group(0)[:30] if date_match else '?'
                    
                    # Check for blank page (header/footer only)
                    is_blank = not is_content_page(str(consolidated_path), i)
                    blank_marker = ' \u26a0 BLANK (header/footer only)' if is_blank else ''
                    
                    # Check ordering
                    order_marker = ''
                    if seq_num < prev_seq and seq_num != prev_seq:
                        order_marker = f' \u2757 OUT OF ORDER (prev was #{prev_seq})'
                        errors.append(f"Page {i+1}: Assignment #{seq_num} after #{prev_seq}")
                    
                    print(f"Page {i+1:3d}: Assignment #{seq_num:3d}  {date_str}{blank_marker}{order_marker}")
                    prev_seq = seq_num
                else:
                    # Non-assignment page (index, etc.)
                    first_line = text.split('\n')[0][:60] if text else '(empty)'
                    print(f"Page {i+1:3d}: [{first_line}]")
        
        # Check for missing assignments
        source_seqs = set(_extract_seq(p) for p in source_pdfs)
        with pdfplumber.open(str(consolidated_path)) as pdf:
            consolidated_seqs = set()
            for page in pdf.pages:
                text = page.extract_text() or ''
                m = re.search(r'[\u2116]\s*(\d+)', text)
                if m:
                    consolidated_seqs.add(int(m.group(1)))
        
        missing = source_seqs - consolidated_seqs
        extra = consolidated_seqs - source_seqs
        
        print()
        if missing:
            print(f"\u2757 MISSING assignments in consolidated PDF: {sorted(missing)}")
            errors.append(f"Missing assignments: {sorted(missing)}")
        if extra:
            print(f"\u26a0 Extra assignments in consolidated PDF (not in folder): {sorted(extra)}")
        
        if errors:
            print(f"\n\u274c VERIFICATION FAILED - {len(errors)} issue(s) found")
        else:
            print(f"\u2705 VERIFICATION PASSED - all {len(source_seqs)} assignments in correct order")
        
        sys.exit(1 if errors else 0)

    # --- Standalone consolidation mode: rebuild combined PDF from existing files ---
    if isinstance(args.consolidate, str):
        # Parse the value, e.g. "zet0326" -> company="zet", month_year="0326"
        consolidate_val = args.consolidate.lower().strip()
        
        # Extract company short code (letters) and month-year (digits)
        match = re.match(r'^([a-z]+)(\d{4})$', consolidate_val)
        if not match:
            print(f"Error: Invalid --consolidate value '{args.consolidate}'. Use format like 'zet0326' (company + MMYY)")
            sys.exit(1)
        
        company_short = match.group(1)
        month_year_str = match.group(2)
        
        if company_short not in COMPANY_SHORT_CODES:
            print(f"Error: Unknown company code '{company_short}'. Valid codes: {', '.join(COMPANY_SHORT_CODES.keys())}")
            sys.exit(1)
        
        month_data = parse_month_arg(month_year_str)
        if not month_data:
            print(f"Error: Invalid month format '{month_year_str}'. Must be MMYY (e.g., 0326)")
            sys.exit(1)
        
        target_year, target_month = month_data
        company_filter = COMPANY_SHORT_CODES[company_short]
        pdf_dir = OUTPUT_ROOT_PATH / f"{company_filter}-pdf" / month_year_str
        
        if not pdf_dir.exists():
            print(f"Error: Folder not found: {pdf_dir}")
            sys.exit(1)
        
        print(f"Standalone consolidation mode")
        print(f"Company: {company_filter}")
        print(f"Period: {target_month:02d}/{target_year}")
        print(f"Source folder: {pdf_dir}")
        
        # Collect consolidated assignment PDFs (exclude index and consolidated files)
        all_consolidated_pdfs = list(pdf_dir.glob(f"\u043d*_consolidated {company_filter} \u0432\u0456\u0434 *.pdf"))
        
        # Sort by sequence number
        def extract_sequence_number(filepath):
            filename = os.path.basename(str(filepath))
            m = re.match(r'\D*(\d+)', filename)
            return int(m.group(1)) if m else 0
        
        all_consolidated_pdfs.sort(key=extract_sequence_number)
        
        print(f"Found {len(all_consolidated_pdfs)} assignment PDFs")
        
        # Debug: list all files with their sequence numbers and check for duplicates
        seq_seen = {}
        for pdf_path in all_consolidated_pdfs:
            seq = extract_sequence_number(pdf_path)
            fname = os.path.basename(str(pdf_path))
            print(f"  [{seq:3d}] {fname[:80]}")
            if seq in seq_seen:
                print(f"  ⚠ DUPLICATE seq {seq}! Previous: {seq_seen[seq][:60]}")
            seq_seen[seq] = fname
        
        unique_paths = set(str(p) for p in all_consolidated_pdfs)
        if len(unique_paths) != len(all_consolidated_pdfs):
            print(f"  ⚠ DUPLICATE PATHS DETECTED: {len(all_consolidated_pdfs)} total, {len(unique_paths)} unique")
        
        if not all_consolidated_pdfs:
            print("No assignment PDFs found to merge.")
            sys.exit(1)
        
        # Find signing index
        index_files = list(pdf_dir.glob("SIGNING_INDEX_*.pdf"))
        signing_index_pdf = index_files[0] if index_files else None
        if signing_index_pdf:
            print(f"Signing index: {signing_index_pdf.name}")
        else:
            print("No signing index found, skipping")
        
        # Build merged PDF
        merged_pdf = PdfWriter()
        # IMPORTANT: Keep all PdfReader objects alive until write() is called.
        # PyPDF2 uses lazy references; if a reader is garbage-collected, page content gets corrupted.
        active_readers = []
        use_both_sides = args.both_sides
        
        if signing_index_pdf and signing_index_pdf.exists():
            try:
                index_reader = PdfReader(str(signing_index_pdf))
                active_readers.append(index_reader)
                index_page_count = len(index_reader.pages)
                for page in index_reader.pages:
                    merged_pdf.add_page(page)
                print(f"Added signing index ({index_page_count} page(s))")
                
                if use_both_sides and index_page_count % 2 == 1:
                    blank_page_reader = create_blank_pdf_page()
                    active_readers.append(blank_page_reader)
                    merged_pdf.add_page(blank_page_reader.pages[0])
                    print(f"Added 1 blank separator page after index (both-sides mode)")
            except Exception as e:
                print(f"Warning: Could not add signing index: {e}")
        
        for pdf_path in all_consolidated_pdfs:
            print(f"Adding: {os.path.basename(str(pdf_path))}")
            try:
                reader = PdfReader(str(pdf_path))
                active_readers.append(reader)
                page_count = len(reader.pages)
                added_count = 0
                skipped_blank = 0
                for page_idx in range(page_count):
                    if is_content_page(str(pdf_path), page_idx):
                        merged_pdf.add_page(reader.pages[page_idx])
                        added_count += 1
                    else:
                        skipped_blank += 1
                        print(f"  ⚠ Skipped blank page {page_idx + 1}/{page_count} (header/footer only)")
                
                if use_both_sides and added_count % 2 == 1:
                    blank_page_reader = create_blank_pdf_page()
                    active_readers.append(blank_page_reader)
                    merged_pdf.add_page(blank_page_reader.pages[0])
                    print(f"  → Added {added_count} pages + 1 blank page")
                else:
                    print(f"  → Added {added_count} pages{f' (removed {skipped_blank} blank)' if skipped_blank else ''}")
            except Exception as e:
                print(f"Warning: Could not add {pdf_path}: {e}")
        
        # Write combined PDF (overwrite if exists)
        sides_suffix = "_BOTH_SIDES" if args.both_sides else ""
        combined_filename = f"CONSOLIDATED_ALL{sides_suffix}_{company_filter}_{target_month:02d}.{target_year}.pdf"
        combined_pdf_path = pdf_dir / combined_filename
        
        with open(combined_pdf_path, 'wb') as output_file:
            merged_pdf.write(output_file)
        
        try:
            total_pages = len(PdfReader(str(combined_pdf_path)).pages)
        except Exception:
            total_pages = None
        
        print(f"\nCombined PDF created: {combined_pdf_path}")
        if total_pages:
            print(f"Total pages: {total_pages}")
        
        # Prompt to print
        print(f"\nPress Enter to print or Ctrl+C to skip...")
        try:
            input()
            os.startfile(str(combined_pdf_path), "print")
            print("Sent to printer")
        except KeyboardInterrupt:
            print("\nPrinting skipped.")
        except Exception as e_print:
            print(f"Could not print automatically: {e_print}")
            print(f"Please print manually from: {combined_pdf_path}")
        
        sys.exit(0)
    
    # --- Parse month argument and determine Excel file path ---
    if not args.month:
        try:
            prev_year, prev_month = get_previous_month()
            default_month_str = f"{prev_month:02d}{str(prev_year)[2:]}"
            user_input = input(f"Введіть місяць для обробки у форматі ММРР [за замовчуванням: {default_month_str}]: ").strip()
            if user_input:
                args.month = user_input
            
            # Ask for company filter
            if not args.zet and not args.zia:
                print("\nОберіть компанію для фільтрації:")
                print("1. ЗЕТТРА")
                print("2. ЗІАВТОТРАНС")
                print("Enter. Всі компанії (без фільтрації)")
                comp_input = input("Ваш вибір: ").strip()
                if comp_input == '1' or comp_input.lower() == 'zet':
                    args.zet = True
                elif comp_input == '2' or comp_input.lower() == 'zia':
                    args.zia = True

            # Ask for processing mode
            if not args.service_only and not args.assignment_only:
                print("\nОберіть режим обробки:")
                print("1. Накази на відрядження та ЦПХ (за замовчуванням)")
                print("2. Тільки накази на відрядження")
                print("3. Тільки договори ЦПХ")
                mode_input = input("Ваш вибір: ").strip()
                if mode_input == '2':
                    args.assignment_only = True
                elif mode_input == '3':
                    args.service_only = True
        except (KeyboardInterrupt, EOFError):
            print("\nСкасовано.")
            sys.exit(0)

    # Re-evaluate company filter based on CLI or interactive inputs
    company_filter = None
    if args.zet:
        company_filter = "ЗЕТТРА"
    elif args.zia:
        company_filter = "ЗІАВТОТРАНС"

    if args.month:
        month_data = parse_month_arg(args.month)
        if not month_data:
            print(f"Error: Invalid month format '{args.month}'. Use MMYY format (e.g., 0525 for May 2025)")
            sys.exit(1)
        target_year, target_month = month_data
        excel_file_path = get_excel_file_path(args.month)
    else:
        target_year, target_month = get_previous_month()
        excel_file_path = get_excel_file_path()
    
    EXCEL_FILE_PATH = excel_file_path
    SHEET_NAME = get_sheet_name(excel_file_path)
    
    # Build short MMYY folder name used in all output directory paths
    month_year_str = f"{target_month:02d}{str(target_year)[2:]}"

    # Verify the Excel file exists
    if not os.path.exists(excel_file_path):
        print(f"Error: Excel file not found at '{excel_file_path}'")
        sys.exit(1)

    # Override blank page removal setting if specified
    remove_blank_pages = REMOVE_BLANK_PAGES and not args.no_blank_removal
    if args.no_blank_removal:
        print("Blank page removal disabled by command line argument")

    print(f"Starting script...")
    print(f"Target month: {target_month:02d}/{target_year}")
    print(f"Reading Excel file: {excel_file_path}")
    print(f"Sheet: {SHEET_NAME}, Table: {TABLE_NAME}")
    if remove_blank_pages:
        print("Blank page removal: ENABLED")
    else:
        print("Blank page removal: DISABLED")
    if args.consolidate:
        print("Consolidation mode: ENABLED - Combining assignments by start date")
    else:
        print("Consolidation mode: DISABLED - Creating individual assignments")

    word_app = None # Initialize Word Application variable
    excel_app = None # Initialize Excel Application variable
    excel_assignments_col_idx = None # Column index for "накази" column
    company_file_counts = {}
    company_sequence_counters = {}  # Track next sequence number per company
    report_lines = []
    
    # Track paper savings for consolidated mode
    total_assignments_individual = 0  # How many individual assignments would be created
    total_assignments_consolidated = 0  # How many consolidated documents actually created
    consolidated_groups_count = 0  # Number of consolidated groups

    try:
        # --- Read Excel Data ---
        # Use openpyxl engine for .xlsm. Check if table name needs header row adjustment.
        # Pandas usually detects the header automatically if table is formatted correctly.
        df = pd.read_excel(excel_file_path, sheet_name=SHEET_NAME, engine='openpyxl')
        
        print(f"Total rows in Excel: {len(df)}")

        # --- Validate required columns exist ---
        optional_keys = {'comment', 'customer', 'freight'}
        required_cols = [col for key, col in COLUMN_MAP.items() if key not in optional_keys]
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Missing required columns in sheet '{SHEET_NAME}': {', '.join(missing_cols)}")
        
        # --- Normalize Company Names ---
        company_col = COLUMN_MAP['company_code']
        if company_col in df.columns:
            def normalize_company(val):
                if pd.isna(val):
                    return val
                val_str = str(val).strip().upper()
                if val_str in ('ЗІА', 'ЗІАВТОТРАНС'):
                    return 'ЗІАВТОТРАНС'
                if val_str in ('ЗЕТ', 'ЗЕТТРА'):
                    return 'ЗЕТТРА'
                return str(val).strip()
            df[company_col] = df[company_col].apply(normalize_company)
        
        # Check if parameter column exists for service agreement detection
        has_comment_col = COLUMN_MAP['comment'] in df.columns
        if not has_comment_col:
            print(f"Warning: Parameter column '{COLUMN_MAP['comment']}' not found. Service agreements won't be detected.")
            if args.service_only:
                print("Error: Cannot process service agreements only without parameter column.")
                sys.exit(1)
            if args.assignment_only:
                print("Warning: Cannot filter out service agreements without parameter column. All rows will be processed as regular assignments.")

        # Check if customer column exists for АВТОСТРАДА detection
        has_customer_col = COLUMN_MAP.get('customer') and COLUMN_MAP['customer'] in df.columns
        if not has_customer_col:
            print(f"Warning: Customer column '{COLUMN_MAP.get('customer', 'замовник')}' not found. АВТОСТРАДА per-day service docs disabled.")

        # --- Convert date column to datetime ---
        df[COLUMN_MAP['date_start']] = pd.to_datetime(df[COLUMN_MAP['date_start']], errors='coerce')
        
        # --- Filter by target month ---
        date_filter = (df[COLUMN_MAP['date_start']].dt.year == target_year) & \
                     (df[COLUMN_MAP['date_start']].dt.month == target_month)
        df = df[date_filter]
        print(f"Rows in target month ({target_month:02d}/{target_year}): {len(df)}")
        
        if len(df) == 0:
            print("No rows to process for the target month. Exiting.")
            sys.exit(0)
        
        # --- Filter: explicit exclude via parameter ---
        if has_comment_col:
            before_explicit = len(df)
            df = df[~df[COLUMN_MAP['comment']].apply(is_exclude_param)]
            explicit_excluded = before_explicit - len(df)
            if explicit_excluded:
                print(f"Excluded rows via ${{exclude}} marker: {explicit_excluded}")
        
        # --- Filter out cash payments (готівка) ---
        if COLUMN_MAP.get('payment_method') and COLUMN_MAP['payment_method'] in df.columns:
            before_payment_filter = len(df)
            payment_filter = df[COLUMN_MAP['payment_method']].str.upper() != 'ГОТІВКА'
            df = df[payment_filter]
            print(f"Rows after excluding cash payments: {len(df)} (excluded {before_payment_filter - len(df)} cash payments)")

        # --- Filter out drivers with status "найм" ---
        # Exceptions: rows with ${assignment} or ${sa;...} in the parameter column are kept regardless
        if COLUMN_MAP.get('driver_status') and COLUMN_MAP['driver_status'] in df.columns:
            before_status_filter = len(df)
            is_najm = df[COLUMN_MAP['driver_status']].astype(str).str.upper().str.contains('НАЙМ', na=False, regex=True)
            if has_comment_col:
                has_assignment_param = df[COLUMN_MAP['comment']].astype(str).str.contains(r'\$\{assignment\}|наказ', na=False, case=False, regex=True)
                has_sa_param = df[COLUMN_MAP['comment']].apply(is_service_agreement)
                najm_exception = has_assignment_param | has_sa_param
            else:
                najm_exception = pd.Series(False, index=df.index)
            status_filter = ~is_najm | najm_exception
            df = df[status_filter]
            excluded = before_status_filter - len(df)
            exceptions = is_najm.sum() - excluded
            print(f"Rows after excluding drivers with status 'найм': {len(df)} (excluded {excluded}, kept {exceptions} найм exceptions via assignment or service agreement)")

        # --- Filter by company if argument is provided ---
        if company_filter:
            before_company_filter = len(df)
            df = df[df[COLUMN_MAP['company_code']].str.upper() == company_filter]
            print(f"Filtered to company '{company_filter}': {len(df)} rows (excluded {before_company_filter - len(df)} from other companies)")

        # --- Filter by service agreement type if specified ---
        if args.service_only and has_comment_col:
            before_service_filter = len(df)
            df = df[df[COLUMN_MAP['comment']].apply(is_service_agreement)]
            print(f"Filtered to service agreements only: {len(df)} rows (excluded {before_service_filter - len(df)} regular assignments)")
        elif args.assignment_only and has_comment_col:
            before_assignment_filter = len(df)
            df = df[~df[COLUMN_MAP['comment']].apply(is_service_agreement)]
            print(f"Filtered to regular assignments only: {len(df)} rows (excluded {before_assignment_filter - len(df)} service agreements)")
        elif not args.service_only and not args.assignment_only and has_comment_col:
            # If neither filter is specified, process both but we'll handle them differently in the loop
            pass

        # --- Verify date order for all drivers ---
        print(f"\n=== VERIFYING DATE ORDER ===")
        is_valid, verification_errors = verify_date_order(df, COLUMN_MAP)
        
        if not is_valid:
            print(f"\n!!! DATE ORDER VERIFICATION FAILED !!!")
            print(f"Found {len(verification_errors)} error(s):\n")
            for error in verification_errors:
                print(f"  ERROR: {error}")
            
            if args.verify_only:
                print("\nVerification only mode - no files will be created.")
                print("Please fix the errors above and run again.")
                sys.exit(1)
            else:
                print("\nWARNING: Date order errors detected!")
                print("You can continue, but the data may have issues.")
                print("Press Enter to continue anyway or Ctrl+C to abort and fix errors...")
                try:
                    input()
                except KeyboardInterrupt:
                    print("\nOperation aborted by user.")
                    sys.exit(0)
        else:
            print(f"Date order verification passed - all dates are sequential for each driver")
            
            if args.verify_only:
                print("\nVerification only mode - verification successful!")
                print("No errors found. You can now run without --verify-only to create files.")
                sys.exit(0)
        
        # --- Calculate consolidation statistics early if needed ---
        consolidation_preview_stats = None
        if args.consolidate and not args.service_only:
            # Filter out service agreements and АВТОСТРАДА rows from consolidation preview
            df_regular_preview = df.copy()
            if has_comment_col:
                df_regular_preview = df_regular_preview[~df_regular_preview[COLUMN_MAP['comment']].apply(is_service_agreement)]
            if has_customer_col:
                df_regular_preview = df_regular_preview[~df_regular_preview[COLUMN_MAP['customer']].apply(is_avtostrada_row)]
            
            if len(df_regular_preview) > 0:
                # Group assignments by company and date to get preview stats
                from collections import defaultdict
                preview_groups = defaultdict(list)
                
                for index, row in df_regular_preview.iterrows():
                    company_code = row.get(COLUMN_MAP['company_code'])
                    date_start = row.get(COLUMN_MAP['date_start'])
                    
                    if pd.notna(date_start) and pd.notna(company_code):
                        if isinstance(date_start, str):
                            date_str = date_start
                        else:
                            date_str = date_start.strftime('%d.%m.%Y')
                        
                        key = (company_code, date_str)
                        preview_groups[key].append(index)
                
                # Calculate preview statistics
                total_assignments_individual = len(df_regular_preview)
                total_assignments_consolidated = len(preview_groups)
                consolidated_groups_count = len([g for g in preview_groups.values() if len(g) > 1])
                papers_saved = total_assignments_individual - total_assignments_consolidated
                
                consolidation_preview_stats = {
                    'unique_combinations': len(preview_groups),
                    'individual': total_assignments_individual,
                    'consolidated': total_assignments_consolidated,
                    'multi_driver_groups': consolidated_groups_count,
                    'papers_saved': papers_saved
                }
        
        # --- Show summary and ask for confirmation ---
        print(f"\n=== PROCESSING SUMMARY ===")
        print(f"Target period: {target_month:02d}/{target_year}")
        print(f"Company filter: {company_filter if company_filter else 'All companies'}")
        
        # Show processing mode
        if args.service_only:
            print(f"Processing mode: Service agreements only (цпх)")
        elif args.assignment_only:
            print(f"Processing mode: Regular assignments only")
        else:
            print(f"Processing mode: All types (assignments and service agreements)")
        
        # Show consolidation mode with statistics
        if args.consolidate:
            print(f"Consolidation: ENABLED - Multiple drivers with same start date will be combined")
            if consolidation_preview_stats:
                print(f"  → Found {consolidation_preview_stats['unique_combinations']} unique date/company combinations")
                print(f"  → Individual assignments: {consolidation_preview_stats['individual']}")
                print(f"  → Consolidated documents: {consolidation_preview_stats['consolidated']}")
                print(f"  → Groups with 2+ drivers: {consolidation_preview_stats['multi_driver_groups']}")
                print(f"  → PAPER SAVINGS: {consolidation_preview_stats['papers_saved']} fewer documents")
        else:
            print(f"Consolidation: DISABLED - Each driver gets individual assignment")
            
        print(f"Total rows to process: {len(df)}")
        
        if len(df) == 0:
            print("No rows to process. Exiting.")
            sys.exit(0)
              # Group by company for summary
        if len(df) > 0:
            company_summary = df[COLUMN_MAP['company_code']].value_counts()
            print("\nRows by company:")
            for company, count in company_summary.items():
                print(f"  {company}: {count} rows")
        
        print("\nPress Enter to continue or Ctrl+C to abort...")
        try:
            input()
        except KeyboardInterrupt:
            print("\nOperation aborted by user.")
            sys.exit(0)

        # --- Initialize sequence counters per company ---
        print("\nInitializing sequence counters...")
        unique_companies = df[COLUMN_MAP['company_code']].unique()
        company_sequence_counters = {}
        
        # Collect unique service agreement groups by timestamp (AFTER filtering)
        service_agreement_groups = {}
        
        if has_comment_col:
            for index, row in df.iterrows():
                # АВТОСТРАДА rows are handled separately - skip them here
                if has_customer_col and is_avtostrada_row(row.get(COLUMN_MAP['customer'])):
                    continue
                if is_service_agreement(row.get(COLUMN_MAP['comment'])):
                    sa_details = extract_service_agreement_details(row.get(COLUMN_MAP['comment']))
                    if sa_details:
                        timestamp = sa_details['timestamp']
                        if not timestamp:
                            driver_name = str(row.get(COLUMN_MAP['driver_full'], '')).strip()
                            driver_clean = re.sub(r'[^\w\s]', '', driver_name).replace(' ', '_').lower()
                            timestamp = f"{driver_clean}_{sa_details['company']}".lower()
                        # Use company code from parameter, not from Excel column
                        param_company = sa_details['company']
                        # Map parameter company codes to full company names
                        company_mapping = {
                            'ZIA': 'ЗІАВТОТРАНС',
                            'ZET': 'ЗЕТТРА'
                        }
                        company_code = company_mapping.get(param_company, param_company)
                        
                        if timestamp not in service_agreement_groups:
                            service_agreement_groups[timestamp] = {
                                'company_code': company_code,
                                'param_company': param_company,  # Keep original parameter company code
                                'money': sa_details['money'],
                                'rows': [],
                                'row_data': [],  # Store complete row data for each row
                                'routes': [],
                                'freights': [],
                                'latest_end_date': None,
                                'first_row_data': None
                            }
                        
                        # Add row index and complete row data
                        service_agreement_groups[timestamp]['rows'].append(index)
                        
                        # Store complete row data for certificate replacements
                        row_data = {
                            'date_start': pd.to_datetime(row.get(COLUMN_MAP['date_start']), errors='coerce').strftime('%d.%m.%Y') if not pd.isna(pd.to_datetime(row.get(COLUMN_MAP['date_start']), errors='coerce')) else '',
                            'date_end': pd.to_datetime(row.get(COLUMN_MAP['date_end']), errors='coerce').strftime('%d.%m.%Y') if not pd.isna(pd.to_datetime(row.get(COLUMN_MAP['date_end']), errors='coerce')) else '',
                            'route_desc': str(row.get(COLUMN_MAP['route_desc'], '')).strip(),
                            'freight': str(row.get(COLUMN_MAP['freight'], '')).strip()
                        }
                        service_agreement_groups[timestamp]['row_data'].append(row_data)
        
        # Initialize company-specific service agreement sequence counters
        company_service_sequence_counters = {}
        service_companies = set()
        for timestamp, group in service_agreement_groups.items():
            service_companies.add(group['company_code'])
        
        for company in service_companies:
            company_service_sequence_counters[company] = get_company_service_sequence_number(company)
            print(f"Service agreements for {company} starting from sequence {company_service_sequence_counters[company]}")
        
        # First, collect all route and freight information for each group
        for timestamp, group in service_agreement_groups.items():
            # Continue collecting route and freight information
            for row_index in group['rows']:
                # Use .loc instead of .iloc since we're working with original indices
                row = df.loc[row_index]
                
                # Collect route information
                route = row.get(COLUMN_MAP['route_desc'], '')
                if route and str(route).strip():
                    group['routes'].append(str(route).strip())
                
                # Collect freight information
                freight = row.get(COLUMN_MAP['freight'], '')
                if freight and str(freight).strip():
                    group['freights'].append(str(freight).strip())
                
                # Track latest end date
                end_date = pd.to_datetime(row.get(COLUMN_MAP['date_end']), errors='coerce')
                if not pd.isna(end_date):
                    if (group['latest_end_date'] is None or 
                        end_date > group['latest_end_date']):
                        group['latest_end_date'] = end_date
                
                # Store first row data for other fields
                if group['first_row_data'] is None:
                    group['first_row_data'] = row
            
            # Determine the earliest start date for this group for sorting
            earliest_start_date = None
            for row_index in group['rows']:
                row = df.loc[row_index]
                start_date = pd.to_datetime(row.get(COLUMN_MAP['date_start']), errors='coerce')
                if not pd.isna(start_date):
                    if earliest_start_date is None or start_date < earliest_start_date:
                        earliest_start_date = start_date
            group['earliest_start_date'] = earliest_start_date
        
        # Sort service agreement groups by company and date (earliest start date first within each company)
        # Create a list of (timestamp, group) tuples sorted by company and date
        sorted_groups = []
        for timestamp, group in service_agreement_groups.items():
            earliest_date = group.get('earliest_start_date')
            company_code = group['company_code']
            if earliest_date is not None:
                sorted_groups.append((company_code, earliest_date, timestamp, group))
            else:
                # If no valid date found, put at the end
                sorted_groups.append((company_code, pd.Timestamp.max, timestamp, group))
        
        # Sort by company first, then by date
        sorted_groups.sort(key=lambda x: (x[0], x[1]))
        
        print(f"Service agreements will be sequenced by company and date order:")
        for company_code, earliest_date, timestamp, group in sorted_groups:
            date_str = earliest_date.strftime('%d.%m.%Y') if earliest_date != pd.Timestamp.max else 'No date'
            print(f"  {company_code}: {date_str} (timestamp {timestamp}), Money {group['money']}")
        
        # Now assign sequence numbers by company in date order
        for company_code, earliest_date, timestamp, group in sorted_groups:
            group['sequence_number'] = company_service_sequence_counters[company_code]
            company_service_sequence_counters[company_code] += 1
        
        print(f"Found {len(service_agreement_groups)} unique service agreement groups by timestamp")
        for timestamp, group in service_agreement_groups.items():
            routes_count = len(group.get('routes', []))
            freights_count = len(group.get('freights', []))
            sequence_number = group.get('sequence_number', 'Unknown')
            print(f"  Timestamp {timestamp}: {len(group['rows'])} rows, Company: {group['company_code']}, Money: {group['money']}, Company Sequence: цпх{sequence_number}")
            print(f"    Routes: {routes_count}, Freights: {freights_count}")

        # --- Build АВТОСТРАДА per-day service document groups ---
        # Group АВТОСТРАДА rows by (driver, company, date) - one service doc per driver per day
        avtostrada_groups = {}
        if has_customer_col:
            for index, row in df.iterrows():
                if not is_avtostrada_row(row.get(COLUMN_MAP['customer'])):
                    continue

                driver_name = str(row.get(COLUMN_MAP['driver_full'], '')).strip()
                company_code = str(row.get(COLUMN_MAP['company_code'], '')).strip().upper()
                date_obj = pd.to_datetime(row.get(COLUMN_MAP['date_start']), errors='coerce')

                if not driver_name or not company_code or pd.isna(date_obj):
                    continue

                date_str = date_obj.strftime('%d.%m.%Y')
                key = (driver_name, company_code, date_str)

                # Key = (driver, company) → ONE document per driver per company;
                # each day inside becomes one ${certificate-sequence} page.
                key = (driver_name, company_code)

                if key not in avtostrada_groups:
                    avtostrada_groups[key] = {
                        'driver_full': driver_name,
                        'company_code': company_code,
                        'rows': [],
                        'days': {},             # date_str → per-day data dict
                        '_sa_timestamps': {},   # {timestamp: money}
                        'per_day_money': 0.0,   # filled after all groups are built
                        'first_date': None,
                        'latest_end_date': None,
                        'first_row_data': None,
                        'sequence_number': None,
                    }

                group = avtostrada_groups[key]
                group['rows'].append(index)

                # Per-day data (routes/freights deduplicated within the day)
                if date_str not in group['days']:
                    group['days'][date_str] = {
                        'routes': [],
                        'freights': [],
                        '_seen_routes': set(),
                        '_seen_freights': set(),
                        'latest_end_date': None,
                    }
                day = group['days'][date_str]

                route = str(row.get(COLUMN_MAP['route_desc'], '')).strip()
                if route and route not in day['_seen_routes']:
                    day['routes'].append(route)
                    day['_seen_routes'].add(route)
                freight = str(row.get(COLUMN_MAP['freight'], '')).strip()
                if freight and freight not in day['_seen_freights']:
                    day['freights'].append(freight)
                    day['_seen_freights'].add(freight)

                # Track end dates (per-day and global)
                end_date = pd.to_datetime(row.get(COLUMN_MAP['date_end']), errors='coerce')
                if not pd.isna(end_date):
                    if day['latest_end_date'] is None or end_date > day['latest_end_date']:
                        day['latest_end_date'] = end_date
                    if group['latest_end_date'] is None or end_date > group['latest_end_date']:
                        group['latest_end_date'] = end_date

                # Track first date for filename
                date_obj_cur = pd.to_datetime(date_str, format='%d.%m.%Y', errors='coerce')
                if not pd.isna(date_obj_cur):
                    if group['first_date'] is None or date_obj_cur < group['first_date']:
                        group['first_date'] = date_obj_cur

                # Record SA timestamp → money
                if has_comment_col:
                    param = row.get(COLUMN_MAP['comment'])
                    sa_details = extract_service_agreement_details(param)
                    if sa_details:
                        timestamp = sa_details['timestamp']
                        if not timestamp:
                            driver_name = str(row.get(COLUMN_MAP['driver_full'], '')).strip()
                            driver_clean = re.sub(r'[^\w\s]', '', driver_name).replace(' ', '_').lower()
                            timestamp = f"{driver_clean}_{sa_details['company']}".lower()
                        try:
                            group['_sa_timestamps'][timestamp] = float(sa_details['money'])
                        except (ValueError, TypeError):
                            pass

                if group['first_row_data'] is None:
                    group['first_row_data'] = row

        # Compute per_day_money:
        # contract_money / (num_avtostrada_days + num_sa_rows_sharing_same_timestamp)
        from collections import defaultdict
        ts_to_keys = defaultdict(list)
        for key, group in avtostrada_groups.items():
            driver_full, company_code = key
            for ts in group['_sa_timestamps']:
                ts_to_keys[(driver_full, company_code, ts)].append(key)

        for (driver_full, company_code, ts), group_keys in ts_to_keys.items():
            # Each (driver, company) has exactly one group; group_keys has one entry
            group = avtostrada_groups[group_keys[0]]
            contract_money = group['_sa_timestamps'][ts]
            num_avtostrada_days = len(group['days'])
            num_sa_rows = len(service_agreement_groups[ts]['rows']) if ts in service_agreement_groups else 0
            total_units = num_avtostrada_days + num_sa_rows
            group['per_day_money'] += contract_money / total_units if total_units else 0

        # Pre-merge detection: if an АВТОСТРАДА group shares a SA timestamp with a regular SA
        # group for the same driver+company, it will be folded into that SA document.
        # Mark those groups NOW so they don't get their own sequence numbers.
        for timestamp, sa_group in service_agreement_groups.items():
            if sa_group['first_row_data'] is None:
                continue
            first_row = sa_group['first_row_data']
            param_driver = extract_driver_from_parameter(first_row.get(COLUMN_MAP.get('comment', ''))) if has_comment_col else None
            sa_driver = param_driver if param_driver else str(first_row.get(COLUMN_MAP['driver_full'], '')).strip()
            sa_company = sa_group['company_code']
            avto_key = (sa_driver, sa_company)
            avto_grp = avtostrada_groups.get(avto_key)
            if avto_grp and timestamp in avto_grp.get('_sa_timestamps', {}):
                avto_grp['_merged_into_sa'] = timestamp   # which SA absorbs it
                sa_group['_merged_avto_key'] = avto_key   # back-reference

        # Assign sequence numbers — skip groups merged into a SA document
        for key in sorted(avtostrada_groups.keys(), key=lambda k: (k[1], k[0])):
            driver_full, company_code = key
            group = avtostrada_groups[key]
            if group.get('_merged_into_sa'):
                group['sequence_number'] = None
                continue
            if company_code not in company_service_sequence_counters:
                company_service_sequence_counters[company_code] = get_company_service_sequence_number(company_code)
            group['sequence_number'] = company_service_sequence_counters[company_code]
            company_service_sequence_counters[company_code] += 1

        print(f"Found {len(avtostrada_groups)} АВТОСТРАДА groups:")
        for key in sorted(avtostrada_groups.keys(), key=lambda k: (k[1], k[0])):
            driver_full, company_code = key
            group = avtostrada_groups[key]
            num_days = len(group['days'])
            if group.get('_merged_into_sa'):
                ts = group['_merged_into_sa']
                sa_seq = service_agreement_groups[ts].get('sequence_number', '?')
                print(f"  {company_code} {format_driver_initials(driver_full)}: {len(group['rows'])} trips, {num_days} days -> MERGED into цпх{sa_seq}")
            else:
                print(f"  {company_code} {format_driver_initials(driver_full)}: {len(group['rows'])} trips, {num_days} days, цпх{group['sequence_number']}")

        for company in unique_companies:
            company_upper = str(company).strip().upper()
            company_word_path = OUTPUT_ROOT_PATH / company_upper / month_year_str
            
            # Get the next available sequence number for regular assignments
            next_seq = get_global_sequence_number(company_word_path, company_upper, target_year, target_month, False)
            company_sequence_counters[company_upper] = next_seq
            print(f"  {company_upper}: regular assignments starting from sequence {next_seq}")
        
        print(f"Regular sequence counters initialized: {company_sequence_counters}")

        # --- Initialize Word Application ---
        print("Initializing Word Application...")
        try:
            # Use EnsureDispatch for better caching and type library access
            word_app = win32com.client.gencache.EnsureDispatch("Word.Application")
        except AttributeError:
             # Fallback if gencache is problematic (rare)
            try:
                word_app = win32com.client.Dispatch("Word.Application")
            except Exception as e_dispatch:
                 print(f"Fatal Error: Could not start Word Application.")
                 print(f"Ensure Microsoft Word is installed and pywin32 is configured correctly.")
                 print(f"Error details: {e_dispatch}")
                 sys.exit(1) # Exit script if Word cannot be started

        word_app.Visible = False # Run Word in the background

        # --- Initialize Excel Application ---
        if args.no_excel_update:
            print("Excel update disabled by command line argument")
            excel_app = None
        else:
            print("Initializing Excel Application...")
            try:
                excel_app = win32com.client.gencache.EnsureDispatch("Excel.Application")
            except AttributeError:
                try:
                    excel_app = win32com.client.Dispatch("Excel.Application")
                except Exception as e_excel_dispatch:
                    print(f"Warning: Could not start Excel Application for writing back sequence numbers.")                
                    print(f"Files will still be created, but sequence numbers won't be written to Excel.")
                    print(f"Error details: {e_excel_dispatch}")
                    excel_app = None

        if excel_app:
            excel_app.Visible = False # Run Excel in the background
            excel_app.DisplayAlerts = False # Suppress Excel alerts
            
            # Find the column index for "накази" column
            try:
                temp_wb = excel_app.Workbooks.Open(excel_file_path)
                temp_ws = temp_wb.Worksheets(SHEET_NAME)
                excel_assignments_col_idx = find_excel_column_index(temp_ws, COLUMN_MAP['assignments_col']) if COLUMN_MAP.get('assignments_col') else None
                temp_wb.Close(SaveChanges=False)
                
                if excel_assignments_col_idx:
                    print(f"Found '{COLUMN_MAP.get('assignments_col')}' column at index {excel_assignments_col_idx}")
                else:
                    print(f"Column for writing assignments back not configured or not found in Excel sheet")
                    excel_app.Quit()
                    excel_app = None
                    
            except Exception as e_col_find:
                print(f"Warning: Error finding column index: {e_col_find}")
                excel_app.Quit()
                excel_app = None

        # --- Process service agreements first (one per timestamp) in sequence order ---
        processed_service_agreements = set()
        
        # --- Setup Excel batch writer ---
        excel_batch_writer = None
        if excel_app and excel_assignments_col_idx:
            excel_batch_writer = ExcelBatchWriter(excel_app, excel_file_path, SHEET_NAME, excel_assignments_col_idx)
        
        # --- Process service agreements first in sequence number order ---
        with excel_batch_writer if excel_batch_writer else nullcontext() as excel_writer:
            # Process service agreements in sequence order (цпх1, цпх2, etc.)
            for company_code, earliest_date, timestamp, group in sorted_groups:
                if timestamp in service_agreement_groups:
                    # Get the first row from this service agreement group for processing
                    first_row_index = group['rows'][0]
                    row = df.loc[first_row_index]
                    index = first_row_index
                    
                    print(f"\nProcessing service agreement цпх{group['sequence_number']} (timestamp {timestamp})...")
                    
                    # Mark this timestamp as processed
                    processed_service_agreements.add(timestamp)
                    
                    # Extract data using COLUMN_MAP
                    data = {key: row.get(excel_col) for key, excel_col in COLUMN_MAP.items()}
                    
                    # Set up service agreement data
                    sa_details = extract_service_agreement_details(data.get('comment'))
                    if not sa_details:
                        print(f"Skipping service agreement: Invalid format")
                        continue
                    
                    # Add formatted data
                    data['money_formatted'] = format_money_ukrainian(sa_details['money'])
                    data['route_contract'] = format_routes_contract(group.get('routes', []))
                    data['freight_contract'] = format_freight_contract(group.get('freights', []))
                    
                    # Use latest end date if available
                    latest_end_date = group.get('latest_end_date')
                    if latest_end_date and not pd.isna(latest_end_date):
                        data['date_end'] = latest_end_date.strftime('%d.%m.%Y')
                    
                    # Use first row data for other fields if current row is missing data
                    first_row_data = group.get('first_row_data')
                    if first_row_data is not None:
                        for key, excel_col in COLUMN_MAP.items():
                            if not data.get(key) or pd.isna(data.get(key)):
                                data[key] = first_row_data.get(excel_col)
                    
                    # Process this service agreement (using existing logic)
                    is_service_row = True
                    
                    # Get essential values
                    company_code = str(data.get('company_code', '')).strip().upper()
                    driver_name = str(data.get('driver_full', '')).strip()
                    
                    # Check for driver override in parameter
                    param_driver = extract_driver_from_parameter(data.get('comment'))
                    if param_driver:
                        driver_name = param_driver
                        print(f"Using driver name from parameter: '{driver_name}'")
                        data['driver_full'] = driver_name
                    
                    # Handle dates
                    date_obj_start = pd.to_datetime(data.get('date_start'), errors='coerce')
                    if isinstance(data.get('date_end'), str) and data.get('date_end'):
                        date_obj_end = pd.to_datetime(data.get('date_end'), format='%d.%m.%Y', errors='coerce')
                    else:
                        date_obj_end = pd.to_datetime(data.get('date_end'), errors='coerce')
                    
                    # Validate essential data
                    if not company_code or not driver_name or pd.isna(date_obj_start):
                        print(f"Skipping service agreement: Missing essential data")
                        continue
                    
                    # Format dates
                    date_str_ddmmyyyy = date_obj_start.strftime('%d.%m.%Y')
                    data['date_start'] = date_obj_start.strftime('%d.%m.%Y')
                    if not pd.isna(date_obj_end):
                        if not isinstance(data.get('date_end'), str):
                            data['date_end'] = date_obj_end.strftime('%d.%m.%Y')
                    
                    # Use service agreement template based on company
                    template_filename = SERVICE_AGREEMENT_TEMPLATES.get(company_code)
                    if not template_filename:
                        print(f"Skipping service agreement: No template found for company '{company_code}'")
                        continue
                    template_path = get_template_path(template_filename)
                    if not template_path.is_file():
                        print(f"Skipping service agreement: Template not found at '{template_path}'")
                        continue
                    
                    # Create output folders with company-specific names
                    service_word_path = OUTPUT_ROOT_PATH / f"{company_code}-ЦПХ"
                    service_pdf_path = OUTPUT_ROOT_PATH / f"{company_code}-ЦПХ-pdf"
                    service_word_path.mkdir(parents=True, exist_ok=True)
                    service_pdf_path.mkdir(parents=True, exist_ok=True)
                    
                    # Format driver initials
                    driver_initials = format_driver_initials(driver_name)
                    
                    # Get sequence number
                    sequence_number = group.get('sequence_number', 1)
                    seq_prefix = "цпх"
                    seq_suffix = str(sequence_number)
                    
                    # Check if file already exists
                    potential_filename = service_pdf_path / f"{company_code} {driver_initials} від {date_str_ddmmyyyy} {seq_prefix}{seq_suffix}.pdf"
                    if potential_filename.exists():
                        print(f"Skipping service agreement: File '{potential_filename}' already exists.")
                        continue
                    
                    # Construct file names
                    base_filename = f"{company_code} {driver_initials} від {date_str_ddmmyyyy} {seq_prefix}{seq_suffix}"
                    docx_filename = service_word_path / f"{base_filename}.docx"
                    pdf_filename = service_pdf_path / f"{base_filename}.pdf"
                    
                    # Process the document (existing logic)
                    word_doc = None
                    try:
                        print(f"Opening template: {template_path}")
                        word_doc = word_app.Documents.Open(str(template_path))
                        print(f"Replacing placeholders for цпх{sequence_number}, Driver: {driver_initials}")
                        
                        # Check whether an АВТОСТРАДА group is merged into this SA document
                        merged_avto_key = group.get('_merged_avto_key')
                        merged_avto = avtostrada_groups.get(merged_avto_key) if merged_avto_key else None

                        # Build certificate row_data: SA trips first, then АВТОСТРАДА days
                        sa_row_data = group.get('row_data', [])
                        avto_row_data = []
                        if merged_avto:
                            sorted_avto_dates = sorted(
                                merged_avto['days'].keys(),
                                key=lambda d: pd.to_datetime(d, format='%d.%m.%Y', errors='coerce')
                            )
                            for ds in sorted_avto_dates:
                                day = merged_avto['days'][ds]
                                day_end = day['latest_end_date']
                                avto_row_data.append({
                                    'date_start': ds,
                                    'date_end': day_end.strftime('%d.%m.%Y') if day_end else ds,
                                    'route_desc': format_routes_contract(day['routes']),
                                    'freight': format_freight_contract(day['freights']),
                                })
                            print(f"Merging {len(avto_row_data)} АВТОСТРАДА day pages into цпх{sequence_number}")

                            # Rebuild the first-page summary placeholders (${route-contract},
                            # ${freight}) so they list exactly ONE numbered item per certificate
                            # page: the SA trips (one each) followed by the АВТОСТРАДА days. A day
                            # with several trips is collapsed to one comma-joined item, so the
                            # contract item count matches the certificate page count.
                            contract_routes = [rd.get('route_desc', '') for rd in sa_row_data]
                            contract_freights = [rd.get('freight', '') for rd in sa_row_data]
                            for ds in sorted_avto_dates:
                                day = merged_avto['days'][ds]
                                contract_routes.append(join_day_items(day['routes']))
                                contract_freights.append(join_day_items(day['freights']))
                            data['route_contract'] = format_routes_contract(contract_routes)
                            data['freight_contract'] = format_freight_contract(contract_freights)

                        combined_row_data = sa_row_data + avto_row_data
                        num_rows = len(combined_row_data)

                        # Build a merged group dict for duplicate_certificate_pages
                        cert_group = dict(group)
                        cert_group['row_data'] = combined_row_data

                        # Replace placeholders
                        find_replace(word_doc, "${sequence}", str(sequence_number))
                        find_replace(word_doc, "${shortdname}", driver_initials)

                        individual_number = get_driver_individual_number(excel_file_path, driver_name)
                        find_replace(word_doc, "${individualnumber}", individual_number)

                        for placeholder, data_key in SERVICE_PLACEHOLDERS.items():
                            replace_value = data.get(data_key, '')
                            if replace_value:
                                find_replace(word_doc, placeholder, replace_value)
                            else:
                                print(f"Warning: No data for service placeholder '{placeholder}' (key: {data_key})")

                        # Duplicate certificate pages (SA trips + АВТОСТРАДА days combined)
                        print(f"Service agreement has {num_rows} total certificate pages ({len(sa_row_data)} SA + {len(avto_row_data)} АВТОСТРАДА)")
                        duplicate_certificate_pages(word_doc, max(num_rows, 1), cert_group)

                        # Save documents
                        print(f"Saving service agreement DOCX: {docx_filename}")
                        word_doc.SaveAs2(str(docx_filename))

                        print(f"Saving service agreement PDF: {pdf_filename}")
                        word_doc.SaveAs2(str(pdf_filename), FileFormat=WD_FORMAT_PDF)

                        # Small delay to ensure PDF is completely written
                        time.sleep(0.5)

                        # Remove blank pages from the PDF
                        if remove_blank_pages:
                            print(f"Checking for blank pages in service agreement PDF...")
                            remove_blank_pages_from_pdf(str(pdf_filename))
                        else:
                            print(f"Blank page removal is disabled")

                        report_lines.append(f"Service Agreement ({len(sa_row_data)} SA + {len(avto_row_data)} АВТО pages): {pdf_filename.name}")

                        # Write sequence number to SA Excel rows
                        if excel_writer:
                            try:
                                excel_row_num = index + 2
                                success = excel_writer.add_sequence_number(
                                    row_num=excel_row_num,
                                    sequence_num=f"{seq_prefix}{seq_suffix}",
                                    hyperlink_path=str(docx_filename)
                                )
                                if success:
                                    print(f"Successfully wrote service agreement identifier {seq_prefix}{seq_suffix} with hyperlink to Excel row {excel_row_num}")
                                else:
                                    print(f"Warning: Failed to write service agreement identifier to Excel for row {excel_row_num}")
                            except Exception as e_excel_write:
                                print(f"Warning: Failed to write back to Excel for row {index + 2}: {e_excel_write}")

                        # Also write same цпх to all merged АВТОСТРАДА rows
                        if merged_avto and excel_writer:
                            for avto_row_idx in merged_avto['rows']:
                                try:
                                    excel_writer.add_sequence_number(
                                        row_num=avto_row_idx + 2,
                                        sequence_num=f"{seq_prefix}{seq_suffix}",
                                        hyperlink_path=str(docx_filename)
                                    )
                                except Exception as e_avto_excel:
                                    print(f"Warning: Failed to write to АВТОСТРАДА Excel row {avto_row_idx + 2}: {e_avto_excel}")
                        
                        # Update file counts
                        company_file_counts[company_code] = company_file_counts.get(company_code, 0) + 1
                        print(f"Successfully created service agreement цпх{sequence_number}.")
                        
                        # Close document
                        word_doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                        word_doc = None
                        
                    except Exception as e_doc:
                        print(f"ERROR processing service agreement цпх{sequence_number}: {e_doc}")
                        if word_doc:
                            try:
                                word_doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                            except Exception as e_close:
                                print(f"Warning: Could not close document after error: {e_close}")
                            word_doc = None
                    finally:
                        time.sleep(0.1)

            # --- Process АВТОСТРАДА service documents (one per driver, certificate page per day) ---
            if avtostrada_groups:
                print(f"\n=== PROCESSING АВТОСТРАДА SERVICE DOCUMENTS ===")

                for key in sorted(avtostrada_groups.keys(), key=lambda k: (k[1], k[0])):
                    driver_full, company_code = key
                    group = avtostrada_groups[key]

                    # Skip groups already merged into a regular SA document
                    if group.get('_merged_into_sa'):
                        print(f"\nSkipping АВТОСТРАДА {format_driver_initials(driver_full)}: merged into цпх of SA timestamp {group['_merged_into_sa']}")
                        continue

                    num_trips = len(group['rows'])
                    sorted_dates = sorted(group['days'].keys(),
                                         key=lambda d: pd.to_datetime(d, format='%d.%m.%Y', errors='coerce'))
                    num_days = len(sorted_dates)
                    per_day_money = group['per_day_money']
                    total_money = per_day_money * num_days

                    print(f"\nАВТОСТРАДА: {format_driver_initials(driver_full)} ({num_trips} trips, {num_days} days, цпх{group['sequence_number']})...")

                    template_filename = SERVICE_AGREEMENT_TEMPLATES.get(company_code)
                    if not template_filename:
                        print(f"Skipping: No service template for company '{company_code}'")
                        continue
                    template_path = get_template_path(template_filename)
                    if not template_path.is_file():
                        print(f"Skipping: Template not found at '{template_path}'")
                        continue

                    service_word_path = OUTPUT_ROOT_PATH / f"{company_code}-ЦПХ"
                    service_pdf_path = OUTPUT_ROOT_PATH / f"{company_code}-ЦПХ-pdf"
                    service_word_path.mkdir(parents=True, exist_ok=True)
                    service_pdf_path.mkdir(parents=True, exist_ok=True)

                    driver_initials = format_driver_initials(driver_full)
                    sequence_number = group['sequence_number']
                    seq_prefix = "цпх"
                    first_date_str = sorted_dates[0] if sorted_dates else ''
                    base_filename = f"{company_code} {driver_initials} від {first_date_str} {seq_prefix}{sequence_number}"
                    docx_filename = service_word_path / f"{base_filename}.docx"
                    pdf_filename = service_pdf_path / f"{base_filename}.pdf"

                    if pdf_filename.exists():
                        print(f"Skipping: File '{pdf_filename}' already exists.")
                        continue

                    # Document-level date range. The first-page summary placeholders list
                    # exactly one numbered item per certificate page (one per day); a day with
                    # several trips is collapsed to one comma-joined item via join_day_items,
                    # so the contract item count matches the certificate page count.
                    latest_end = group['latest_end_date']
                    end_date_str = latest_end.strftime('%d.%m.%Y') if latest_end and not pd.isna(latest_end) else first_date_str
                    contract_routes = []
                    contract_freights = []
                    for ds in sorted_dates:
                        day = group['days'][ds]
                        contract_routes.append(join_day_items(day['routes']))
                        contract_freights.append(join_day_items(day['freights']))
                    route_contract = format_routes_contract(contract_routes)
                    freight_contract = format_freight_contract(contract_freights)
                    money_formatted = format_money_ukrainian(str(total_money)) if total_money > 0 else '0 гривень 00 коп.'
                    individual_number = get_driver_individual_number(excel_file_path, driver_full)

                    # Build one row_data entry per day for certificate pages
                    row_data = []
                    for ds in sorted_dates:
                        day = group['days'][ds]
                        day_end = day['latest_end_date']
                        day_end_str = day_end.strftime('%d.%m.%Y') if day_end else ds
                        row_data.append({
                            'date_start': ds,
                            'date_end': day_end_str,
                            'route_desc': format_routes_contract(day['routes']),
                            'freight': format_freight_contract(day['freights']),
                        })

                    word_doc = None
                    try:
                        print(f"Opening template: {template_path}")
                        word_doc = word_app.Documents.Open(str(template_path))

                        find_replace(word_doc, "${sequence}", str(sequence_number))
                        find_replace(word_doc, "${shortdname}", driver_initials)
                        find_replace(word_doc, "${individualnumber}", individual_number)
                        find_replace(word_doc, "${date}", first_date_str)
                        find_replace(word_doc, "${driver}", driver_full)
                        find_replace(word_doc, "${route-contract}", route_contract)
                        find_replace(word_doc, "${start}", first_date_str)
                        find_replace(word_doc, "${end}", end_date_str)
                        find_replace(word_doc, "${freight}", freight_contract)
                        find_replace(word_doc, "${money}", money_formatted)

                        # One ${certificate-sequence} page per day; money = total / num_days per page
                        cert_data = {
                            'money': str(total_money),
                            'row_data': row_data,
                        }
                        duplicate_certificate_pages(word_doc, num_days, cert_data)

                        print(f"Saving АВТОСТРАДА DOCX: {docx_filename}")
                        word_doc.SaveAs2(str(docx_filename))

                        print(f"Saving АВТОСТРАДА PDF: {pdf_filename}")
                        word_doc.SaveAs2(str(pdf_filename), FileFormat=WD_FORMAT_PDF)

                        time.sleep(0.5)

                        if remove_blank_pages:
                            remove_blank_pages_from_pdf(str(pdf_filename))

                        report_lines.append(f"АВТОСТРАДА Service ({num_trips} trips, {num_days} days): {pdf_filename.name}")

                        # Write same цпх identifier to every row in the group
                        if excel_writer:
                            for row_idx in group['rows']:
                                try:
                                    success = excel_writer.add_sequence_number(
                                        row_num=row_idx + 2,
                                        sequence_num=f"{seq_prefix}{sequence_number}",
                                        hyperlink_path=str(docx_filename)
                                    )
                                    if success:
                                        print(f"Wrote {seq_prefix}{sequence_number} to Excel row {row_idx + 2}")
                                except Exception as e_excel:
                                    print(f"Warning: Failed to write to Excel row {row_idx + 2}: {e_excel}")

                        company_file_counts[company_code] = company_file_counts.get(company_code, 0) + 1
                        print(f"Successfully created АВТОСТРАДА service doc цпх{sequence_number} ({num_days} certificate pages)")

                        word_doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                        word_doc = None

                    except Exception as e_doc:
                        print(f"ERROR processing АВТОСТРАДА service doc цпх{sequence_number}: {e_doc}")
                        print(traceback.format_exc())
                        if word_doc:
                            try:
                                word_doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                            except:
                                pass
                            word_doc = None
                    finally:
                        time.sleep(0.1)

            # --- Process regular assignments ---
            # If consolidation mode is enabled, group by date and company first
            # Variables to store consolidation statistics for report
            consolidation_stats = None
            consolidated_pdf_files = []  # Track all consolidated PDFs for final merge
            driver_assignments_index = {}  # Track which assignments each driver is involved in {driver_initials: [seq_nums]}
            
            if args.consolidate:
                print(f"\n=== CONSOLIDATION MODE: Grouping assignments by date and company ===")
                
                # Filter out service agreements and АВТОСТРАДА rows from regular assignments
                df_regular = df.copy()
                if has_comment_col:
                    df_regular = df_regular[~df_regular[COLUMN_MAP['comment']].apply(is_service_agreement)]
                if has_customer_col:
                    df_regular = df_regular[~df_regular[COLUMN_MAP['customer']].apply(is_avtostrada_row)]
                
                # Group assignments by company and date
                consolidated_groups = group_assignments_by_date_and_company(df_regular, COLUMN_MAP)
                
                # Track already processed row indices
                processed_rows = set()
                
                print(f"Found {len(consolidated_groups)} unique date/company combinations")
                
                # Calculate paper savings
                total_assignments_individual = len(df_regular)
                total_assignments_consolidated = len(consolidated_groups)
                consolidated_groups_count = len([g for g in consolidated_groups.values() if len(g) > 1])
                papers_saved = total_assignments_individual - total_assignments_consolidated
                
                # Store statistics for final report
                consolidation_stats = {
                    'unique_combinations': len(consolidated_groups),
                    'individual': total_assignments_individual,
                    'consolidated': total_assignments_consolidated,
                    'multi_driver_groups': consolidated_groups_count,
                    'papers_saved': papers_saved
                }
                
                print(f"Individual assignments: {total_assignments_individual}")
                print(f"Consolidated documents: {total_assignments_consolidated}")
                print(f"Groups with 2+ drivers: {consolidated_groups_count}")
                print(f"PAPER SAVINGS: {papers_saved} fewer documents ({papers_saved} papers saved)")
                
                # Process each consolidated group
                for (company_code, date_str), row_indices in sorted(consolidated_groups.items()):
                    num_drivers = len(row_indices)
                    
                    if num_drivers == 0:
                        continue
                    
                    print(f"\n{'='*60}")
                    print(f"Processing consolidated group: {company_code} on {date_str} ({num_drivers} drivers)")
                    print(f"{'='*60}")
                    
                    # Collect driver data for this group
                    driver_data_list = []
                    all_rows_data = []
                    
                    for idx in row_indices:
                        row = df_regular.loc[idx]
                        data = {key: row.get(excel_col) for key, excel_col in COLUMN_MAP.items()}
                        
                        # Get driver info
                        driver_name = str(data.get('driver_full', '')).strip()
                        
                        # Check for driver override in parameter
                        param_driver = extract_driver_from_parameter(data.get('comment'))
                        if param_driver:
                            driver_name = param_driver
                            data['driver_full'] = driver_name
                        
                        # Validate essential data
                        if not driver_name:
                            print(f"Skipping row {idx + 2}: Missing driver name")
                            continue
                        
                        # Format dates
                        date_obj_start = pd.to_datetime(data.get('date_start'), errors='coerce')
                        date_obj_end = pd.to_datetime(data.get('date_end'), errors='coerce')
                        
                        if pd.isna(date_obj_start):
                            print(f"Skipping row {idx + 2}: Invalid start date")
                            continue
                        
                        driver_initials = format_driver_initials(driver_name)
                        
                        driver_info = {
                            'driver_full': driver_name,
                            'driver_initials': driver_initials,
                            'truck_model': str(data.get('truck_model', '')).strip(),
                            'plate_number': str(data.get('plate_number', '')).strip(),
                            'route_desc': str(data.get('route_desc', '')).strip(),
                            'date_start': date_obj_start.strftime('%d.%m.%Y'),
                            'date_end': date_obj_end.strftime('%d.%m.%Y') if not pd.isna(date_obj_end) else '',
                            'row_index': idx
                        }
                        
                        driver_data_list.append(driver_info)
                        all_rows_data.append((idx, data))
                        processed_rows.add(idx)
                    
                    if not driver_data_list:
                        print(f"No valid drivers in this group, skipping")
                        continue
                    
                    # Sort drivers alphabetically
                    driver_data_list.sort(key=lambda x: x['driver_full'])
                    
                    # Get template
                    template_filename = TEMPLATES.get(company_code)
                    if not template_filename:
                        print(f"Skipping consolidated group: No template for company '{company_code}'")
                        continue
                    
                    template_path = get_template_path(template_filename)
                    if not template_path.is_file():
                        print(f"Skipping consolidated group: Template not found at '{template_path}'")
                        continue
                    
                    # Create output folders
                    company_word_path = OUTPUT_ROOT_PATH / company_code / month_year_str
                    company_pdf_path = OUTPUT_ROOT_PATH / f"{company_code}-pdf" / month_year_str
                    company_word_path.mkdir(parents=True, exist_ok=True)
                    company_pdf_path.mkdir(parents=True, exist_ok=True)
                    
                    # Get sequence number
                    next_seq_num = company_sequence_counters[company_code]
                    company_sequence_counters[company_code] += 1
                    
                    # Create consolidated filename with full driver names (not initials)
                    driver_names_list = [d['driver_initials'] for d in driver_data_list]
                    base_filename = create_consolidated_filename(company_code, driver_names_list, date_str, next_seq_num)
                    docx_filename = truncate_path_to_limit(company_word_path / f"{base_filename}.docx")
                    pdf_filename = truncate_path_to_limit(company_pdf_path / f"{base_filename}.pdf")
                    
                    # Check if file exists
                    if docx_filename.exists():
                        print(f"Skipping: File '{docx_filename}' already exists")
                        # Still add the PDF to tracking list for stats
                        if pdf_filename.exists():
                            consolidated_pdf_files.append(str(pdf_filename))
                        continue
                    
                    # Process the consolidated document
                    word_doc = None
                    try:
                        print(f"Opening template: {template_path}")
                        word_doc = word_app.Documents.Open(str(template_path))
                        print(f"Creating consolidated assignment for {num_drivers} drivers")
                        
                        # Replace common placeholders (Date, Sequence)
                        find_replace(word_doc, "Sequence", str(next_seq_num))
                        # Format date with Ukrainian month name
                        date_ukrainian = format_date_ukrainian(date_str)
                        find_replace(word_doc, "Date", date_ukrainian)
                        find_replace(word_doc, "start", date_str)  # If template has a start placeholder
                        
                        # Insert multiple driver lines
                        if not insert_multiple_driver_lines(word_doc, driver_data_list, args.short_names):
                            print("Warning: Failed to insert multiple driver lines, continuing anyway")
                        
                        # Save documents
                        print(f"Saving consolidated DOCX: {docx_filename}")
                        word_doc.SaveAs2(str(docx_filename))
                        
                        print(f"Saving consolidated PDF: {pdf_filename}")
                        word_doc.SaveAs2(str(pdf_filename), FileFormat=WD_FORMAT_PDF)
                        
                        time.sleep(0.5)
                        
                        # Remove blank pages
                        if remove_blank_pages:
                            print(f"Checking for blank pages in consolidated PDF...")
                            remove_blank_pages_from_pdf(str(pdf_filename))
                        
                        report_lines.append(f"Consolidated Assignment ({num_drivers} drivers): {pdf_filename.name}")
                        
                        # Write sequence numbers to Excel for all rows in this group
                        if excel_writer:
                            for driver_info in driver_data_list:
                                try:
                                    excel_row_num = driver_info['row_index'] + 2
                                    success = excel_writer.add_sequence_number(
                                        row_num=excel_row_num,
                                        sequence_num=f"н{next_seq_num}_консолідовано",
                                        hyperlink_path=str(docx_filename)
                                    )
                                    if success:
                                        print(f"Wrote sequence н{next_seq_num}_консолідовано to row {excel_row_num}")
                                except Exception as e_excel:
                                    print(f"Warning: Failed to write to Excel row {driver_info['row_index'] + 2}: {e_excel}")
                        
                        # Update file counts
                        company_file_counts[company_code] = company_file_counts.get(company_code, 0) + 1
                        print(f"Successfully created consolidated assignment н{next_seq_num}")
                        
                        # Track PDF for final merge
                        consolidated_pdf_files.append(str(pdf_filename))
                        
                        # Track driver involvement for signing index
                        # Count trips per driver and store as (sequence_number, trip_count)
                        driver_trip_counts = {}
                        for driver_info in driver_data_list:
                            driver_initials = driver_info['driver_initials']
                            driver_trip_counts[driver_initials] = driver_trip_counts.get(driver_initials, 0) + 1
                        
                        for driver_initials, trip_count in driver_trip_counts.items():
                            if driver_initials not in driver_assignments_index:
                                driver_assignments_index[driver_initials] = []
                            driver_assignments_index[driver_initials].append((next_seq_num, trip_count))
                        
                        # Close document
                        word_doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                        word_doc = None
                        
                    except Exception as e_doc:
                        print(f"ERROR processing consolidated group: {e_doc}")
                        print(traceback.format_exc())
                        if word_doc:
                            try:
                                word_doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                            except:
                                pass
                            word_doc = None
                    finally:
                        time.sleep(0.1)
                
                print(f"\n{'='*60}")
                print(f"CONSOLIDATION COMPLETE")
                print(f"Total papers saved: {papers_saved}")
                print(f"{'='*60}\n")
                
                # --- Create signing index document ---
                signing_index_pdf = None
                if driver_assignments_index:
                    try:
                        print(f"\n{'='*60}")
                        print(f"CREATING SIGNING INDEX")
                        print(f"{'='*60}\n")
                        
                        # Register Ukrainian font
                        try:
                            font_path = r"C:\Windows\Fonts\arial.ttf"
                            pdfmetrics.registerFont(TTFont('Arial', font_path))
                            font_path_bold = r"C:\Windows\Fonts\arialbd.ttf"
                            pdfmetrics.registerFont(TTFont('Arial-Bold', font_path_bold))
                            font_name = 'Arial'
                            font_name_bold = 'Arial-Bold'
                        except Exception as e_font:
                            print(f"Warning: Could not load Arial font: {e_font}")
                            font_name = 'Helvetica'
                            font_name_bold = 'Helvetica-Bold'
                        
                        # Create PDF filename
                        if company_filter:
                            index_company = company_filter
                        else:
                            index_company = "ALL"
                        
                        index_filename = f"SIGNING_INDEX_{index_company}_{target_month:02d}.{target_year}.pdf"
                        if company_filter:
                            month_output_path = OUTPUT_ROOT_PATH / f"{index_company}-pdf" / month_year_str
                        else:
                            month_output_path = OUTPUT_ROOT_PATH / month_year_str
                        month_output_path.mkdir(parents=True, exist_ok=True)
                        signing_index_pdf = month_output_path / index_filename
                        
                        # Create PDF with reportlab
                        can = canvas.Canvas(str(signing_index_pdf), pagesize=A4)
                        page_width, page_height = A4
                        
                        # Add header line with month, year, and company
                        month_names_ukr = {
                            1: 'Січень', 2: 'Лютий', 3: 'Березень', 4: 'Квітень',
                            5: 'Травень', 6: 'Червень', 7: 'Липень', 8: 'Серпень',
                            9: 'Вересень', 10: 'Жовтень', 11: 'Листопад', 12: 'Грудень'
                        }
                        month_name = month_names_ukr.get(target_month, str(target_month))
                        company_name = company_filter if company_filter else "Всі компанії"
                        
                        header_text = f"{month_name} {target_year}, {company_name}"
                        
                        # Draw header (centered, bold, larger font)
                        can.setFont(font_name_bold, 14)
                        header_width = can.stringWidth(header_text, font_name_bold, 14)
                        can.drawString((page_width - header_width) / 2, page_height - 60, header_text)
                        
                        # Starting position for driver list
                        current_y = page_height - 100
                        margin_left = 50
                        checkbox_size = 10
                        line_height = 20
                        
                        # Sort drivers alphabetically
                        sorted_drivers = sorted(driver_assignments_index.items())
                        
                        # Add each driver with checkbox and assignment numbers
                        for driver_initials, seq_data in sorted_drivers:
                            # Check if we need a new page
                            if current_y < 80:
                                can.showPage()
                                current_y = page_height - 60
                                can.setFont(font_name, 12)
                            
                            # Sort by sequence number and format with trip counts
                            seq_data.sort(key=lambda x: x[0])
                            seq_parts = []
                            for seq_num, trip_count in seq_data:
                                if trip_count > 1:
                                    seq_parts.append(f"{seq_num} ({trip_count})")
                                else:
                                    seq_parts.append(str(seq_num))
                            seq_nums_str = ", ".join(seq_parts)
                            assignments_count = len(seq_data)
                            
                            # Draw empty checkbox
                            checkbox_x = margin_left
                            checkbox_y = current_y - checkbox_size
                            can.setStrokeColorRGB(0.2, 0.2, 0.2)
                            can.setFillColorRGB(1, 1, 1)  # White fill for empty checkbox
                            can.setLineWidth(0.8)
                            can.rect(checkbox_x, checkbox_y, checkbox_size, checkbox_size, stroke=1, fill=1)
                            
                            # Draw driver initials (bold) and assignment count
                            text_x = margin_left + checkbox_size + 8
                            can.setFillColorRGB(0, 0, 0)  # Black text
                            can.setFont(font_name_bold, 12)
                            driver_text = f"({assignments_count}) {driver_initials} "
                            driver_text_width = can.stringWidth(driver_text, font_name_bold, 12)
                            
                            # Build wrapped sequence-number lines
                            max_text_width = page_width - text_x - 30  # right margin
                            seq_tokens = seq_parts  # already built above: ["1", "2", "3 (2)", ...]
                            wrap_lines = []
                            cur_line = ""  # sequence tokens for current line
                            cur_width = driver_text_width  # first line starts after driver prefix
                            for ti, token in enumerate(seq_tokens):
                                is_last_token = (ti == len(seq_tokens) - 1)
                                token_str = token + ("." if is_last_token else ", ")
                                token_w = can.stringWidth(token_str, font_name, 12)
                                if cur_line and cur_width + token_w > max_text_width:
                                    wrap_lines.append(cur_line)
                                    cur_line = token_str
                                    cur_width = driver_text_width + token_w
                                else:
                                    cur_line += token_str
                                    cur_width += token_w
                            if cur_line:
                                wrap_lines.append(cur_line)
                            
                            # Draw first line (with driver prefix)
                            can.drawString(text_x, current_y - 8, driver_text)
                            can.setFont(font_name, 12)
                            can.drawString(text_x + driver_text_width, current_y - 8, wrap_lines[0] if wrap_lines else "")
                            current_y -= line_height
                            
                            # Draw continuation lines (indented to align with seq numbers)
                            for extra_line in wrap_lines[1:]:
                                if current_y < 80:
                                    can.showPage()
                                    current_y = page_height - 60
                                    can.setFont(font_name, 12)
                                can.drawString(text_x + driver_text_width, current_y - 8, extra_line)
                                current_y -= line_height
                            
                            print(f"Added: {driver_initials} - assignments {seq_nums_str}")
                        
                        # Save PDF
                        can.save()
                        
                        print(f"\n{'='*60}")
                        print(f"✓ Signing index created successfully!")
                        print(f"Location: {signing_index_pdf}")
                        print(f"Drivers listed: {len(driver_assignments_index)}")
                        print(f"{'='*60}\n")
                        
                    except Exception as e_index:
                        print(f"ERROR creating signing index: {e_index}")
                        print(traceback.format_exc())
                        signing_index_pdf = None
                
                # --- Merge all consolidated PDFs into one final combined PDF ---
                # Collect all consolidated PDFs from disk (not from the tracking list)
                # to handle cases where files were created in previous runs
                try:
                    print(f"\n{'='*60}")
                    print(f"CREATING COMBINED PDF")
                    print(f"{'='*60}\n")
                    
                    # Collect all consolidated PDF files from the company PDF directories
                    all_consolidated_pdfs = []
                    if company_filter:
                        # Single company mode
                        pdf_dir = OUTPUT_ROOT_PATH / f"{company_filter}-pdf" / month_year_str
                        if pdf_dir.exists():
                            all_consolidated_pdfs = list(pdf_dir.glob(f"н*_consolidated {company_filter} від *.pdf"))
                    else:
                        # All companies mode - collect from all company PDF directories
                        for pdf_dir in OUTPUT_ROOT_PATH.glob("*-pdf"):
                            if pdf_dir.is_dir():
                                month_subdir = pdf_dir / month_year_str
                                if month_subdir.is_dir():
                                    all_consolidated_pdfs.extend(month_subdir.glob("н*_consolidated * від *.pdf"))
                    
                    # Sort PDF files by sequence number extracted from filename
                    def extract_sequence_number(filepath):
                        filename = os.path.basename(str(filepath))
                        # Use \D* to skip any non-digit prefix (avoids Cyrillic encoding issues with 'н')
                        match = re.match(r'\D*(\d+)', filename)
                        return int(match.group(1)) if match else 0
                    
                    all_consolidated_pdfs.sort(key=extract_sequence_number)
                    
                    print(f"Found {len(all_consolidated_pdfs)} consolidated PDFs to merge")
                    
                    # Debug: list all files with their sequence numbers and check for duplicates
                    seq_seen = {}
                    for pdf_path in all_consolidated_pdfs:
                        seq = extract_sequence_number(pdf_path)
                        fname = os.path.basename(str(pdf_path))
                        print(f"  [{seq:3d}] {fname[:80]}")
                        if seq in seq_seen:
                            print(f"  ⚠ DUPLICATE seq {seq}! Previous: {seq_seen[seq][:60]}")
                        seq_seen[seq] = fname
                    
                    unique_paths = set(str(p) for p in all_consolidated_pdfs)
                    if len(unique_paths) != len(all_consolidated_pdfs):
                        print(f"  ⚠ DUPLICATE PATHS DETECTED: {len(all_consolidated_pdfs)} total, {len(unique_paths)} unique")
                    
                    if not all_consolidated_pdfs:
                        print("No consolidated PDFs found to merge.")
                        raise Exception("No PDFs to merge")
                    
                    # Create the merged PDF writer
                    merged_pdf = PdfWriter()
                    # IMPORTANT: Keep all PdfReader objects alive until write() is called.
                    # PyPDF2 uses lazy references; if a reader is garbage-collected, page content gets corrupted.
                    active_readers = []
                    
                    # Add signing index as first pages if it exists
                    if signing_index_pdf and signing_index_pdf.exists():
                        print(f"Adding signing index as first pages...")
                        try:
                            index_reader = PdfReader(str(signing_index_pdf))
                            active_readers.append(index_reader)
                            index_page_count = len(index_reader.pages)
                            for page in index_reader.pages:
                                merged_pdf.add_page(page)
                            print(f"✓ Added signing index ({index_page_count} page(s))")
                            
                            # Add blank page after index only in both-sides mode with odd page count
                            if args.both_sides and index_page_count % 2 == 1:
                                blank_page_reader = create_blank_pdf_page()
                                active_readers.append(blank_page_reader)
                                merged_pdf.add_page(blank_page_reader.pages[0])
                                print(f"✓ Added 1 blank separator page after index (both-sides mode)")
                            
                        except Exception as e_index_pdf:
                            print(f"Warning: Could not add signing index: {e_index_pdf}")
                    
                    # Add all consolidated assignment PDFs
                    # With --both-sides: pad odd-page-count assignments with a blank page
                    use_both_sides = args.both_sides
                    for pdf_path in all_consolidated_pdfs:
                        print(f"Adding: {os.path.basename(str(pdf_path))}")
                        try:
                            assignment_reader = PdfReader(str(pdf_path))
                            active_readers.append(assignment_reader)
                            page_count = len(assignment_reader.pages)
                            
                            # Add only content pages, skip blank pages (header/footer only)
                            added_count = 0
                            skipped_blank = 0
                            for page_idx in range(page_count):
                                if is_content_page(str(pdf_path), page_idx):
                                    merged_pdf.add_page(assignment_reader.pages[page_idx])
                                    added_count += 1
                                else:
                                    skipped_blank += 1
                                    print(f"  ⚠ Skipped blank page {page_idx + 1}/{page_count} (header/footer only)")
                            
                            if use_both_sides and added_count % 2 == 1:
                                blank_page_reader = create_blank_pdf_page()
                                active_readers.append(blank_page_reader)
                                merged_pdf.add_page(blank_page_reader.pages[0])
                                print(f"  → Added {added_count} pages + 1 blank page (odd page count, both-sides mode)")
                            else:
                                print(f"  → Added {added_count} pages{f' (removed {skipped_blank} blank)' if skipped_blank else ''}")
                                
                        except Exception as e_pdf:
                            print(f"Warning: Could not add {pdf_path}: {e_pdf}")
                    
                    # Create combined PDF filename with date range
                    if company_filter:
                        combined_company = company_filter
                    else:
                        # Use first company from consolidated files
                        combined_company = "ALL"
                    
                    sides_suffix = "_BOTH_SIDES" if args.both_sides else ""
                    combined_filename = f"CONSOLIDATED_ALL{sides_suffix}_{combined_company}_{target_month:02d}.{target_year}.pdf"
                    if company_filter:
                        month_output_path = OUTPUT_ROOT_PATH / f"{combined_company}-pdf" / month_year_str
                    else:
                        month_output_path = OUTPUT_ROOT_PATH / month_year_str
                    month_output_path.mkdir(parents=True, exist_ok=True)
                    combined_pdf_path = month_output_path / combined_filename
                    
                    # Write the combined PDF
                    with open(combined_pdf_path, 'wb') as output_file:
                        merged_pdf.write(output_file)
                    
                    print(f"\n{'='*60}")
                    mode_label = 'BOTH SIDES' if args.both_sides else 'SINGLE SIDED'
                    print(f"✓ Combined PDF ({mode_label}) created successfully!")
                    print(f"Location: {combined_pdf_path}")
                    try:
                        total_pages = len(PdfReader(str(combined_pdf_path)).pages)
                        print(f"Total pages: {total_pages}")
                    except Exception as e_total:
                        total_pages = None
                        print(f"Warning: Could not read combined PDF page count: {e_total}")
                    print(f"{'='*60}\n")
                    
                    if total_pages is not None:
                        report_lines.append(f"\nCombined PDF ({mode_label}): {combined_filename} ({total_pages} pages)")
                    else:
                        report_lines.append(f"\nCombined PDF ({mode_label}): {combined_filename} (page count unavailable)")
                    
                    # Prompt to print the combined PDF
                    print(f"Press Enter to print {combined_filename}...")
                    input()
                    try:
                        os.startfile(str(combined_pdf_path), "print")
                        print("✓ Sent to printer")
                    except Exception as e_print:
                        print(f"Could not print automatically: {e_print}")
                        print(f"Please print manually from: {combined_pdf_path}")
                    
                except Exception as e_merge:
                    print(f"ERROR creating combined PDF: {e_merge}")
                    print(traceback.format_exc())
                    report_lines.append(f"\nWarning: Failed to create combined PDF: {e_merge}")
            
            else:
                # --- Regular (non-consolidated) assignment processing ---
                for index, row in df.iterrows():
                    print(f"\nProcessing row {index + 2}...") # +2 assumes header row + 0-based index
                    try:
                        # --- Extract data using COLUMN_MAP ---
                        data = {key: row.get(excel_col) for key, excel_col in COLUMN_MAP.items()}
                    
                        # --- Check if this is a service agreement ---
                        is_service_row = has_comment_col and is_service_agreement(data.get('comment'))

                        # Skip service agreements - they were already processed in sequence order
                        if is_service_row:
                            print(f"Skipping row {index + 2}: Service agreement already processed in sequence order")
                            continue

                        # Skip АВТОСТРАДА rows - they are processed as per-day service documents
                        if has_customer_col and is_avtostrada_row(data.get('customer')):
                            print(f"Skipping row {index + 2}: АВТОСТРАДА row - processed as daily service document")
                            continue
                            
                    except Exception as e:
                        print(f"Error extracting data from row {index + 2}: {e}")
                        continue

                    # --- Get Essential Values ---
                    company_code = str(data.get('company_code', '')).strip().upper() # Normalize
                    driver_name = str(data.get('driver_full', '')).strip()
                    
                    # --- Check for driver override in parameter ---
                    param_driver = extract_driver_from_parameter(data.get('comment'))
                    if param_driver:
                        driver_name = param_driver
                        print(f"Using driver name from parameter: '{driver_name}'")
                        # Update the data dictionary with the overridden driver name
                        data['driver_full'] = driver_name

                    # Handle dates - ensure they are datetime objects
                    date_obj_start = pd.to_datetime(data.get('date_start'), errors='coerce')
                    
                    # For service agreements, end date might already be formatted from aggregation
                    if is_service_row and isinstance(data.get('date_end'), str) and data.get('date_end'):
                        # End date already formatted from aggregation
                        date_obj_end = pd.to_datetime(data.get('date_end'), format='%d.%m.%Y', errors='coerce')
                    else:
                        date_obj_end = pd.to_datetime(data.get('date_end'), errors='coerce')

                    # --- Validate Essential Data ---
                    if not company_code:
                        print(f"Skipping row {index + 2}: Missing or invalid Company Code ('{COLUMN_MAP['company_code']}')")
                        continue
                    if not driver_name:
                        print(f"Skipping row {index + 2}: Missing Driver Name ('{COLUMN_MAP['driver_full']}')")
                        continue
                    if pd.isna(date_obj_start):
                         print(f"Skipping row {index + 2}: Missing or invalid Start Date ('{COLUMN_MAP['date_start']}')")
                         continue
                    # End date might be optional depending on logic, check if needed for replacement
                    if pd.isna(date_obj_end) and "end" in PLACEHOLDERS:
                        print(f"Warning: Row {index + 2}: Missing End Date ('{COLUMN_MAP['date_end']}'). Placeholder 'end' might not be filled.")
                        # Set a default or leave as None if find_replace handles it
                        data['date_end'] = "" # Replace with empty string if missing

                    # Format date for filename and potentially for replacement text
                    date_str_ddmmyyyy = date_obj_start.strftime('%d.%m.%Y')
                    # Format dates for Word replacements if specific format needed (e.g., only date part)
                    data['date_start'] = date_obj_start.strftime('%d.%m.%Y') # Or keep as object if Word handles it
                    
                    # Handle end date formatting - might already be formatted for service agreements
                    if not pd.isna(date_obj_end):
                        if not (is_service_row and isinstance(data.get('date_end'), str)):
                            # Only reformat if not already formatted from aggregation
                            data['date_end'] = date_obj_end.strftime('%d.%m.%Y')


                    # --- Determine Template ---
                    # Use regular assignment template
                    template_filename = TEMPLATES.get(company_code)
                    if not template_filename:
                        print(f"Skipping row {index + 2}: Invalid Company Code '{company_code}'. No template mapping found.")
                        continue

                    template_path = get_template_path(template_filename)
                    if not template_path.is_file():
                        print(f"Skipping row {index + 2}: Template file not found at '{template_path}'")
                        continue

                    # --- Create Output Folders ---
                    # Regular assignments use company-specific folders with MMYY subfolder
                    company_word_path = OUTPUT_ROOT_PATH / company_code / month_year_str
                    company_pdf_path = OUTPUT_ROOT_PATH / f"{company_code}-pdf" / month_year_str
                    company_word_path.mkdir(parents=True, exist_ok=True)
                    company_pdf_path.mkdir(parents=True, exist_ok=True)

                    # --- Format Driver Initials ---
                    driver_initials = format_driver_initials(driver_name)

                    # --- Get Sequence Number ---
                    # Use regular assignment sequence counter
                    next_seq_num = company_sequence_counters[company_code]
                    company_sequence_counters[company_code] += 1  # Increment for next use
                    seq_prefix = "н"
                    seq_suffix = str(next_seq_num)
                    file_extension = ".docx"  # Regular assignments are saved as DOCX

                    # --- Construct File Names ---
                    base_filename = f"{company_code} {driver_initials} від {date_str_ddmmyyyy} {seq_prefix}{seq_suffix}"
                    # Regular assignments have both DOCX and PDF versions in company folders
                    docx_filename = company_word_path / f"{base_filename}.docx"
                    pdf_filename = company_pdf_path / f"{base_filename}.pdf"

                    # --- Check for Existing File ---
                    if docx_filename.exists():
                        print(f"Skipping row {index + 2}: Output file '{docx_filename}' already exists.")
                        continue

                    # --- Open Template and Replace Placeholders ---
                    word_doc = None # Initialize doc variable for this iteration
                    try:
                        print(f"Opening template: {template_path}")
                        word_doc = word_app.Documents.Open(str(template_path)) # Path object needs conversion to string
                        print(f"Replacing placeholders for Seq: {seq_suffix}, Driver: {driver_initials}")

                        # Regular assignment template uses standard placeholders
                        # Replace "Sequence"
                        find_replace(word_doc, "Sequence", seq_suffix)

                        # Replace "shortdrname" - use full name by default, initials if --short-names is specified
                        driver_name_for_signature = driver_initials if args.short_names else driver_name
                        find_replace(word_doc, "shortdrname", driver_name_for_signature)

                        # Replace "Date" with Ukrainian month format
                        date_ukrainian = format_date_ukrainian(date_str_ddmmyyyy)
                        find_replace(word_doc, "Date", date_ukrainian)

                        # Replace other placeholders from COLUMN_MAP
                        for placeholder, data_key in PLACEHOLDERS.items():
                            replace_value = data.get(data_key, '') # Get data, default to empty string if missing
                            if replace_value:  # Only replace if we have data
                                find_replace(word_doc, placeholder, replace_value)
                            else:
                                print(f"Warning: No data for placeholder '{placeholder}' (key: {data_key})")

                        # --- Save Documents ---
                        # For regular assignments, save as DOCX first, then PDF
                        print(f"Saving DOCX: {docx_filename}")
                        word_doc.SaveAs2(str(docx_filename))

                        print(f"Saving PDF: {pdf_filename}")
                        word_doc.SaveAs2(str(pdf_filename), FileFormat=WD_FORMAT_PDF)
                        
                        # Small delay to ensure PDF is completely written
                        time.sleep(0.5)
                        
                        # Remove blank pages from the PDF
                        if remove_blank_pages:
                            print(f"Checking for blank pages in assignment PDF...")
                            remove_blank_pages_from_pdf(str(pdf_filename))
                        else:
                            print(f"Blank page removal is disabled")
                        
                        # Track for reporting
                        report_lines.append(f"Assignment: {docx_filename.name}")
                        
                        # --- Write sequence number with hyperlink back to Excel ---
                        if excel_writer:
                            try:
                                excel_row_num = index + 2
                                success = excel_writer.add_sequence_number(
                                    row_num=excel_row_num,
                                    sequence_num=f"{seq_prefix}{seq_suffix}",
                                    hyperlink_path=str(docx_filename)
                                )
                                if success:
                                    print(f"Successfully wrote sequence number {seq_prefix}{seq_suffix} with hyperlink to Excel row {excel_row_num}")
                                else:
                                    print(f"Warning: Failed to write sequence number to Excel for row {excel_row_num}")
                            except Exception as e_excel_write:
                                print(f"Warning: Failed to write back to Excel for row {index + 2}: {e_excel_write}")
                        else:
                            if not excel_app:
                                print("Excel COM not available - skipping write-back")
                            elif not excel_assignments_col_idx:
                                print(f"Column '{COLUMN_MAP.get('assignments_col', 'накази')}' not configured or found - skipping write-back")

                        # --- Update File Counts ---
                        company_file_counts[company_code] = company_file_counts.get(company_code, 0) + 1
                        print(f"Successfully created files for row {index + 2}.")

                        # --- Close Document ---
                        word_doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                        word_doc = None # Release COM object for the document

                    except Exception as e_doc:
                        print(f"ERROR processing row {index + 2}: {e_doc}")
                        if word_doc:
                            try:
                                word_doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
                            except Exception as e_close:
                                print(f"Warning: Could not close document after error: {e_close}")
                            word_doc = None # Ensure release even on error

                    finally:
                        # Small delay can sometimes help COM stability with rapid file operations
                        time.sleep(0.1)

        # --- End of Loop ---

        # --- Build Report ---
        if company_file_counts:
            report_lines.append("\n--- Report ---")
            
            # Add consolidation statistics if available
            if consolidation_stats:
                report_lines.append("\n=== CONSOLIDATION MODE SUMMARY ===")
                report_lines.append(f"Found {consolidation_stats['unique_combinations']} unique date/company combinations")
                report_lines.append(f"Individual assignments: {consolidation_stats['individual']}")
                report_lines.append(f"Consolidated documents: {consolidation_stats['consolidated']}")
                report_lines.append(f"Groups with 2+ drivers: {consolidation_stats['multi_driver_groups']}")
                report_lines.append(f"PAPER SAVINGS: {consolidation_stats['papers_saved']} fewer documents ({consolidation_stats['papers_saved']} papers saved)")
                report_lines.append("")
            
            for company, count in company_file_counts.items():
                report_lines.append(f"{company}: {count} file(s) created.")
        else:
            report_lines.append("\nNo files were created.")

    except FileNotFoundError:
        print(f"Fatal Error: Excel file not found at '{excel_file_path}'")
    except ValueError as ve:
        print(f"Fatal Error: Configuration or Data Error - {ve}")
    except ImportError:
         print("Fatal Error: Required libraries (pandas, openpyxl, pywin32) not found.")
         print("Please install them: pip install pandas openpyxl pywin32")
    except Exception as e_main:
        print(f"\n--- An Unexpected Error Occurred ---")
        import traceback
        print(traceback.format_exc()) # Print detailed traceback
        report_lines.append(f"\nScript terminated early due to error: {e_main}")

    finally:
        # --- Quit Word Application ---
        if word_app:
            try:
                print("\nQuitting Word Application...")
                word_app.Quit()
                # Optional: Explicitly release COM object
                # import pythoncom
                # pythoncom.CoUninitialize() # Or del word_app
                del word_app # Helps Python's garbage collector release the COM object
            except Exception as e_quit:
                print(f"Warning: Error while quitting Word: {e_quit}")

        # --- Quit Excel Application ---
        if excel_app:
            try:
                print("Quitting Excel Application...")
                excel_app.DisplayAlerts = True  # Re-enable alerts
                excel_app.Quit()
                del excel_app
            except Exception as e_excel_quit:
                print(f"Warning: Error while quitting Excel: {e_excel_quit}")

        # --- Display Report ---
        print("\n-------------------------------------")
        print("Script execution finished.")
        for line in report_lines:
            print(line)
        print("-------------------------------------")

        # --- Play completion sound ---
        try:
            winsound.PlaySound(COMPLETE_SOUND, winsound.SND_FILENAME)
        except Exception as e_sound:
            print(f"Warning: Could not play completion sound: {e_sound}")