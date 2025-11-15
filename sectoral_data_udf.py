import xlwings as xw
import sqlite3
import configparser
import logging
from logging.handlers import RotatingFileHandler
import time
import os
import sys
import functools
from datetime import datetime

# --- Configuration & Logging Setup ---

# Add script directory to path to find modules (for xlwings UDF server)
sys.path.append(os.path.dirname(__file__))

# Setup logging
def setup_logging():
    """Configures a rotating file logger."""
    log_file = 'query_log.txt'
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    
    # Prevent duplicate handlers if script is reloaded
    if logger.hasHandlers():
        return logger

    # 1MB rotating log file, keep 1 backup
    handler = RotatingFileHandler(log_file, maxBytes=1_048_576, backupCount=1)
    
    formatter = logging.Formatter(
        '%(asctime)s | %(levelname)s | %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    
    return logger

logger = setup_logging()

# Load configuration
def get_config():
    """Finds and reads the config.ini file."""
    config = configparser.ConfigParser()
    
    # Try script directory, then current working directory
    script_dir = os.path.dirname(__file__)
    config_paths = [
        os.path.join(script_dir, 'config.ini'),
        'config.ini'
    ]
    
    found_config = config.read(config_paths)
    if not found_config:
        logger.error("config.ini file not found in search paths.")
        raise FileNotFoundError("config.ini file not found.")
        
    return config['Database']

try:
    config = get_config()
    # Build absolute path to DB from script location
    script_dir = os.path.dirname(__file__)
    DB_PATH = os.path.join(script_dir, config.get('db_path'))
    TABLE_NAME = config.get('table_name')
    
    # Field validation: Prevent SQL injection on column/table names
    # Only allow these fields to be queried.
    SAFE_FIELDS = ('curr_ttm_ebitda_margins', 'sector', 'date')
    
    # Basic validation for table name (allow alphanumeric and underscore)
    if not (TABLE_NAME and TABLE_NAME.replace('_', '').isalnum()):
        raise ValueError(f"Invalid table name in config: {TABLE_NAME}")

    logger.info(f"Config loaded. DB_PATH: {DB_PATH}, TABLE_NAME: {TABLE_NAME}")

except Exception as e:
    logger.error(f"Failed to load configuration: {e}")
    # This will cause UDFs to fail, which is intended.

# --- Database Connection ---

def get_db_connection():
    """Establishes a connection to the SQLite database."""
    if not os.path.exists(DB_PATH):
        logger.error(f"Database file not found at {DB_PATH}")
        raise FileNotFoundError(f"Database file not found: {DB_PATH}")
    
    conn = sqlite3.connect(DB_PATH)
    # Return rows as tuples, which is efficient
    conn.row_factory = sqlite3.Row
    return conn

# --- Performance & Logging Decorator ---

def log_and_time(func):
    """Decorator to log function execution time and errors."""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        func_name = func.__name__
        # Format args for logging, handle potential datetime objects
        log_args = []
        for arg in args:
            if isinstance(arg, datetime):
                log_args.append(arg.strftime('%Y-%m-%d'))
            else:
                log_args.append(str(arg))
        
        start_time = time.perf_counter()
        
        try:
            result = func(*args, **kwargs)
            status = "SUCCESS"
            return result
        
        except Exception as e:
            status = f"FAILURE: {e}"
            logger.error(f"Function: {func_name}, Args: {log_args}, Error: {e}")
            # Propagate error to Excel UDF
            return f"Error: {e}"
        
        finally:
            end_time = time.perf_counter()
            exec_time_ms = (end_time - start_time) * 1000
    
            log_message = (
                f"Function: {func_name} | "
                f"Args: {log_args} | "
                f"Status: {status} | "
                f"Time: {exec_time_ms:.2f} ms"
            )
            logger.info(log_message)
            
    return wrapper

# --- Internal Core Query Functions (Cached) ---

def _validate_field(field):
    """Check field against a safelist to prevent SQL injection."""
    if field not in SAFE_FIELDS:
        raise ValueError(f"Invalid field: '{field}'.")
    return field

def _format_date(date_obj):
    """Convert Excel datetime object to YYYY-MM-DD string."""
    if isinstance(date_obj, datetime):
        return date_obj.strftime('%Y-%m-%d')
    if isinstance(date_obj, str):
        # Add basic parsing to handle YYYY-MM-DD HH:MM:SS from DB
        if ' ' in date_obj:
            return date_obj.split(' ')[0]
        return date_obj
    # Handle Excel's number-based dates if xlwings doesn't convert them
    if isinstance(date_obj, (int, float)):
        # This conversion is tricky and platform-dependent.
        # xlwings *usually* handles this.
        logger.warning(f"Received numeric date {date_obj}, treating as string.")
        return str(date_obj) # Fallback, but might fail in SQL
    raise ValueError("Invalid date format. Expected YYYY-MM-DD or Excel date.")

@functools.lru_cache(maxsize=256) # Cache up to 256 recent queries
@log_and_time
def _query_single_data(sector, field, date_str):
    """Internal function to fetch a single data point."""
    safe_field = _validate_field(field)
    
    # Use date() function in SQL to ignore time part
    query = f"SELECT {safe_field} FROM {TABLE_NAME} WHERE sector = ? AND date(date) = ?"
    
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(query, (sector, date_str))
        row = cursor.fetchone()
        
    if row:
        return row[0]
    else:
        return "Error: No data found."

@functools.lru_cache(maxsize=128)
@log_and_time
def _query_series(sector, field, start_date_str, end_date_str):
    """Internal function to fetch a time series."""
    safe_field = _validate_field(field)
    
    query = f"SELECT date(date) as formatted_date, {safe_field} FROM {TABLE_NAME} WHERE sector = ? AND date(date) BETWEEN ? AND ? ORDER BY formatted_date"
    
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(query, (sector, start_date_str, end_date_str))
        # Convert from list of Row objects to list of tuples
        rows = [tuple(row) for row in cursor.fetchall()]
    
    # Add headers for the spilled array
    return [('Date', field)] + rows

@functools.lru_cache(maxsize=64)
@log_and_time
def _query_matrix(date_str, field):
    """Internal function to fetch all sectors for a given date."""
    safe_field = _validate_field(field)
    
    query = f"SELECT sector, date(date) as formatted_date, {safe_field} FROM {TABLE_NAME} WHERE date(date) = ? ORDER BY sector"
    
    with get_db_connection() as conn:
        cursor = conn.cursor()
        rows = [tuple(row) for row in cursor.fetchall()]
        
    return [('Sector', 'Date', field)] + rows

@functools.lru_cache(maxsize=128)
@log_and_time
def _query_all_growth(sector, field):
    """Internal function to fetch all data for a given sector."""
    safe_field = _validate_field(field)
    
    query = f"SELECT date(date) as formatted_date, {safe_field} FROM {TABLE_NAME} WHERE sector = ? ORDER BY formatted_date"
    
    with get_db_connection() as conn:
        cursor = conn.cursor()
        rows = [tuple(row) for row in cursor.fetchall()]
        
    return [('Date', field)] + rows

# --- Excel UDF Definitions ---

@xw.func
@xw.arg('date', datetime, doc="Date as YYYY-MM-DD or Excel date.")
def get_sectoral_quarterly_data(sector, field, date):
    """
    Retrieves a single data point for a given sector, field, and date.
    Example: =get_sectoral_quarterly_data("IT", "curr_ttm_ebitda_margins", "2025-06-30")
    """
    try:
        date_str = _format_date(date)
        return _query_single_data(sector, field, date_str)
    except Exception as e:
        logger.error(f"UDF get_sectoral_quarterly_data Error: {e}")
        return f"Error: {e}"

@xw.func
@xw.arg('start_date', datetime, doc="Start date (YYYY-MM-DD or Excel date).")
@xw.arg('end_date', datetime, doc="End date (YYYY-MM-DD or Excel date).")
@xw.ret(expand='table')
def get_series(sector, field, start_date, end_date):
    """
    Retrieves a time series for a sector between two dates.
    Example: =get_series("Capital Goods", "curr_ttm_ebitda_margins", "2022-03-31", "2025-09-30")
    """
    try:
        start_date_str = _format_date(start_date)
        end_date_str = _format_date(end_date)
        return _query_series(sector, field, start_date_str, end_date_str)
    except Exception as e:
        logger.error(f"UDF get_series Error: {e}")
        return [[f"Error: {e}"]] # Return as 2D array for spilling

@xw.func
@xw.arg('date', datetime, doc="Date as YYYY-MM-DD or Excel date.")
@xw.ret(expand='table')
def get_quarterly_matrix(date, field):
    """
    Retrieves data for all sectors on a specific date.
    Example: =get_quarterly_matrix("2025-06-30", "curr_ttm_ebitda_margins")
    """
    try:
        date_str = _format_date(date)
        return _query_matrix(date_str, field)
    except Exception as e:
        logger.error(f"UDF get_quarterly_matrix Error: {e}")
        return [[f"Error: {e}"]]

@xw.func
@xw.ret(expand='table')
def get_all_revenue_growth(sector, field):
    """
    Retrieves the entire history for a single sector.
    Example: =get_all_revenue_growth("Healthcare", "curr_ttm_ebitda_margins")
    """
    try:
        return _query_all_growth(sector, field)
    except Exception as e:
        logger.error(f"UDF get_all_revenue_growth Error: {e}")
        return [[f"Error: {e}"]]

# --- Main entry point for xlwings UDT run ---
if __name__ == "__main__":
    # This part is for running with 'xlwings udt run'
    # It tells xlwings which functions to expose.
    xw.serve()