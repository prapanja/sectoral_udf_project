<<<<<<< HEAD
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
=======
# sectoral_data_udf.py
import os
import sqlite3
import time
import logging
from logging.handlers import RotatingFileHandler
from functools import lru_cache
import configparser
from datetime import datetime
import xlwings as xw  # used only for Excel UDF exposure; functions callable from CLI too

# ---------------------------
# Config loader
# ---------------------------
def load_config():
    cfg = configparser.ConfigParser()
    # search locations: script dir, cwd
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [os.path.join(script_dir, 'config.ini'), os.path.join(os.getcwd(), 'config.ini')]
    found = False
    for c in candidates:
        if os.path.exists(c):
            cfg.read(c)
            found = True
            break
    if not found:
        raise FileNotFoundError("config.ini not found in script directory or current working directory.")
    db_path = cfg.get('database', 'sqlite_path', fallback='sectoral_ebitda_margins.db')
    table_name = cfg.get('database', 'table_name', fallback='sectoral_ebitda_margins')
    date_format = cfg.get('database', 'date_format', fallback='%Y-%m-%d')
    return db_path, table_name, date_format

DB_PATH, TABLE_NAME, DATE_FORMAT = load_config()

# ---------------------------
# Logging
# ---------------------------
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'query_log.txt')
logger = logging.getLogger('sectoral_udf')
logger.setLevel(logging.INFO)
handler = RotatingFileHandler(LOG_FILE, maxBytes=1_000_000, backupCount=3)
formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
handler.setFormatter(formatter)
if not logger.handlers:
    logger.addHandler(handler)

def log_call(func_name, params, start_time, status='SUCCESS', error_msg=None):
    duration_ms = (time.perf_counter() - start_time) * 1000
    msg = f"{func_name} | params={params} | duration_ms={duration_ms:.2f} | status={status}"
    if error_msg:
        msg += f" | error={error_msg}"
    logger.info(msg)

# ---------------------------
# Database helper
# ---------------------------
def get_connection():
    # ensure path is absolute if relative in config
    db_abs = DB_PATH if os.path.isabs(DB_PATH) else os.path.abspath(DB_PATH)
    if not os.path.exists(db_abs):
        raise FileNotFoundError(f"SQLite DB not found at {db_abs}")
    conn = sqlite3.connect(db_abs, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
    conn.row_factory = sqlite3.Row
    return conn

def _validate_date(date_str):
    try:
        # allow date strings like '2025-06-30'
        datetime.strptime(date_str, DATE_FORMAT)
        return True
    except Exception:
        return False

# ---------------------------
# CACHES
# ---------------------------
# small LRU caches for lookups
@lru_cache(maxsize=1024)
def _cached_get_value(sector, field, date_str):
    conn = get_connection()
    cur = conn.cursor()
    q = f"SELECT {field} FROM {TABLE_NAME} WHERE sector = ? AND date = ? LIMIT 1"
    cur.execute(q, (sector, date_str))
    row = cur.fetchone()
    conn.close()
    if row is None:
        return None
    return row[0]

# ---------------------------
# Utility to format spill outputs for Excel
# ---------------------------
def format_two_column_series(rows):
    # rows: list of (date, value)
    data = [["date", "value"]]
    for d, v in rows:
        data.append([d, v])
    return data

def format_three_column_matrix(rows):
    data = [["sector", "date", "value"]]
    for sector, date, value in rows:
        data.append([sector, date, value])
    return data

# ---------------------------
# UDF implementations
# ---------------------------

# 1. get_sectoral_quarterly_data(sector, field, date)
@xw.func
def get_sectoral_quarterly_data(sector: str, field: str, date: str):
    start = time.perf_counter()
    func_name = 'get_sectoral_quarterly_data'
    params = {'sector': sector, 'field': field, 'date': date}
    try:
        if not (sector and field and date):
            raise ValueError("One or more inputs are empty.")
        if not _validate_date(date):
            raise ValueError(f"date must be in {DATE_FORMAT} format.")
        # protect against injection: we parameterize values; field is a column name so validate allowed columns
        allowed_columns = _get_columns()
        if field not in allowed_columns:
            raise ValueError(f"Field '{field}' not in table columns: {allowed_columns}")
        value = _cached_get_value(sector, field, date)
        if value is None:
            raise ValueError("No data found for given (sector, date).")
        log_call(func_name, params, start)
        return value
    except Exception as e:
        log_call(func_name, params, start, status='FAILURE', error_msg=str(e))
        return f"#ERROR: {e}"

# 2. get_series(sector, field, start_date, end_date) -> two columns (date, value) spilled
@xw.func
@xw.arg('start_date', pd=None)  # ensures Excel range mapping works
def get_series(sector: str, field: str, start_date: str, end_date: str):
    start_t = time.perf_counter()
    func_name = 'get_series'
    params = {'sector': sector, 'field': field, 'start_date': start_date, 'end_date': end_date}
    try:
        if not (sector and field and start_date and end_date):
            raise ValueError("Missing parameter.")
        if not (_validate_date(start_date) and _validate_date(end_date)):
            raise ValueError(f"Dates must be in {DATE_FORMAT} format.")
        allowed_columns = _get_columns()
        if field not in allowed_columns:
            raise ValueError(f"Field '{field}' not in table columns.")
        conn = get_connection()
        cur = conn.cursor()
        q = f"SELECT date, {field} FROM {TABLE_NAME} WHERE sector = ? AND date BETWEEN ? AND ? ORDER BY date"
        cur.execute(q, (sector, start_date, end_date))
        rows = cur.fetchall()
        conn.close()
        if not rows:
            raise ValueError("No rows found for given range.")
        series = [(r['date'], r[field]) for r in rows]
        out = format_two_column_series(series)
        log_call(func_name, params, start_t)
        return out
    except Exception as e:
        log_call(func_name, params, start_t, status='FAILURE', error_msg=str(e))
        return [[f"#ERROR: {e}"]]

# 3. get_quarterly_matrix(date, field) -> all sectors for that date (sector, date, value)
@xw.func
def get_quarterly_matrix(date: str, field: str):
    start_t = time.perf_counter()
    func_name = 'get_quarterly_matrix'
    params = {'date': date, 'field': field}
    try:
        if not (date and field):
            raise ValueError("Missing parameter.")
        if not _validate_date(date):
            raise ValueError(f"Date must be in {DATE_FORMAT} format.")
        allowed_columns = _get_columns()
        if field not in allowed_columns:
            raise ValueError(f"Field '{field}' not a column.")
        conn = get_connection()
        cur = conn.cursor()
        q = f"SELECT sector, date, {field} FROM {TABLE_NAME} WHERE date = ? ORDER BY sector"
        cur.execute(q, (date,))
        rows = cur.fetchall()
        conn.close()
        if not rows:
            raise ValueError("No rows found for that date.")
        out = format_three_column_matrix([(r['sector'], r['date'], r[field]) for r in rows])
        log_call(func_name, params, start_t)
        return out
    except Exception as e:
        log_call(func_name, params, start_t, status='FAILURE', error_msg=str(e))
        return [[f"#ERROR: {e}"]]

# 4. get_all_revenue_growth(sector, field) -> all dates for a sector (date, value)
@xw.func
def get_all_revenue_growth(sector: str, field: str):
    start_t = time.perf_counter()
    func_name = 'get_all_revenue_growth'
    params = {'sector': sector, 'field': field}
    try:
        if not (sector and field):
            raise ValueError("Missing parameter.")
        allowed_columns = _get_columns()
        if field not in allowed_columns:
            raise ValueError(f"Field '{field}' not a column.")
        conn = get_connection()
        cur = conn.cursor()
        q = f"SELECT date, {field} FROM {TABLE_NAME} WHERE sector = ? ORDER BY date"
        cur.execute(q, (sector,))
        rows = cur.fetchall()
        conn.close()
        if not rows:
            raise ValueError("No rows found for that sector.")
        out = format_two_column_series([(r['date'], r[field]) for r in rows])
        log_call(func_name, params, start_t)
        return out
    except Exception as e:
        log_call(func_name, params, start_t, status='FAILURE', error_msg=str(e))
        return [[f"#ERROR: {e}"]]

# ---------------------------
# Helper to get columns from table (caches)
# ---------------------------
@lru_cache(maxsize=1)
def _get_columns():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({TABLE_NAME})")
    cols = [r[1] for r in cur.fetchall()]  # r[1] is column name
    conn.close()
    return cols

# ---------------------------
# CLI test runner (quick tests)
# ---------------------------
if __name__ == "__main__":
    print("Running quick CLI tests against DB:", DB_PATH)
    # Replace these with actual sample values from your DB
    sample_sector = "Capital Goods"
    sample_date = "2025-06-30"
    field = "curr_ttm_ebitda_margins"

    print("Columns in table:", _get_columns())

    # Test single lookup
    print("Single lookup:", get_sectoral_quarterly_data(sample_sector, field, sample_date))

    # Test matrix
    matrix = get_quarterly_matrix(sample_date, field)
    print("Matrix rows:", matrix[:5])

    # Test series
    series = get_series(sample_sector, field, "2022-03-31", "2025-09-30")
    print("Series sample:", series[:6])

    # Test all dates for a sector
    all_rg = get_all_revenue_growth(sample_sector, field)
    print("All RG sample:", all_rg[:6])
>>>>>>> 40ec65d0811ae239e4578cacc8ca1fc0b06dc8ba
