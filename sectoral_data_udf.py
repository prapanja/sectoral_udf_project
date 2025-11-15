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
