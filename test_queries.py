# This is a test script to prove the Python code works.
import sectoral_data_udf
from datetime import datetime
import os

# --- Test 1: Get Single Data Point ---
print("--- TEST 1: GET SINGLE DATA ---")
sector = "Auto & Auto Components"
field = "curr_ttm_ebitda_margins"

date_obj = datetime(2009, 3, 31)
date_str = sectoral_data_udf._format_date(date_obj)

try:
    data = sectoral_data_udf._query_single_data(sector, field, date_str)
    print(f"Data for {sector} on {date_str}:")
    print(data)
    print("---------------------------------")
except Exception as e:
    print(f"TEST 1 FAILED: {e}")
    print("---------------------------------")


# --- Test 2: Get Data Series ---
print("\n--- TEST 2: GET DATA SERIES ---")
start_obj = datetime(2009, 3, 31)
end_obj = datetime(2010, 3, 31)
start_str = sectoral_data_udf._format_date(start_obj)
end_str = sectoral_data_udf._format_date(end_obj)

try:
    series_data = sectoral_data_udf._query_series(sector, field, start_str, end_str)
    print(f"Data series for {sector} between {start_str} and {end_str}:")
    # Print the first 5 rows
    for row in series_data[:6]:
        print(row)
    print("---------------------------------")
except Exception as e:
    print(f"TEST 2 FAILED: {e}")
    print("---------------------------------")


# --- Test 3: Check Log File ---
print("\n--- TEST 3: CHECK LOG FILE ---")
log_file_path = os.path.join(os.path.dirname(__file__), 'query_log.txt')
if os.path.exists(log_file_path):
    print(f"SUCCESS: Log file 'query_log.txt' was found!")
else:
    print(f"FAILURE: Log file 'query_log.txt' was NOT found.")
print("---------------------------------")

print("\n*** TEST COMPLETE ***")
