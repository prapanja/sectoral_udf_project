Excel-Integrated Financial Data Retrieval using SQLite

ITUS Capital - Data Analytics Internship Project

1. Project Objective

This project connects a Microsoft Excel workbook directly to a local SQLite database (sectoral_ebitda_margins.db). It allows an analyst to use simple Excel formulas to retrieve specific financial data, such as year-on-year revenue growth (curr_ttm_ebitda_margins) for any sector on any date.

The system is built in Python and uses the xlwings library to create a high-speed "bridge" between Excel (the "front-end") and the SQLite database (the "back-end").

2. File Structure

Your project folder should contain the following files:

sectoral_ebitda_margins.db: The core SQLite database file containing all financial data.

sectoral_data_udf.py: The main Python "robot" script. It contains all the Excel functions, database logic, caching, and logging.

config.ini: The configuration file. It tells the Python script the name of the database (db_path) and the data table (table_name).

apply_index.py: A one-time utility script that applies the schema.sql index to the database. This makes queries much faster.

schema.sql: A text file containing the SQL command to create a performance index on the (sector, date) columns.

test_queries.py: A simple Python script to prove that the database connection and query logic are working perfectly, without needing Excel.

example.xlsx: The Excel workbook where you type the formulas.

README.md: This file.

query_log.txt: (Generated automatically) A log file that records every function call, its parameters, and how long it took to run.

3. Function Reference

This system provides four custom Excel formulas:

=get_sectoral_quarterly_data(sector, field, date)

Description: Retrieves a single data point.

Example: =get_sectoral_quarterly_data("IT", "curr_ttm_ebitda_margins", "2025-06-30")

=get_series(sector, field, start_date, end_date)

Description: Returns a 2-column table (Date, Value) of data for one sector between two dates.

Example: =get_series("Auto & Auto Components", "curr_ttm_ebitda_margins", "2009-03-31", "2010-03-31")

=get_quarterly_matrix(date, field)

Description: Returns a 3-column table (Sector, Date, Value) of all sectors for a single date.

Example: =get_quarterly_matrix("2025-06-30", "curr_ttm_ebitda_margins")

=get_all_revenue_growth(sector, field)

Description: Returns a 2-column table (Date, Value) of the entire history for one sector.

Example: =get_all_revenue_growth("name", "curr_ttm_ebitda_margins")

4. How to Set Up the Project

Install Python: Ensure Python is installed (from python.org) and the "Add Python to PATH" box was checked during installation.

Install xlwings: Open a Command Prompt (cmd) and run:

pip install xlwings

Optimize Database: In your Command Prompt, navigate to the project folder and run the one-time organizer script:

# (First, navigate to folder)

cd <project-folder-path>

# (Now, run the script)

python apply_index.py

Install Excel Add-in: In the same Command Prompt, install the xlwings "bridge" into Excel. This may require running the Command Prompt as an Administrator.

xlwings addin install

5. How to Use the Excel Functions

There are two ways to "turn on" the Python robot.

Method A: Manual Server (Most Reliable)

This method forces the connection and is excellent for testing.

Close Excel completely.

Open a Command Prompt (cmd).

Navigate to project folder:

cd <project-folder-path>

Manually turn on the "robot":

python sectoral_data_udf.py

Message: xlwings server running... The window will "freeze." This is GOOD and means the robot is on.

MINIMIZE this window (do not close it).

Now, open example.xlsx file. The formulas will now work.

When finished, close Excel and the Command Prompt window.

Method B: Automatic Startup (Advanced)

This method tells Excel to run the Python script automatically in the background. This requires security permissions.

Configure Excel: Open your example.xlsx file.

Click the xlwings tab in the ribbon.

Interpreter box: Paste the full path to Python (find with where python), e.g., <path-to-your-python.exe>

PYTHONPATH box: Type a single dot: .

UDF Modules box: Type the name of the script: sectoral_data_udf

Trust Excel: Go to File > Options > Trust Center > Trust Center Settings... > Macro Settings and check the box Trust access to the VBA project object model.

Save and Restart: Save and close example.xlsx. When you re-open it, the formulas should work automatically. This may require Administrator rights to function correctly.

6. How to Run the Python Test Script

To prove the Python logic works without needing Excel, you can run the test_queries.py script.

Open a Command Prompt.

Navigate to your project folder.

Run the test script:

python test_queries.py

You will see this output printed directly to the screen, proving the code works:

--- TEST 1: GET SINGLE DATA ---
Data for Auto & Auto Components on 2009-03-31:
0.09641

---

--- TEST 2: GET DATA SERIES ---
Data series for Auto & Auto Components between 2009-03-31 and 2010-03-31:
('Date', 'curr_ttm_ebitda_margins')
('2009-03-31', 0.09641)
('2009-06-30', 0.10053)
('2009-09-30', 0.10755)
('2009-12-31', 0.12526)
('2010-03-31', 0.14113)

---

--- TEST 3: CHECK LOG FILE ---
SUCCESS: Log file 'query_log.txt' was found!

---

**_ TEST COMPLETE _**
