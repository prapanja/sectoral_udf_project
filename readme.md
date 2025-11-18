
# Sectoral UDF Project

This project provides a complete Excel–Python integration system for retrieving sectoral data from a SQLite database.  
It uses **xlwings** to expose Python functions as Excel UDFs and provides clean, fast access to quarterly and time-series metrics.

---

##  Project Structure

```

sectoral_udf_project/
│
├── sectoral_data_udf.py          # Main Python UDF module
├── config.ini                     # Configuration (DB path, table name, date format)
├── schema.sql                     # Index creation for performance
├── sectoral_ebitda_margins.db     # SQLite database (project dataset)
├── README.md                      # Documentation
└── .venv/                         # Virtual environment (excluded from Git)

````

---

##  Installation & Setup

### 1. Create and activate the virtual environment

```bash
python -m venv .venv
.\.venv\Scripts\activate   # Windows
````

### 2. Install dependencies

```bash
pip install xlwings==0.30.12
```

### 3. Install the xlwings Excel Add-in

(If not already installed)

```bash
xlwings addin install
```

Or copy the downloaded `xlwings.xlam` manually into:

```
%appdata%\Microsoft\Excel\XLSTART
```

---

##  Configure the Database (config.ini)

config.ini must contain:

```ini
[database]
sqlite_path = sectoral_ebitda_margins.db
table_name = sectoral_ebitda_margins
date_format = %Y-%m-%d
```

---

##  Create Database Index (Performance)

Run this:

```bash
sqlite3 sectoral_ebitda_margins.db < schema.sql
```

`schema.sql`:

```sql
CREATE INDEX IF NOT EXISTS idx_sector_date
ON sectoral_ebitda_margins (sector, date);

CREATE INDEX IF NOT EXISTS idx_date
ON sectoral_ebitda_margins (date);
```

---

## Python UDFs Available in Excel

---

###  `get_sectoral_quarterly_data(sector, field, date)`

Returns a single quarterly value.

**Example Excel formula:**

```
=get_sectoral_quarterly_data("Capital Goods","curr_ttm_ebitda_margins","2025-06-30")
```

---

###  `get_series(sector, field, start_date, end_date)`

Returns a spill range:
`date | value`

**Example:**

```
=get_series("Capital Goods","curr_ttm_ebitda_margins","2022-03-31","2025-09-30")
```

---

###  `get_quarterly_matrix(date, field)`

Returns:
`sector | date | value`

**Example:**

```
=get_quarterly_matrix("2025-06-30","curr_ttm_ebitda_margins")
```

---

###  `get_all_revenue_growth(sector, field)`

All dates for that sector.

**Example:**

```
=get_all_revenue_growth("Capital Goods","curr_ttm_ebitda_margins")
```

---

##  Excel Setup Instructions 

Inside Excel → **xlwings tab**:

###  Interpreter (FULL PATH):

```
C:\Users\Desktop\sectoral_udf_project\.venv\Scripts\python.exe
```

###  UDF Modules:

```
C:\Users\Desktop\sectoral_udf_project\sectoral_data_udf.py
```

###  Then click:

```
Restart UDF Server
```

(Enable **Show Console** to see errors.)

---

##  Testing the Functions

After restarting the UDF server, test:

```
=get_sectoral_quarterly_data("Capital Goods","curr_ttm_ebitda_margins","2025-06-30")
```

---

##  GitHub Submission Steps

```bash
git init
git add .
git commit -m "Initial project commit"
git branch -M main
git remote add origin <your-repo-url>
git push -u origin main
```

---



