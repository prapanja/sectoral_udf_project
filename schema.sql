-- schema.sql
-- Run this with sqlite3 CLI: sqlite3 sectoral_ebitda_margins.db < schema.sql
PRAGMA foreign_keys = OFF;

-- If accord_code exists, create index on accord_code,date; otherwise use sector
-- (SQLite doesn't support conditional DDL — run safe commands)
CREATE INDEX IF NOT EXISTS idx_table_date_on_sector
ON sectoral_ebitda_margins (sector, date);

-- Also create a generic date index
CREATE INDEX IF NOT EXISTS idx_table_date
ON sectoral_ebitda_margins (date);
