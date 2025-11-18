<<<<<<< HEAD
CREATE INDEX IF NOT EXISTS idx_sector_date
ON sectoral_ebitda_margins (sector, date);
=======
-- schema.sql
PRAGMA foreign_keys = OFF;

CREATE INDEX IF NOT EXISTS idx_table_date_on_sector
ON sectoral_ebitda_margins (sector, date);

CREATE INDEX IF NOT EXISTS idx_table_date
ON sectoral_ebitda_margins (date);
>>>>>>> 40ec65d0811ae239e4578cacc8ca1fc0b06dc8ba
