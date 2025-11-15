import sqlite3
import configparser
import os
import sys

def get_config():
    """Reads configuration from config.ini"""
    config = configparser.ConfigParser()
    script_dir = os.path.dirname(__file__)
    config_path = os.path.join(script_dir, 'config.ini')
    
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"config.ini not found at {config_path}")
    
    config.read(config_path)
    return config['Database']

def apply_schema(db_path, table_name, schema_file_path):
    """
    Applies the SQL schema (index) to the existing database.
    """
    if not os.path.exists(db_path):
        print(f"Error: Database file not found at {db_path}")
        print("Please make sure 'sectoral_ebitda_margins.db' is in the same directory.")
        return

    if not os.path.exists(schema_file_path):
        print(f"Error: Schema file not found at {schema_file_path}")
        return

    try:
        with open(schema_file_path, 'r') as f:
            sql_script = f.read()
            
        # Replace placeholder table name in schema.sql with config table_name
        sql_script = sql_script.replace("sectoral_ebitda_margins", table_name)

        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        print(f"Connecting to {db_path}...")
        cursor.executescript(sql_script)
        conn.commit()
        
        print(f"Index 'idx_sector_date' created successfully on table '{table_name}'.")
        
    except sqlite3.Error as e:
        print(f"An error occurred while applying the index: {e}")
        print("The index might already exist, which is okay.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    finally:
        if conn:
            conn.close()

if __name__ == "__main__":
    try:
        config = get_config()
        DB_PATH = config.get('db_path')
        TABLE_NAME = config.get('table_name')
        SCHEMA_FILE = 'schema.sql'
        
        # Ensure paths are relative to this script
        script_dir = os.path.dirname(__file__)
        db_full_path = os.path.join(script_dir, DB_PATH)
        schema_full_path = os.path.join(script_dir, SCHEMA_FILE)

        apply_schema(db_full_path, TABLE_NAME, schema_full_path)
        
    except Exception as e:
        print(f"Failed to run: {e}")
        sys.exit(1)