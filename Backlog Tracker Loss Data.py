
from tkinter import filedialog
import os, sys
import pandas as pd
import datetime as dt
import pyodbc
#my Package
parent_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.append(parent_dir)
from PythonTools import CREDS

downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
loss_file_path = filedialog.askopenfilename(
    title="Select Tracker Loss Data Excel File",
    filetypes=[("Excel files", "*.xlsx")],
    initialdir=downloads_folder
)

def connect_db():
    # Create a connection to the SQL database
    dbconn_str = (
                r'DRIVER={ODBC Driver 18 for SQL Server};'
                fr'SERVER={CREDS['DB_IP']}\SQLEXPRESS;'
                r'DATABASE=NARENCO_O&M_AE;'
                fr'UID={CREDS['DB_UID']};'
                fr'PWD={CREDS['DB_PWD']};'
                r'Encrypt=no;'
            )
    
    dbconnection = pyodbc.connect(dbconn_str)
    c = dbconnection.cursor()
    return c, dbconnection

def input_backlog_data_to_SQL(sheet_name, df):
    c, dbconnection = connect_db()
    clean_sheet_name = sheet_name.replace(", LLC", "").replace(" Solar", "").strip()
    table_name = f"[{clean_sheet_name} Tracker Loss Data]"
    
    # Tracker columns are typically from index 1 up to the last 3 columns (which are usually totals/meter data)
    tracker_cols = df.columns[1:-3]

    try:
        # Check if table exists
        c.execute(f"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{clean_sheet_name} Tracker Loss Data'")
        if c.fetchone()[0] == 0:
            # Create table
            columns_def = ["[Timestamp] DATETIME"]
            for tracker_col in tracker_cols:
                columns_def.append(f"[{tracker_col}] FLOAT")
            
            create_table_sql = f"CREATE TABLE {table_name} ({', '.join(columns_def)})"
            c.execute(create_table_sql)
            dbconnection.commit()
            print(f"Created table: {table_name}")
        
        # Check for missing columns and add them if necessary
        c.execute(f"SELECT TOP 0 * FROM {table_name}")
        existing_columns = [column[0] for column in c.description]
        
        for tracker_col in tracker_cols:
            if tracker_col not in existing_columns:
                try:
                    alter_sql = f"ALTER TABLE {table_name} ADD [{tracker_col}] FLOAT"
                    c.execute(alter_sql)
                    dbconnection.commit()
                    print(f"Added column {tracker_col} to {table_name}")
                except pyodbc.Error:
                    pass

        # Get existing timestamps to avoid duplicates
        c.execute(f"SELECT [Timestamp] FROM {table_name}")
        existing_timestamps = set(row[0] for row in c.fetchall())

        rows_to_insert = []
        columns = ["[Timestamp]"] + [f"[{col}]" for col in tracker_cols]
        placeholders = ["?"] * len(columns)
        insert_sql = f"INSERT INTO {table_name} ({', '.join(columns)}) VALUES ({', '.join(placeholders)})"

        for index, row in df.iterrows():
            timestamp = row.iloc[0]
            if timestamp in existing_timestamps:
                continue
            
            values = [timestamp.to_pydatetime() if isinstance(timestamp, pd.Timestamp) else timestamp]
            for col in tracker_cols:
                val = row[col]
                if pd.isna(val) or val == '':
                    values.append(None)
                else:
                    try:
                        values.append(float(val))
                    except:
                        values.append(None)
            
            rows_to_insert.append(values)
            existing_timestamps.add(timestamp)

        if rows_to_insert:
            c.fast_executemany = True
            c.executemany(insert_sql, rows_to_insert)
            dbconnection.commit()
            print(f"Inserted {len(rows_to_insert)} rows for {sheet_name}")
        else:
            print(f"No new data for {sheet_name}")

    except pyodbc.Error as e:
        print(f"Database error for {sheet_name}: {e}")
    finally:
        dbconnection.close()

if loss_file_path:
    xls = pd.ExcelFile(loss_file_path)
    for sheet_name in xls.sheet_names:
        if sheet_name != "Sheet 1":
            print(f"Processing {sheet_name} | {dt.datetime.now().time()}")
            df = pd.read_excel(loss_file_path, sheet_name=sheet_name, skiprows=2)
            df = df.iloc[1:] # Skip units row
            
            def parse_date(x):
                if isinstance(x, (int, float)):
                    return pd.to_datetime(x, unit='D', origin='1899-12-30')
                return pd.to_datetime(x, errors='coerce')

            df.iloc[:, 0] = df.iloc[:, 0].apply(parse_date)
            df.dropna(subset=[df.columns[0]], inplace=True)
            input_backlog_data_to_SQL(sheet_name, df)
    print("Done.")
