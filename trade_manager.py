import os
import glob
import openpyxl
import win32com.client as win32
import sqlite3
import pandas as pd
from sqlalchemy import create_engine

# Initialize the database
conn = sqlite3.connect('db.sqlite3')
engine = create_engine('sqlite:///db.sqlite3')

def create_table():
    """Create table to store trade data."""
    conn.execute('''
    CREATE TABLE IF NOT EXISTS repo_trades (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        "Trade Date" TEXT,
        "Security" TEXT,
        "Quantity" INTEGER,
        "Rate" REAL
    );
    ''')
    conn.commit()

def load_excel_data(file_path):
    """Load trade data from Excel into pandas DataFrame."""
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    data = sheet.iter_rows(values_only=True)
    headers = next(data)
    df = pd.DataFrame(data, columns=headers)
    return df

def save_to_sql(df):
    """Save DataFrame to SQLite."""
    df.to_sql('repo_trades', engine, if_exists='append', index=False)


def get_latest_trade_data_file():
    """Find the latest Excel file that ends with 'trade_data.xlsm'."""
    files = glob.glob('*trade_data.xlsm')  # Search for all files ending with 'trade_data.xlsm'
    if files:
        # Sort by modified time and return the most recent file
        latest_file = max(files, key=os.path.getmtime)
        return latest_file
    else:
        raise FileNotFoundError("No files ending with 'trade_data.xlsm' were found in the current directory.")


def run_vba_macro():
    """Run VBA macro to refresh data in Excel."""
    xl = win32.Dispatch('Excel.Application')

    # Get the latest Excel file that ends with 'trade_data.xlsm'
    file_path = os.path.abspath(get_latest_trade_data_file())

    wb = xl.Workbooks.Open(file_path)
    xl.Application.Run("RefreshData")
    wb.Close(SaveChanges=1)
    xl.Quit()


def main():
    create_table()

    # Get the latest Excel file that ends with 'trade_data.xlsm'
    excel_file = get_latest_trade_data_file()

    # Load data from the Excel file
    df = load_excel_data(excel_file)

    # Rename columns to match the SQL table schema
    df.columns = ['trade_date', 'security', 'quantity', 'rate']

    # Save DataFrame to SQL
    save_to_sql(df)

    # Run VBA macro
    run_vba_macro()

    print(f"Trade data from {excel_file} has been updated.")

if __name__ == "__main__":
    main()
