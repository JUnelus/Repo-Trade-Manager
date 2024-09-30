# Repo-Trade-Manager

Repo Trade Manager is a Python-based tool designed to automate the process of managing repo trade data for a trading desk. The project integrates Python, Excel (with VBA macros), and SQLite to provide an automated solution for loading trade data, running macros to refresh Excel data, and saving the results in a SQLite database.

## Features

- **Dynamic Excel File Loading**: Automatically detects and loads the most recent Excel file ending with `trade_data.xlsm`.
- **VBA Macro Automation**: Automates the execution of Excel VBA macros to refresh data in the workbook.
- **SQLite Database Integration**: Loads repo trade data from Excel and saves it into a SQLite database for future reference and analysis.
- **Error Handling**: Ensures smooth operation, handling issues like missing files or database schema mismatches.
- **Easily Extensible**: Modify to support other types of financial trades or data pipelines.

## Technologies Used

- **Python**: Main programming language.
- **VBA**: Used to automate actions inside Excel (e.g., refreshing data).
- **SQLite**: A lightweight SQL database used to store repo trade data.
- **Excel**: Acts as the user interface for inputting and viewing trade data.

## Prerequisites

To run the project, you need the following installed:

- Python 3.8+
- pip: Python package manager to install dependencies.
- Excel (with macros enabled) on a Windows system.

## Installation

1. **Clone the Repository**:

    ```bash
    git clone https://github.com/junelus/repo-trade-manager.git
    cd repo-trade-manager
    ```

2. **Install Dependencies**: Use pip to install the necessary libraries:

    ```bash
    pip install -r requirements.txt
    ```

3. **Enable Macros in Excel**: Ensure macros are enabled in Excel by going to:

    ```
    File > Options > Trust Center > Trust Center Settings > Macro Settings
    ```

    Select `Enable all macros` and `Trust access to the VBA project object model`.

## Project Structure

```plaintext
repo-trade-manager/
├── README.md                 # Project documentation
├── requirements.txt          # Python dependencies
├── fetch_repo_trade_data.py  # Python script to fetch trade data
├── trade_manager.py          # Main Python script
├── db.sqlite3                # SQLite database (auto-created)
└── any file ending with 'trade_data.xlsm'  # Excel file(s)
```

# How It Works

The Repo Trade Manager script performs the following steps:

## Dynamic Excel File Detection
The script automatically finds the latest Excel file ending with `trade_data.xlsm` in the current directory.

## API Integration (Alpha Vantage)

This project uses the Alpha Vantage API to fetch financial data. To set up:

1. Get a free API key from [Alpha Vantage](https://www.alphavantage.co/support/#api-key).
2. Update `ALPHA_VANTAGE_API_KEY` in `.env` with your API key.
3. Run the Python script to pull real-time trade data into `trade_data.xlsx`:
   ```bash
   python fetch_repo_trade_data.py
    ```
   
## Loading Data from Excel
- The script reads trade data from the detected Excel file using `openpyxl`.
- It assumes the Excel sheet contains columns such as `Trade Date`, `Security`, `Quantity`, and `Rate`.
- The data is then loaded into a Pandas DataFrame for processing.

## Saving Data to SQLite
- After loading the data from Excel, it renames the columns to match the SQLite database schema and saves it into the `repo_trades` table in an SQLite database.

## Running the VBA Macro
- The script runs the `RefreshData` VBA macro inside the Excel file to refresh all data connections and updates the Excel workbook.

# Usage

## 1. Prepare Excel Data
Ensure that you have an Excel file ending with `trade_data.xlsm` in the project directory. This file should contain the trade data with the following columns:

| Trade Date  | Security  | Quantity | Rate |
| ----------- | --------- | -------- | ---- |
| YYYY-MM-DD  | Stock/Asset | Integer | Float|

## 2. Run the Script
To run the project, execute the following command in your terminal:

```bash
python trade_manager.py
```

This will:
- Load the most recent `trade_data.xlsm` file.
- Save the trade data into the SQLite database.
- Run the `RefreshData` macro inside the Excel workbook.

## 3. Check Data in SQLite
To check whether the data has been successfully loaded into the SQLite database, run the `check_data_in_sqlite()` function within `trade_manager.py`:

```python
import sqlite3

def check_data_in_sqlite():
    """Check if data is loaded into SQLite."""
    conn = sqlite3.connect('db.sqlite3')
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM repo_trades")
    rows = cursor.fetchall()
    for row in rows:
        print(row)
    conn.close()
```

## 4. Customization
Feel free to modify the following:
- Excel columns and data structure.
- VBA macro functionality.
- Additional logic to enhance SQL queries or support other trade-related features.

# Example Data

```csv
Trade Date,Security,Quantity,Rate
2024-09-30,AAPL,100,227.50
2024-09-30,GOOGL,150,1350.75
```

# Contributing
Contributions are welcome! If you’d like to contribute, please fork the repository and use a feature branch. Pull requests are warmly welcomed.