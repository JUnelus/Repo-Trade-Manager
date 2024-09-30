import os
import requests
import openpyxl
import pandas as pd
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Alpha Vantage API configuration
ALPHA_VANTAGE_API_KEY = os.getenv('ALPHA_VANTAGE_API_KEY')
ALPHA_VANTAGE_BASE_URL = 'https://www.alphavantage.co/query'

stock_symbol = 'AAPL'
stock_name = 'apple'

def fetch_repo_trade_data(symbol, interval='1min'):
    """Fetch repo trade data from Alpha Vantage API."""
    url = ALPHA_VANTAGE_BASE_URL
    params = {
        'function': 'TIME_SERIES_INTRADAY',
        'symbol': symbol,
        'interval': interval,
        'apikey': ALPHA_VANTAGE_API_KEY
    }
    response = requests.get(url, params=params)
    data = response.json()
    time_series = data.get(f"Time Series ({interval})", {})
    records = []
    for time, value in time_series.items():
        record = {
            'trade_date': time,
            'security': symbol,
            'quantity': 1000,  # Static quantity for the example
            'rate': value.get('1. open')  # Using the 'open' price as rate
        }
        records.append(record)
    return pd.DataFrame(records)

def write_to_excel(data, file_path=f'{stock_name}_trade_data.xlsx'):
    """Write fetched data to Excel."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Repo Trades"

    # Write header
    headers = ['Trade Date', 'Security', 'Quantity', 'Rate']
    ws.append(headers)

    # Write data
    for row in data.itertuples(index=False):
        ws.append(list(row))

    wb.save(file_path)
    print(f"Data written to {file_path}")


def main():
    df = fetch_repo_trade_data(stock_symbol)
    write_to_excel(df)

    print("Repo trade data fetched and saved to Excel.")

if __name__ == "__main__":
    main()
