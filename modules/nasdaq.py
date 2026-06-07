"""Fetch NASDAQ-listed stocks via yfinance / pandas_datareader."""

import pandas as pd
import requests


def get_nasdaq() -> pd.DataFrame:
    """Return NASDAQ-listed companies as a DataFrame (via NASDAQ FTP)."""
    url = "https://api.nasdaq.com/api/screener/stocks"
    params = {
        "tableonly": "true",
        "limit": 5000,
        "offset": 0,
        "exchange": "NASDAQ",
        "download": "true",
    }
    headers = {"User-Agent": "Mozilla/5.0"}
    response = requests.get(url, params=params, headers=headers, timeout=30)
    response.raise_for_status()
    data = response.json()
    rows = data["data"]["rows"]
    df = pd.DataFrame(rows)
    return df


if __name__ == "__main__":
    df = get_nasdaq()
    print(df.head())
    print(f"\nTotal NASDAQ companies: {len(df)}")
