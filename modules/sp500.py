"""Fetch S&P 500 constituents from Wikipedia."""

import pandas as pd


def get_sp500() -> pd.DataFrame:
    """Return S&P 500 tickers and company info as a DataFrame."""
    url = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
    tables = pd.read_html(url)
    df = tables[0]
    df.columns = [c.strip() for c in df.columns]
    return df


if __name__ == "__main__":
    df = get_sp500()
    print(df.head())
    print(f"\nTotal S&P 500 companies: {len(df)}")
