# finance-utils

A collection of Python utilities for working with financial market data.

---

## Features

| Module | Description |
|---|---|
| `modules/text_parser.py` | Parse plain-text watchlist files into structured Company records |
| `modules/watchlist_to_excel.py` | Convert a watchlist text file to a styled Excel spreadsheet |
| `modules/sp500.py` | Fetch the S&P 500 constituent list from Wikipedia |
| `modules/nasdaq.py` | Fetch NASDAQ-listed stocks from the NASDAQ API |
| `modules/stocks_to_excel.py` | Export live SP500/NASDAQ data to a multi-sheet Excel workbook |

---

## Project Structure

```
finance-utils/
├── main.py                        # CLI entry point
├── requirements.txt
├── inputs/                        # Place watchlist .txt files here
│   └── input-05-18-2026.txt       # Sample watchlist
├── outputs/                       # Generated Excel files land here
└── modules/
    ├── __init__.py
    ├── text_parser.py             # Watchlist text file parser (Company model)
    ├── watchlist_to_excel.py      # Styled Excel writer (mirrors StocksToExcel)
    ├── sp500.py                   # S&P 500 live data
    ├── nasdaq.py                  # NASDAQ live data
    └── stocks_to_excel.py         # Live data → Excel export
```

---

## Setup

```bash
pip install -r requirements.txt
```

---

## Usage

### 1. Convert a watchlist text file to Excel (main feature)

Mirrors the Java **StocksToExcel** project. Reads a plain-text watchlist and
produces a styled `.xlsx` file.

```bash
python main.py --watchlist inputs/input-05-18-2026.txt
# Output: outputs/input-05-18-2026.xlsx
```

Custom output path:

```bash
python main.py --watchlist inputs/my-stocks.txt --watchlist-output outputs/my-stocks.xlsx
```

**Input file format** (two supported patterns):

```
Company Name (EXCHANGE: TICKER)
Company Name (TICKER)
```

Example:

```
Energy Transfer (NYSE: ET)
Intel (NASDAQ: INTC)
Analog Devices' (NASDAQ: ADI)
MongoDB's (NASDAQ: MDB)
Wendy's Co. (NASDAQ: WEN)
```

**Excel output:**

| Symbol | Name            | Exchange |
|--------|-----------------|----------|
| ADI    | Analog Devices' | NASDAQ   |
| ET     | Energy Transfer | NYSE     |
| INTC   | Intel           | NASDAQ   |
| MDB    | MongoDB's       | NASDAQ   |
| WEN    | Wendy's Co.     | NASDAQ   |

- Dark header (RGB 44, 62, 80) with bold white text
- Alternating light-blue / white rows
- Sorted by ticker symbol
- Duplicate entries removed (by ticker + exchange)
- Name cleaning: strips `Inc.` and `Corp` suffixes

---

### 2. Fetch live S&P 500 + NASDAQ data to Excel

```bash
python main.py --fetch
# Output: outputs/stocks.xlsx  (sheets: SP500, NASDAQ)
```

Options:

```
--fetch-output PATH   Custom output path (default: outputs/stocks.xlsx)
--no-sp500            Skip the S&P 500 sheet
--no-nasdaq           Skip the NASDAQ sheet
```

---

### 3. Both at once

```bash
python main.py --watchlist inputs/input-05-18-2026.txt --fetch
```

---

### Use modules directly in Python

```python
from modules.text_parser import read_and_parse
from modules.watchlist_to_excel import export_watchlist_to_excel
from modules.sp500 import get_sp500
from modules.nasdaq import get_nasdaq
from modules.stocks_to_excel import export_to_excel

# Parse a watchlist file
companies = read_and_parse("inputs/input-05-18-2026.txt")

# Export watchlist to styled Excel
export_watchlist_to_excel("inputs/input-05-18-2026.txt")

# Get live data as DataFrames
sp500_df = get_sp500()
nasdaq_df = get_nasdaq()

# Export live data to Excel
export_to_excel(output_path="outputs/stocks.xlsx")
```

---

## Planned Additions

- NYSE stock list
- Dow Jones 30 constituents
- Russell 2000 / Russell 1000
- Stock price history download (via `yfinance`)
- Portfolio tracker
- Dividend yield screener
