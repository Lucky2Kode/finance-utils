# RUN Guide — finance-utils

Step-by-step instructions to run every feature with all available options and examples.

---

## Prerequisites

### 1. Navigate to the project folder

```bash
cd "/Users/swethatanguturi/Desktop/Amigos/Tech సాంకేతికత/Tech 3 సాంకేతికత/GitRepos/Finance/finance-utils"
```

### 2. Install dependencies (first time only)

```bash
pip install -r requirements.txt
```

---

## Feature 1 — Watchlist to Excel (`--watchlist`)

Reads a stock watchlist file (`.txt` or `.xlsx`), appends new symbols to
`outputs/A2ZStocks.xlsx`, and clears processed entries from the input file.

### How it works

| Input type | Parsed from | After run |
|---|---|---|
| `.txt` | One stock per line in supported formats | Parsed lines removed; unparsed lines kept |
| `.xlsx` | Symbol / Name / Exchange columns | Processed rows deleted from input sheet |

**Output files:**
- `outputs/A2ZStocks.xlsx` — master workbook (Companies + Summary sheets)
- `outputs/Unparsed.txt` — lines that could not be parsed (`.txt` input only)

---

### Option A — Default (uses `inputs/A2ZStocks.txt`)

```bash
python main.py --watchlist
```

**Example output:**
```
Input file   : inputs/A2ZStocks.txt (text)
Existing     : 607 stocks
New added    : 5 stocks
Total unique : 612 stocks
Unparsed     : 0 lines (none)
Cleared      : parsed lines removed from A2ZStocks.txt
Saved        : outputs/A2ZStocks.xlsx
```

---

### Option B — Custom `.txt` input file

```bash
python main.py --watchlist inputs/my-stocks.txt
```

**Supported line formats in `.txt`:**

```
Apple Inc. (NASDAQ: AAPL)          # Name (EXCHANGE: TICKER)
Palantir (PLTR)                    # Name (TICKER) — no exchange
(NYSEARCA: XLE)                    # (EXCHANGE: TICKER) — no name
NASDAQ: MRVL                       # EXCHANGE: TICKER — bare format
```

---

### Option C — Excel file as input

Pass an existing `.xlsx` file. Each row is read and appended to `A2ZStocks.xlsx`.
After processing, those rows are removed from the input file.

```bash
python main.py --watchlist inputs/my-stocks.xlsx
```

**Expected columns in the input Excel (any order, case-insensitive):**

| Symbol | Name | Exchange |
|---|---|---|
| AAPL | Apple | NASDAQ |
| TSLA | Tesla | NASDAQ |

**Example output:**
```
Input file   : inputs/my-stocks.xlsx (Excel)
Existing     : 607 stocks
New added    : 12 stocks
Total unique : 619 stocks
failed rows  : 0 (kept in input file)
Cleared      : processed rows removed from my-stocks.xlsx
Saved        : outputs/A2ZStocks.xlsx
```

---

### Option D — Custom output path

```bash
python main.py --watchlist inputs/A2ZStocks.txt --watchlist-output outputs/MyPortfolio.xlsx
```

---

### Re-running safely

Re-running the same file is always safe — duplicates are skipped by ticker symbol.

```bash
# Run once
python main.py --watchlist

# Run again — adds 0 new stocks, no duplicates
python main.py --watchlist
```

---

### Output Excel — Sheet: Companies

| Symbol | Name | Exchange | Date Added |
|---|---|---|---|
| AAPL | Apple | NASDAQ | 2026-06-06 |
| TSLA | Tesla | NASDAQ | 2026-06-06 |
| ... | | | |

- Sorted alphabetically by Symbol
- Dark header (bold white text)
- Alternating light-blue / white rows

### Output Excel — Sheet: Summary

| Letter | Count |
|---|---|
| A | 55 |
| B | 35 |
| ... | |
| Z | 6 |
| **Total** | **607** |

---

## Feature 2 — Fetch Live Market Data (`--fetch`)

Downloads live S&P 500 and NASDAQ stock lists and writes them to `outputs/stocks.xlsx`.

> Requires an internet connection.

---

### Option A — Fetch both S&P 500 and NASDAQ (default)

```bash
python main.py --fetch
```

**Example output:**
```
Fetching S&P 500 data...
  Written 503 rows to sheet 'SP500'
Fetching NASDAQ data...
  Written 3847 rows to sheet 'NASDAQ'
```

**Output:** `outputs/stocks.xlsx` with two sheets — `SP500` and `NASDAQ`

---

### Option B — Fetch S&P 500 only

```bash
python main.py --fetch --no-nasdaq
```

---

### Option C — Fetch NASDAQ only

```bash
python main.py --fetch --no-sp500
```

---

### Option D — Custom output path

```bash
python main.py --fetch --fetch-output outputs/live-data.xlsx
```

---

## Feature 3 — Run Both Features Together

```bash
python main.py --watchlist --fetch
```

Runs the watchlist conversion first, then fetches live data. Both output files are generated independently.

---

## Quick Reference

| Command | What it does |
|---|---|
| `python main.py --watchlist` | Process default `inputs/A2ZStocks.txt` |
| `python main.py --watchlist inputs/file.txt` | Process a specific `.txt` file |
| `python main.py --watchlist inputs/file.xlsx` | Process a specific `.xlsx` file |
| `python main.py --watchlist --watchlist-output outputs/X.xlsx` | Custom output path |
| `python main.py --fetch` | Fetch live SP500 + NASDAQ |
| `python main.py --fetch --no-sp500` | Fetch NASDAQ only |
| `python main.py --fetch --no-nasdaq` | Fetch SP500 only |
| `python main.py --fetch --fetch-output outputs/X.xlsx` | Custom fetch output |
| `python main.py --watchlist --fetch` | Run both features |
| `python main.py --help` | Show all options |

---

## File Structure Reference

```
finance-utils/
├── main.py                        # Entry point — run this
├── requirements.txt               # pip dependencies
├── inputs/
│   └── A2ZStocks.txt              # Default watchlist input (edit to add stocks)
├── outputs/
│   ├── A2ZStocks.xlsx             # Master watchlist Excel (auto-updated)
│   ├── stocks.xlsx                # Live fetched data (SP500 + NASDAQ)
│   └── Unparsed.txt               # Lines that could not be parsed (if any)
└── modules/
    ├── parser.py                  # Text + Excel input parsing
    ├── watchlist_to_excel.py      # Watchlist → Excel writer
    ├── sp500.py                   # S&P 500 live fetch
    ├── nasdaq.py                  # NASDAQ live fetch
    └── stocks_to_excel.py         # Live data → Excel export
```
