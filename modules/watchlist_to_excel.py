"""
Convert a plain-text stock watchlist to a styled Excel spreadsheet.

Output file: A2ZStocks.xlsx (fixed, in outputs/)
Sheets:
  Companies  — Symbol | Name | Exchange | Date Added
               - Sorted by ticker
               - Dark header, alternating blue/white rows
               - Incremental: re-runs only append symbols not already present
  Summary    — Letter | Count  (number of stocks per first letter)
"""

from datetime import date
from pathlib import Path
from typing import Optional

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from modules.parser import read_and_parse

_OUTPUT_FILENAME = "A2ZStocks.xlsx"
_HEADER_FILL = PatternFill("solid", fgColor="2C3E50")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_EVEN_FILL = PatternFill("solid", fgColor="DCE6F1")
_ODD_FILL = PatternFill("solid", fgColor="FFFFFF")
_COMPANIES_HEADERS = ["Symbol", "Name", "Exchange", "Date Added"]
_SUMMARY_HEADERS = ["Letter", "Count"]


def _style_header(ws, headers: list[str]) -> None:
    for col, title in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center")


def _auto_size(ws, num_cols: int) -> None:
    for col_idx in range(1, num_cols + 1):
        max_len = max(
            len(str(ws.cell(row=r, column=col_idx).value or ""))
            for r in range(1, ws.max_row + 1)
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 3


def _load_existing(output_path: Path) -> dict[str, dict]:
    """Read the Companies sheet and return {TICKER: {name, exchange, date_added}}."""
    existing: dict[str, dict] = {}
    if not output_path.exists():
        return existing

    wb = load_workbook(output_path)
    if "Companies" not in wb.sheetnames:
        return existing

    ws = wb["Companies"]
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]

    def col(name):
        return headers.index(name)

    for row in ws.iter_rows(min_row=2, values_only=True):
        ticker = row[col("Symbol")]
        if ticker:
            existing[ticker.upper()] = {
                "name": row[col("Name")],
                "exchange": row[col("Exchange")],
                "date_added": row[col("Date Added")],
            }
    return existing


def _write_companies(ws, records: list[tuple[str, dict]]) -> None:
    _style_header(ws, _COMPANIES_HEADERS)
    for row_idx, (ticker, data) in enumerate(records, start=2):
        fill = _EVEN_FILL if row_idx % 2 == 0 else _ODD_FILL
        for col_idx, value in enumerate(
            [ticker, data["name"], data["exchange"], data["date_added"]], start=1
        ):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.fill = fill
    _auto_size(ws, len(_COMPANIES_HEADERS))


def _write_summary(ws, records: list[tuple[str, dict]]) -> None:
    _style_header(ws, _SUMMARY_HEADERS)
    letter_counts: dict[str, int] = {}
    for ticker, _ in records:
        letter = ticker[0].upper()
        letter_counts[letter] = letter_counts.get(letter, 0) + 1

    for row_idx, letter in enumerate(sorted(letter_counts), start=2):
        fill = _EVEN_FILL if row_idx % 2 == 0 else _ODD_FILL
        ws.cell(row=row_idx, column=1, value=letter).fill = fill
        ws.cell(row=row_idx, column=2, value=letter_counts[letter]).fill = fill

    # Total row — bold, same dark style as header
    total_row = ws.max_row + 1
    total_font = Font(bold=True, color="FFFFFF")
    for col, value in enumerate(["Total", sum(letter_counts.values())], start=1):
        cell = ws.cell(row=total_row, column=col, value=value)
        cell.fill = _HEADER_FILL
        cell.font = total_font
        cell.alignment = Alignment(horizontal="center")

    _auto_size(ws, len(_SUMMARY_HEADERS))


def export_watchlist_to_excel(
    input_path: str | Path,
    output_path: Optional[str | Path] = None,
) -> Path:
    """
    Parse *input_path* and incrementally update *output_path* (A2ZStocks.xlsx).

    - Stocks already present (matched by Symbol) are skipped.
    - New stocks are appended with today's date.
    - Summary sheet is rebuilt on every run.
    """
    input_path = Path(input_path)

    if output_path is None:
        output_path = Path(__file__).parent.parent / "outputs" / _OUTPUT_FILENAME
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Load what's already in the file
    existing = _load_existing(output_path)
    before = len(existing)

    # Parse new companies and merge
    today = date.today().isoformat()
    new_companies, unparsed = read_and_parse(input_path)

    for company in new_companies:
        key = company.ticker.upper()
        if key not in existing:
            existing[key] = {
                "name": company.name,
                "exchange": company.exchange,
                "date_added": today,
            }

    added = len(existing) - before

    # Write unparsed lines to Unparsed.txt
    unparsed_path = output_path.parent / "Unparsed.txt"
    if unparsed:
        unparsed_path.write_text("\n".join(unparsed) + "\n", encoding="utf-8")
    elif unparsed_path.exists():
        unparsed_path.unlink()

    # Clear successfully parsed lines from the input file — keep only unparsed lines
    input_path.write_text(
        "\n".join(unparsed) + ("\n" if unparsed else ""), encoding="utf-8"
    )

    print(f"Input file   : {input_path}")
    print(f"Existing     : {before} stocks")
    print(f"New added    : {added} stocks")
    print(f"Total unique : {len(existing)} stocks")
    print(f"Unparsed     : {len(unparsed)} lines{f' → {unparsed_path}' if unparsed else ' (none)'}")
    print(f"Cleared      : parsed lines removed from {input_path.name}")

    # Sort by ticker
    records = sorted(existing.items(), key=lambda x: x[0])

    # Build workbook
    wb = Workbook()
    ws_companies = wb.active
    ws_companies.title = "Companies"
    _write_companies(ws_companies, records)

    ws_summary = wb.create_sheet("Summary")
    _write_summary(ws_summary, records)

    wb.save(output_path)
    print(f"Saved        : {output_path}")
    return output_path


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python -m modules.watchlist_to_excel <input-file> [output-file]")
        sys.exit(1)
    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    export_watchlist_to_excel(inp, out)
