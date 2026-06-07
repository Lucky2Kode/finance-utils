"""Export stock lists to Excel with one sheet per index."""

from pathlib import Path
from typing import Optional

import pandas as pd

from modules.sp500 import get_sp500
from modules.nasdaq import get_nasdaq


def export_to_excel(
    output_path: Optional[Path] = None,
    include_sp500: bool = True,
    include_nasdaq: bool = True,
) -> Path:
    """
    Fetch stock lists and write them to an Excel workbook.

    Args:
        output_path: Destination .xlsx file. Defaults to outputs/stocks.xlsx.
        include_sp500: Include S&P 500 sheet.
        include_nasdaq: Include NASDAQ sheet.

    Returns:
        Path to the written file.
    """
    if output_path is None:
        output_path = Path(__file__).parent.parent / "outputs" / "stocks.xlsx"

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        if include_sp500:
            print("Fetching S&P 500 data...")
            sp500_df = get_sp500()
            sp500_df.to_excel(writer, sheet_name="SP500", index=False)
            print(f"  Written {len(sp500_df)} rows to sheet 'SP500'")

        if include_nasdaq:
            print("Fetching NASDAQ data...")
            nasdaq_df = get_nasdaq()
            nasdaq_df.to_excel(writer, sheet_name="NASDAQ", index=False)
            print(f"  Written {len(nasdaq_df)} rows to sheet 'NASDAQ'")

    print(f"\nSaved: {output_path}")
    return output_path


if __name__ == "__main__":
    export_to_excel()
