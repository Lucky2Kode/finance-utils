"""Entry point — run all finance utilities."""

import argparse
from pathlib import Path

from modules.stocks_to_excel import export_to_excel
from modules.watchlist_to_excel import export_watchlist_to_excel


def parse_args():
    parser = argparse.ArgumentParser(
        description="Finance Utilities",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Convert default watchlist (inputs/A2ZStocks.txt) to Excel
  python main.py --watchlist

  # Convert a specific watchlist file
  python main.py --watchlist inputs/input-05-18-2026.txt

  # Fetch live SP500 + NASDAQ data and export to Excel
  python main.py --fetch

  # Both
  python main.py --watchlist --fetch
""",
    )

    parser.add_argument(
        "--watchlist",
        type=Path,
        nargs="?",
        const=Path("inputs/A2ZStocks.txt"),
        default=None,
        metavar="FILE",
        help="Parse a plain-text watchlist file and write Symbol|Name|Exchange to Excel "
             "(default: inputs/A2ZStocks.txt)",
    )
    parser.add_argument(
        "--watchlist-output",
        type=Path,
        default=None,
        metavar="FILE",
        help="Output path for watchlist Excel (default: outputs/<input_stem>.xlsx)",
    )

    parser.add_argument(
        "--fetch",
        action="store_true",
        help="Fetch live SP500 and NASDAQ data and export to Excel",
    )
    parser.add_argument(
        "--fetch-output",
        type=Path,
        default=Path("outputs/stocks.xlsx"),
        metavar="FILE",
        help="Output path for fetched stocks Excel (default: outputs/stocks.xlsx)",
    )
    parser.add_argument("--no-sp500", action="store_true", help="Skip S&P 500 (with --fetch)")
    parser.add_argument("--no-nasdaq", action="store_true", help="Skip NASDAQ (with --fetch)")

    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()

    if not args.watchlist and not args.fetch:
        print("No action specified. Use --watchlist <file> and/or --fetch.")
        print("Run with --help for usage.")
    else:
        if args.watchlist:
            export_watchlist_to_excel(args.watchlist, args.watchlist_output)

        if args.fetch:
            export_to_excel(
                output_path=args.fetch_output,
                include_sp500=not args.no_sp500,
                include_nasdaq=not args.no_nasdaq,
            )
