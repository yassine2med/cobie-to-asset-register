"""Command line: python -m cobieregister convert FILE.xlsx [-o register.csv]"""

from __future__ import annotations

import argparse
import sys

from .convert import convert, to_csv


def main(argv: list[str] | None = None) -> int:
    """Return 0 on success, 1 when --strict finds warnings, 2 when the file cannot be read."""
    parser = argparse.ArgumentParser(prog="cobieregister", description="COBie to asset register")
    sub = parser.add_subparsers(dest="command", required=True)
    cv = sub.add_parser("convert", help="write an asset register CSV from a COBie workbook")
    cv.add_argument("file")
    cv.add_argument("-o", "--output", help="CSV path (default: print to screen)")
    cv.add_argument("--strict", action="store_true", help="exit 1 when any row has a warning")
    args = parser.parse_args(argv)
    try:
        res = convert(args.file)
    except FileNotFoundError:
        print(f"File not found: {args.file}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"Cannot read {args.file}: {exc}", file=sys.stderr)
        return 2
    text = to_csv(res)
    if args.output:
        with open(args.output, "w", encoding="utf-8", newline="") as fh:
            fh.write(text)
        print(f"{len(res.rows)} assets written to {args.output}")
    else:
        print(text, end="")
    for w in res.warnings:
        print(f"warning: {w}", file=sys.stderr)
    return 1 if (args.strict and res.warnings) else 0


if __name__ == "__main__":
    sys.exit(main())
