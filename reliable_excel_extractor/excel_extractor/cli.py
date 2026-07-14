"""Command-line interface for workbook extraction."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from .config import ExtractionConfig
from .pipeline import ExcelExtractor


def _parser() -> argparse.ArgumentParser:
    """Build the command-line argument parser."""

    parser = argparse.ArgumentParser(
        description=(
            "Inspect a complex Excel workbook and extract auditable raw and "
            "analytical tables."
        )
    )
    parser.add_argument("workbook", type=Path, help="Source Excel workbook")
    parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="Directory receiving extraction tables",
    )
    parser.add_argument(
        "--config",
        type=Path,
        help="Optional JSON extraction configuration",
    )
    parser.add_argument(
        "--format",
        choices=("csv", "jsonl"),
        default="csv",
        help="Output table format",
    )
    return parser


def main() -> None:
    """Run the command-line extractor."""

    args = _parser().parse_args()
    config = ExtractionConfig()
    if args.config:
        with args.config.open("r", encoding="utf-8") as stream:
            config = ExtractionConfig.from_dict(json.load(stream))

    result = ExcelExtractor(config).extract(args.workbook)
    output = result.to_directory(args.output, file_format=args.format)
    counts = result.metadata["table_row_counts"]
    print(f"Extraction written to: {output}")
    for name, count in counts.items():
        print(f"  {name}: {count}")


if __name__ == "__main__":
    main()
