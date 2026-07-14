from pathlib import Path

import pandas as pd


input_path = Path(
    "extracted_output/analytical_observations.csv"
)
output_dir = Path("extracted_output/tables")
output_dir.mkdir(parents=True, exist_ok=True)

observations = pd.read_csv(input_path)

for block_id, table in observations.groupby("block_id"):
    safe_name = (
        str(block_id)
        .replace("/", "_")
        .replace("\\", "_")
        .replace(":", "_")
    )

    table.to_csv(
        output_dir / f"{safe_name}.csv",
        index=False,
    )


----------------------------

from __future__ import annotations

import html
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries


OUTPUT_DIR = Path("extracted_output")
RAW_CELLS_PATH = OUTPUT_DIR / "raw_cells.csv"
BLOCKS_PATH = OUTPUT_DIR / "detected_blocks.csv"
HTML_OUTPUT_PATH = OUTPUT_DIR / "detected_blocks_preview.html"


def first_available_value(row: pd.Series) -> Any:
    """Return the best displayable representation of a source cell."""

    cached_value = row.get("cached_value")
    if pd.notna(cached_value):
        return cached_value

    resolved_merged_value = row.get("resolved_merged_value")
    if pd.notna(resolved_merged_value):
        return resolved_merged_value

    raw_value = row.get("raw_value")
    if pd.notna(raw_value):
        return raw_value

    formula = row.get("formula")
    if pd.notna(formula):
        return formula

    return None


def reconstruct_block(
    raw_cells: pd.DataFrame,
    sheet_name: str,
    source_range: str,
) -> pd.DataFrame:
    """Reconstruct one detected block as an Excel-shaped DataFrame."""

    min_col, min_row, max_col, max_row = range_boundaries(source_range)

    block_cells = raw_cells[
        (raw_cells["sheet_name"] == sheet_name)
        & raw_cells["row_index"].between(min_row, max_row)
        & raw_cells["column_index"].between(min_col, max_col)
    ].copy()

    block_cells["preview_value"] = block_cells.apply(
        first_available_value,
        axis=1,
    )

    matrix = block_cells.pivot_table(
        index="row_index",
        columns="column_index",
        values="preview_value",
        aggfunc="first",
        dropna=False,
    )

    matrix = matrix.reindex(
        index=range(min_row, max_row + 1),
        columns=range(min_col, max_col + 1),
    )

    matrix.columns = [
        get_column_letter(column_index)
        for column_index in matrix.columns
    ]
    matrix.index.name = "Excel row"

    return matrix


def build_html_report(
    raw_cells: pd.DataFrame,
    detected_blocks: pd.DataFrame,
) -> str:
    """Build one HTML report containing all detected blocks."""

    sections: list[str] = []

    for _, block in detected_blocks.iterrows():
        block_id = str(block["block_id"])
        sheet_name = str(block["sheet_name"])
        source_range = str(block["source_range"])

        matrix = reconstruct_block(
            raw_cells=raw_cells,
            sheet_name=sheet_name,
            source_range=source_range,
        )

        block_class = block.get("block_class", "")
        status = block.get("interpretation_status", "")
        confidence = block.get("detection_confidence", "")

        sections.append(
            f"""
            <section class="block">
                <h2>{html.escape(block_id)}</h2>

                <div class="metadata">
                    <strong>Sheet:</strong> {html.escape(sheet_name)}
                    &nbsp;|&nbsp;
                    <strong>Range:</strong> {html.escape(source_range)}
                    &nbsp;|&nbsp;
                    <strong>Class:</strong> {html.escape(str(block_class))}
                    &nbsp;|&nbsp;
                    <strong>Status:</strong> {html.escape(str(status))}
                    &nbsp;|&nbsp;
                    <strong>Confidence:</strong>
                    {html.escape(str(confidence))}
                </div>

                <div class="table-container">
                    {matrix.to_html(
                        border=0,
                        na_rep="",
                        escape=True,
                    )}
                </div>
            </section>
            """
        )

    return f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <title>Detected Excel Blocks</title>

        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 24px;
                background: #f5f6f7;
            }}

            h1 {{
                margin-bottom: 30px;
            }}

            .block {{
                margin-bottom: 40px;
                padding: 20px;
                background: white;
                border: 1px solid #d9d9d9;
                border-radius: 8px;
            }}

            .metadata {{
                margin-bottom: 16px;
                color: #444;
            }}

            .table-container {{
                overflow-x: auto;
            }}

            table {{
                border-collapse: collapse;
                white-space: nowrap;
            }}

            th,
            td {{
                border: 1px solid #cccccc;
                padding: 6px 10px;
                text-align: left;
            }}

            th {{
                background: #eeeeee;
                position: sticky;
                top: 0;
            }}

            td:empty {{
                background: #fafafa;
            }}
        </style>
    </head>

    <body>
        <h1>Detected Excel Blocks</h1>
        {''.join(sections)}
    </body>
    </html>
    """


def main() -> None:
    raw_cells = pd.read_csv(
        RAW_CELLS_PATH,
        low_memory=False,
    )
    detected_blocks = pd.read_csv(
        BLOCKS_PATH,
        low_memory=False,
    )

    report = build_html_report(
        raw_cells=raw_cells,
        detected_blocks=detected_blocks,
    )

    HTML_OUTPUT_PATH.write_text(
        report,
        encoding="utf-8",
    )

    print(f"Preview created: {HTML_OUTPUT_PATH.resolve()}")


if __name__ == "__main__":
    main()
