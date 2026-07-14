"""End-to-end workbook extraction pipeline."""

from __future__ import annotations

import uuid
from pathlib import Path
from typing import Any

import pandas as pd

from .config import ExtractionConfig
from .interpret import block_record, interpret_block
from .models import ExtractionResult, ExtractionWarning
from .normalize import (
    block_column_records,
    block_row_records,
    normalize_block,
)
from .regions import detect_candidate_blocks
from .validate import (
    validate_coverage,
    validate_observations,
    validate_required_sheets,
    validate_simple_sum_formulas,
)
from .workbook import (
    anchor_warnings,
    extract_raw_cells,
    inspect_sheet,
    load_workbook_pair,
    workbook_inventory,
)


OUTPUT_COLUMNS: dict[str, list[str]] = {
    "workbook_inventory": ["workbook_id", "file_name"],
    "sheet_inventory": ["workbook_id", "sheet_name"],
    "raw_cells": ["workbook_id", "sheet_name", "cell_address"],
    "detected_blocks": ["workbook_id", "block_id", "sheet_name"],
    "block_columns": ["workbook_id", "block_id", "column_index"],
    "block_rows": ["workbook_id", "block_id", "row_index"],
    "analytical_observations": [
        "workbook_id",
        "observation_id",
        "block_id",
    ],
    "extraction_warnings": ["warning_code", "severity", "message"],
}


def _frame(records: list[dict[str, Any]], name: str) -> pd.DataFrame:
    """Create a stable DataFrame even when no records were produced."""

    if records:
        return pd.DataFrame.from_records(records)
    return pd.DataFrame(columns=OUTPUT_COLUMNS[name])


class ExcelExtractor:
    """Extract auditable raw and analytical tables from one workbook."""

    def __init__(self, config: ExtractionConfig | None = None) -> None:
        self.config = config or ExtractionConfig()

    def extract(self, path: str | Path) -> ExtractionResult:
        """Run the full extraction pipeline.

        :param path: Path to the source workbook.
        :return: Raw, interpreted, normalized, and warning tables.
        """

        workbook_path = Path(path).expanduser().resolve()
        if not workbook_path.exists():
            raise FileNotFoundError(workbook_path)
        workbook_id = str(uuid.uuid5(uuid.NAMESPACE_URL, str(workbook_path)))

        formula_workbook, value_workbook = load_workbook_pair(workbook_path)
        workbook_record = workbook_inventory(
            workbook_path, formula_workbook, workbook_id
        )

        sheet_records: list[dict[str, Any]] = []
        raw_cell_records: list[dict[str, Any]] = []
        block_records: list[dict[str, Any]] = []
        block_column_record_list: list[dict[str, Any]] = []
        block_row_record_list: list[dict[str, Any]] = []
        observation_records: list[dict[str, Any]] = []
        warnings: list[ExtractionWarning] = []

        warnings.extend(
            validate_required_sheets(
                formula_workbook.sheetnames,
                self.config.template.required_sheets,
                workbook_id,
            )
        )

        for sheet_name in formula_workbook.sheetnames:
            formula_sheet = formula_workbook[sheet_name]
            value_sheet = value_workbook[sheet_name]
            if (
                formula_sheet.sheet_state != "visible"
                and not self.config.include_hidden_sheets
            ):
                continue

            raw_cells, cell_warnings = extract_raw_cells(
                formula_sheet,
                value_sheet,
                workbook_id,
                self.config,
            )
            raw_cell_records.extend(raw_cells)
            warnings.extend(cell_warnings)
            sheet_records.append(
                inspect_sheet(formula_sheet, raw_cells, workbook_id)
            )

            sheet_rule = self.config.template.sheet_rules.get(sheet_name)
            if sheet_rule:
                warnings.extend(
                    anchor_warnings(
                        formula_sheet,
                        raw_cells,
                        workbook_id,
                        sheet_rule.expected_anchor_labels,
                    )
                )

            candidates = detect_candidate_blocks(
                formula_sheet,
                raw_cells,
                self.config,
                sheet_rule,
            )
            for candidate in candidates:
                interpretation = interpret_block(
                    candidate,
                    raw_cells,
                    self.config,
                    workbook_id,
                    sheet_rule,
                )
                block_records.append(
                    block_record(interpretation, workbook_id)
                )
                block_column_record_list.extend(
                    block_column_records(interpretation, workbook_id)
                )
                block_row_record_list.extend(
                    block_row_records(
                        interpretation,
                        raw_cells,
                        self.config,
                        workbook_id,
                    )
                )
                warnings.extend(interpretation.warnings)

                observations, normalization_warnings = normalize_block(
                    interpretation,
                    raw_cells,
                    self.config,
                    workbook_id,
                )
                observation_records.extend(observations)
                warnings.extend(normalization_warnings)

        warnings.extend(
            validate_observations(observation_records, workbook_id)
        )
        if self.config.evaluate_simple_sum_formulas:
            warnings.extend(
                validate_simple_sum_formulas(raw_cell_records, workbook_id)
            )
        coverage_metadata, coverage_warnings = validate_coverage(
            raw_cell_records,
            block_records,
            workbook_id,
        )
        warnings.extend(coverage_warnings)

        warning_records = [warning.to_record() for warning in warnings]
        tables = {
            "workbook_inventory": _frame(
                [workbook_record], "workbook_inventory"
            ),
            "sheet_inventory": _frame(
                sheet_records, "sheet_inventory"
            ),
            "raw_cells": _frame(raw_cell_records, "raw_cells"),
            "detected_blocks": _frame(
                block_records, "detected_blocks"
            ),
            "block_columns": _frame(
                block_column_record_list, "block_columns"
            ),
            "block_rows": _frame(block_row_record_list, "block_rows"),
            "analytical_observations": _frame(
                observation_records, "analytical_observations"
            ),
            "extraction_warnings": _frame(
                warning_records, "extraction_warnings"
            ),
        }
        metadata = {
            "workbook_id": workbook_id,
            "source_path": str(workbook_path),
            "config": self.config.to_dict(),
            "coverage": coverage_metadata,
            "table_row_counts": {
                name: len(frame) for name, frame in tables.items()
            },
        }
        return ExtractionResult(tables=tables, metadata=metadata)


def extract_workbook(
    path: str | Path,
    config: ExtractionConfig | None = None,
) -> ExtractionResult:
    """Convenience wrapper around :class:`ExcelExtractor`."""

    return ExcelExtractor(config=config).extract(path)
