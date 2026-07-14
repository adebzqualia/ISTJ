"""Post-extraction validation checks."""

from __future__ import annotations

import json
import math
import re
from collections import Counter, defaultdict
from typing import Any

from openpyxl.utils import range_boundaries

from .models import ExtractionWarning, WarningSeverity
from .utils import is_missing, is_number, normalize_label


SIMPLE_SUM_PATTERN = re.compile(
    r"^=SUM\((?:(?:'(?P<quoted>[^']+)'|(?P<plain>[^!]+))!)?"
    r"(?P<range>\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)\)$",
    re.IGNORECASE,
)


def validate_required_sheets(
    actual_sheet_names: list[str],
    required_sheet_names: tuple[str, ...],
    workbook_id: str,
) -> list[ExtractionWarning]:
    """Validate required worksheet presence."""

    actual = {normalize_label(name) for name in actual_sheet_names}
    return [
        ExtractionWarning(
            warning_code="WB003",
            severity=WarningSeverity.BLOCKING,
            message=f"Required worksheet is missing: {name!r}.",
            workbook_id=workbook_id,
            suggested_action="Use the expected template or update configuration.",
        )
        for name in required_sheet_names
        if normalize_label(name) not in actual
    ]


def validate_observations(
    observations: list[dict[str, Any]],
    workbook_id: str,
) -> list[ExtractionWarning]:
    """Validate observation provenance and duplicate semantic keys."""

    warnings: list[ExtractionWarning] = []
    source_counts = Counter(
        (
            item.get("source_sheet"),
            item.get("source_value_cell"),
            item.get("block_id"),
        )
        for item in observations
    )
    for (sheet, cell, block_id), count in source_counts.items():
        if count > 1:
            warnings.append(
                ExtractionWarning(
                    warning_code="VA003",
                    severity=WarningSeverity.ERROR,
                    message=(
                        f"Source cell was assigned to {count} observations."
                    ),
                    workbook_id=workbook_id,
                    sheet_name=sheet,
                    block_id=block_id,
                    cell_address=cell,
                    suggested_action="Review overlapping block normalization.",
                )
            )

    semantic_counts: defaultdict[tuple[Any, ...], list[dict[str, Any]]] = (
        defaultdict(list)
    )
    for item in observations:
        key = (
            item.get("block_id"),
            normalize_label(item.get("kpi_name")),
            item.get("period_start"),
            item.get("period_end"),
            normalize_label(item.get("scenario")),
            item.get("dimensions_json"),
            item.get("row_type"),
        )
        semantic_counts[key].append(item)

    for items in semantic_counts.values():
        if len(items) <= 1:
            continue
        first = items[0]
        warnings.append(
            ExtractionWarning(
                warning_code="VA002",
                severity=WarningSeverity.WARNING,
                message=(
                    f"Duplicate analytical key detected across {len(items)} "
                    "observations."
                ),
                workbook_id=workbook_id,
                sheet_name=first.get("source_sheet"),
                block_id=first.get("block_id"),
                cell_address=first.get("source_value_cell"),
                suggested_action=(
                    "Confirm whether the duplicate represents a valid split "
                    "or a repeated source row."
                ),
            )
        )
    return warnings


def validate_simple_sum_formulas(
    raw_cells: list[dict[str, Any]],
    workbook_id: str,
    *,
    absolute_tolerance: float = 1e-8,
    relative_tolerance: float = 1e-6,
) -> list[ExtractionWarning]:
    """Reconcile direct ``SUM(A1:A5)`` formulas against cached values."""

    lookup = {
        (record["sheet_name"], record["row_index"], record["column_index"]): record
        for record in raw_cells
    }
    warnings: list[ExtractionWarning] = []
    for record in raw_cells:
        formula = record.get("formula")
        if not formula:
            continue
        match = SIMPLE_SUM_PATTERN.fullmatch(formula.replace(" ", ""))
        if not match:
            continue
        source_sheet = (
            match.group("quoted")
            or match.group("plain")
            or record["sheet_name"]
        ).strip()
        min_col, min_row, max_col, max_row = range_boundaries(
            match.group("range").replace("$", "")
        )
        values: list[float] = []
        unresolved = False
        for row in range(min_row, max_row + 1):
            for column in range(min_col, max_col + 1):
                source = lookup.get((source_sheet, row, column))
                if not source:
                    continue
                value = (
                    source.get("cached_value")
                    if source.get("formula") is not None
                    else source.get("raw_value")
                )
                if is_missing(value):
                    continue
                if not is_number(value):
                    unresolved = True
                    break
                values.append(float(value))
            if unresolved:
                break
        cached_value = record.get("cached_value")
        if unresolved or not is_number(cached_value):
            continue
        expected = sum(values)
        if not math.isclose(
            float(cached_value),
            expected,
            abs_tol=absolute_tolerance,
            rel_tol=relative_tolerance,
        ):
            warnings.append(
                ExtractionWarning(
                    warning_code="VA001",
                    severity=WarningSeverity.ERROR,
                    message=(
                        f"Cached SUM result {cached_value!r} does not reconcile "
                        f"to source cells ({expected!r})."
                    ),
                    workbook_id=workbook_id,
                    sheet_name=record["sheet_name"],
                    cell_address=record["cell_address"],
                    suggested_action=(
                        "Recalculate the workbook and review formula dependencies."
                    ),
                )
            )
    return warnings


def validate_coverage(
    raw_cells: list[dict[str, Any]],
    block_records: list[dict[str, Any]],
    workbook_id: str,
) -> tuple[dict[str, Any], list[ExtractionWarning]]:
    """Measure coverage of non-empty cells by detected block ranges."""

    by_sheet: dict[str, list[tuple[int, int, int, int]]] = defaultdict(list)
    for block in block_records:
        by_sheet[block["sheet_name"]].append(
            (
                block["min_row"],
                block["max_row"],
                block["min_column"],
                block["max_column"],
            )
        )

    non_empty = 0
    covered = 0
    uncovered_examples: list[str] = []
    for record in raw_cells:
        value = (
            record.get("formula")
            if record.get("formula") is not None
            else record.get("raw_value")
        )
        if is_missing(value) and is_missing(record.get("resolved_merged_value")):
            continue
        non_empty += 1
        row = record["row_index"]
        column = record["column_index"]
        is_covered = any(
            min_row <= row <= max_row and min_col <= column <= max_col
            for min_row, max_row, min_col, max_col in by_sheet.get(
                record["sheet_name"], []
            )
        )
        if is_covered:
            covered += 1
        elif len(uncovered_examples) < 20:
            uncovered_examples.append(
                f"{record['sheet_name']}!{record['cell_address']}"
            )

    ratio = covered / non_empty if non_empty else 1.0
    metadata = {
        "non_empty_source_cells": non_empty,
        "cells_covered_by_detected_blocks": covered,
        "cell_detection_coverage_ratio": ratio,
        "uncovered_cell_examples": uncovered_examples,
    }
    warnings: list[ExtractionWarning] = []
    if non_empty and ratio < 0.80:
        warnings.append(
            ExtractionWarning(
                warning_code="VA004",
                severity=WarningSeverity.WARNING,
                message=(
                    f"Only {ratio:.1%} of non-empty cells are covered by "
                    "detected blocks."
                ),
                workbook_id=workbook_id,
                suggested_action=(
                    "Review region thresholds and uncovered-cell examples."
                ),
            )
        )
    return metadata, warnings
