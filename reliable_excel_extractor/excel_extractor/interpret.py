"""Heuristic and deterministic interpretation of candidate blocks."""

from __future__ import annotations

import statistics
from collections import Counter
from typing import Any

from .config import ExtractionConfig, SheetRule
from .models import (
    BlockCandidate,
    BlockInterpretation,
    ExtractionWarning,
    WarningSeverity,
)
from .periods import parse_period_label
from .utils import is_missing, is_number, normalize_label, technical_name


KNOWN_RECORD_FIELDS = {
    "entity",
    "country",
    "region",
    "business",
    "business line",
    "product",
    "category",
    "period",
    "date",
    "year",
    "month",
    "quarter",
    "kpi",
    "kpi name",
    "metric",
    "measure",
    "value",
    "amount",
    "scenario",
    "currency",
    "unit",
}


def _cell_lookup(
    raw_cells: list[dict[str, Any]],
) -> dict[tuple[int, int], dict[str, Any]]:
    """Index raw-cell records by coordinate."""

    return {
        (record["row_index"], record["column_index"]): record
        for record in raw_cells
    }


def effective_value(record: dict[str, Any] | None) -> Any:
    """Return a value suitable for structural interpretation."""

    if not record:
        return None
    if not is_missing(record.get("resolved_merged_value")):
        return record["resolved_merged_value"]
    if record.get("formula") is not None:
        if not is_missing(record.get("cached_value")):
            return record["cached_value"]
        return record["formula"]
    return record.get("raw_value")


def _is_header_like(record: dict[str, Any] | None) -> bool:
    """Return whether a cell has common header characteristics."""

    value = effective_value(record)
    if is_missing(value):
        return False
    if isinstance(value, str):
        return True
    if isinstance(value, int) and 1900 <= value <= 2200:
        return True
    return bool(record and record.get("is_bold"))


def _is_numeric_like(record: dict[str, Any] | None) -> bool:
    """Return whether a cell resolves to a numeric measure."""

    return is_number(effective_value(record))


def _row_values(
    lookup: dict[tuple[int, int], dict[str, Any]],
    row: int,
    min_col: int,
    max_col: int,
) -> list[Any]:
    """Return effective values for one bounded row."""

    return [
        effective_value(lookup.get((row, column)))
        for column in range(min_col, max_col + 1)
    ]


def _title_row_score(values: list[Any]) -> float:
    """Score whether a row resembles a sparse title or section label."""

    present = [value for value in values if not is_missing(value)]
    if not present:
        return 0.7
    if len(present) <= 2 and all(isinstance(value, str) for value in present):
        return 1.0
    return 0.0


def _header_candidate_score(
    candidate: BlockCandidate,
    lookup: dict[tuple[int, int], dict[str, Any]],
    title_rows: int,
    header_depth: int,
) -> tuple[float, dict[int, list[Any]], list[str]]:
    """Score one title/header split and reconstruct header paths."""

    width = candidate.width
    header_start = candidate.min_row + title_rows
    header_end = header_start + header_depth - 1
    body_start = header_end + 1
    if header_depth == 0:
        header_end = header_start - 1
        body_start = header_start
    if body_start > candidate.max_row:
        return -1.0, {}, ["No body rows remain after the header."]

    reasons: list[str] = []
    title_score = 1.0
    if title_rows:
        title_scores: list[float] = []
        for row in range(candidate.min_row, header_start):
            records = [
                lookup.get((row, column))
                for column in range(candidate.min_col, candidate.max_col + 1)
            ]
            present = [
                record
                for record in records
                if not is_missing(effective_value(record))
            ]
            origins = {
                record.get("merged_anchor") or record.get("cell_address")
                for record in present
            }
            values = [effective_value(record) for record in present]
            if origins and len(origins) <= 2 and all(
                isinstance(value, str) for value in values
            ):
                title_scores.append(1.0)
            else:
                title_scores.append(_title_row_score(values))
        title_score = statistics.fmean(title_scores) if title_scores else 1.0

    header_records = [
        lookup.get((row, column))
        for row in range(header_start, header_end + 1)
        for column in range(candidate.min_col, candidate.max_col + 1)
    ]
    header_row_coverages = [
        sum(
            not is_missing(effective_value(lookup.get((row, column))))
            for column in range(candidate.min_col, candidate.max_col + 1)
        )
        / width
        for row in range(header_start, header_end + 1)
    ]
    header_row_origin_coverages: list[float] = []
    for row in range(header_start, header_end + 1):
        row_records = [
            lookup.get((row, column))
            for column in range(candidate.min_col, candidate.max_col + 1)
        ]
        origins = {
            record.get("merged_anchor") or record.get("cell_address")
            for record in row_records
            if record and not is_missing(effective_value(record))
        }
        header_row_origin_coverages.append(len(origins) / width)
    body_records = [
        lookup.get((row, column))
        for row in range(body_start, candidate.max_row + 1)
        for column in range(candidate.min_col, candidate.max_col + 1)
    ]
    header_present = [
        record for record in header_records if not is_missing(effective_value(record))
    ]
    body_present = [
        record for record in body_records if not is_missing(effective_value(record))
    ]

    if header_depth == 0:
        header_like_ratio = 0.35
        completeness = 0.40
    else:
        header_like_ratio = (
            sum(_is_header_like(record) for record in header_present)
            / max(1, len(header_present))
        )
        columns_with_header = sum(
            any(
                not is_missing(effective_value(lookup.get((row, column))))
                for row in range(header_start, header_end + 1)
            )
            for column in range(candidate.min_col, candidate.max_col + 1)
        )
        completeness = columns_with_header / width

    numeric_body_ratio = (
        sum(_is_numeric_like(record) for record in body_present)
        / max(1, len(body_present))
    )
    body_rows = list(range(body_start, candidate.max_row + 1))
    row_coverages = [
        sum(
            not is_missing(effective_value(lookup.get((row, column))))
            for column in range(candidate.min_col, candidate.max_col + 1)
        )
        / width
        for row in body_rows
    ]
    body_coverage = statistics.median(row_coverages) if row_coverages else 0.0

    header_paths: dict[int, list[Any]] = {}
    duplicates = 0
    seen: Counter[tuple[str, ...]] = Counter()
    for column in range(candidate.min_col, candidate.max_col + 1):
        if header_depth:
            path = [
                effective_value(lookup.get((row, column)))
                for row in range(header_start, header_end + 1)
            ]
            path = [value for value in path if not is_missing(value)]
        else:
            path = [f"column_{column - candidate.min_col + 1}"]
        header_paths[column] = path
        key = tuple(normalize_label(item) for item in path)
        seen[key] += 1
    duplicates = sum(count - 1 for count in seen.values() if count > 1)
    uniqueness = 1.0 - duplicates / max(1, width)

    sparse_header_penalty = 0.0
    if header_depth and width >= 4:
        sparse_header_penalty += 0.12 * sum(
            coverage <= 0.40 for coverage in header_row_coverages
        )
        sparse_header_penalty += 0.30 * sum(
            coverage <= 0.20
            for coverage in header_row_origin_coverages
        )
    score = (
        0.20 * title_score
        + 0.22 * header_like_ratio
        + 0.18 * completeness
        + 0.18 * min(1.0, numeric_body_ratio + 0.25)
        + 0.12 * body_coverage
        + 0.10 * uniqueness
        - sparse_header_penalty
    )

    if header_like_ratio >= 0.7:
        reasons.append("Candidate header is predominantly label-like.")
    if completeness >= 0.8:
        reasons.append("Most columns have a resolved header path.")
    if numeric_body_ratio >= 0.4:
        reasons.append("Body contains a substantial numeric area.")
    if duplicates:
        reasons.append(f"{duplicates} duplicate header path(s) remain.")

    return score, header_paths, reasons


def _detect_row_label_columns(
    candidate: BlockCandidate,
    lookup: dict[tuple[int, int], dict[str, Any]],
    data_start_row: int,
    maximum: int,
) -> int:
    """Estimate the number of leading dimension or row-label columns."""

    result = 0
    for column in range(
        candidate.min_col,
        min(candidate.max_col, candidate.min_col + maximum - 1) + 1,
    ):
        values = [
            effective_value(lookup.get((row, column)))
            for row in range(data_start_row, candidate.max_row + 1)
        ]
        present = [value for value in values if not is_missing(value)]
        if not present:
            break
        text_ratio = sum(isinstance(value, str) for value in present) / len(
            present
        )
        numeric_ratio = sum(is_number(value) for value in present) / len(present)
        if text_ratio >= 0.45 and numeric_ratio <= 0.40:
            result += 1
        else:
            break
    return result


def _column_role(
    header_path: list[Any],
    values: list[Any],
) -> str:
    """Infer a generic semantic role for one source column."""

    label = normalize_label(header_path[-1] if header_path else "")
    if label in {"period", "date", "year", "month", "quarter"}:
        return "period"
    if label in {"kpi", "kpi name", "metric", "measure"}:
        return "kpi"
    if label in {"unit", "currency"}:
        return label
    if label in {"scenario", "version"}:
        return "scenario"
    if label in {"value", "amount"}:
        return "measure"

    present = [value for value in values if not is_missing(value)]
    if not present:
        return "unknown"
    numeric_ratio = sum(is_number(value) for value in present) / len(present)
    text_ratio = sum(isinstance(value, str) for value in present) / len(present)
    period_ratio = sum(
        parse_period_label(value)["period_parse_confidence"] >= 0.75
        for value in present
    ) / len(present)
    if period_ratio >= 0.7:
        return "period"
    if numeric_ratio >= 0.7:
        return "measure"
    if text_ratio >= 0.6:
        return "dimension"
    return "unknown"


def _classify_block(
    candidate: BlockCandidate,
    header_depth: int,
    row_label_columns: int,
    header_paths: dict[int, list[Any]],
    lookup: dict[tuple[int, int], dict[str, Any]],
    data_start_row: int,
) -> tuple[str, dict[int, str], list[str]]:
    """Classify the block and infer source-column roles."""

    reasons: list[str] = []
    roles: dict[int, str] = {}
    for column, path in header_paths.items():
        values = [
            effective_value(lookup.get((row, column)))
            for row in range(data_start_row, candidate.max_row + 1)
        ]
        roles[column] = _column_role(path, values)

    normalized_headers = {
        normalize_label(path[-1])
        for path in header_paths.values()
        if path
    }
    known_header_hits = len(normalized_headers & KNOWN_RECORD_FIELDS)
    measure_columns = sum(role == "measure" for role in roles.values())
    dimension_columns = sum(role == "dimension" for role in roles.values())

    body_rows = candidate.max_row - data_start_row + 1
    per_row_non_empty = []
    per_row_text = []
    for row in range(data_start_row, candidate.max_row + 1):
        values = [
            effective_value(lookup.get((row, column)))
            for column in range(candidate.min_col, candidate.max_col + 1)
        ]
        present = [value for value in values if not is_missing(value)]
        per_row_non_empty.append(len(present))
        per_row_text.append(sum(isinstance(value, str) for value in present))

    median_non_empty = statistics.median(per_row_non_empty) if per_row_non_empty else 0
    median_text = statistics.median(per_row_text) if per_row_text else 0

    if (
        row_label_columns >= 1
        and candidate.width - row_label_columns >= 2
        and measure_columns >= 2
        and (
            header_depth >= 2
            or not ({"value", "amount"} & normalized_headers)
        )
    ):
        reasons.append("Leading labels are followed by multiple numeric columns.")
        return "cross_tab_matrix", roles, reasons

    if known_header_hits >= 2 and header_depth >= 1:
        reasons.append("Header names match a record-oriented schema.")
        return "rectangular_record_table", roles, reasons

    if candidate.width <= 3 and row_label_columns >= 1 and measure_columns >= 1:
        reasons.append("Compact label/value structure detected.")
        return "kpi_value_block", roles, reasons

    if median_non_empty <= 2 and median_text >= 1 and body_rows >= 2:
        reasons.append("Sparse label/value rows resemble a form.")
        return "form_like_block", roles, reasons

    all_values = [
        effective_value(lookup.get((row, column)))
        for row in range(candidate.min_row, candidate.max_row + 1)
        for column in range(candidate.min_col, candidate.max_col + 1)
    ]
    present = [value for value in all_values if not is_missing(value)]
    text_ratio = (
        sum(isinstance(value, str) for value in present) / max(1, len(present))
    )
    if text_ratio >= 0.85:
        reasons.append("Region is predominantly narrative text.")
        return "notes_block", roles, reasons

    if candidate.official_table_name:
        reasons.append("Official Excel table retained as a record table.")
        return "rectangular_record_table", roles, reasons

    return "unclassified", roles, reasons


def _row_types(
    candidate: BlockCandidate,
    lookup: dict[tuple[int, int], dict[str, Any]],
    data_start_row: int,
    row_label_columns: int,
    sheet_rule: SheetRule | None,
) -> dict[int, str]:
    """Classify body rows as detail, subtotal, total, or grand total."""

    total_labels = {
        normalize_label(value)
        for value in (
            sheet_rule.known_total_labels if sheet_rule else ("total",)
        )
    }
    subtotal_labels = {
        normalize_label(value)
        for value in (
            sheet_rule.known_subtotal_labels if sheet_rule else ("subtotal",)
        )
    }
    grand_total_labels = {
        normalize_label(value)
        for value in (
            sheet_rule.known_grand_total_labels
            if sheet_rule
            else ("grand total",)
        )
    }

    result: dict[int, str] = {}
    for row in range(data_start_row, candidate.max_row + 1):
        labels = [
            normalize_label(effective_value(lookup.get((row, column))))
            for column in range(
                candidate.min_col,
                min(
                    candidate.max_col,
                    candidate.min_col + max(1, row_label_columns) - 1,
                )
                + 1,
            )
        ]
        labels = [label for label in labels if label]
        formulas = [
            str(lookup.get((row, column), {}).get("formula") or "").upper()
            for column in range(candidate.min_col, candidate.max_col + 1)
        ]
        if any(label in grand_total_labels for label in labels):
            result[row] = "grand_total"
        elif any(label in subtotal_labels for label in labels) or any(
            "SUBTOTAL(" in formula for formula in formulas
        ):
            result[row] = "subtotal"
        elif any(label in total_labels for label in labels) or any(
            "SUM(" in formula for formula in formulas
        ):
            result[row] = "total"
        else:
            result[row] = "detail"
    return result


def interpret_block(
    candidate: BlockCandidate,
    raw_cells: list[dict[str, Any]],
    config: ExtractionConfig,
    workbook_id: str,
    sheet_rule: SheetRule | None = None,
) -> BlockInterpretation:
    """Interpret headers, roles, and row types for a candidate block."""

    lookup = _cell_lookup(raw_cells)
    max_header_depth = min(
        candidate.height - 1,
        sheet_rule.maximum_header_depth
        if sheet_rule
        else config.maximum_header_depth,
    )
    best_score = -1.0
    best_title_rows = 0
    best_header_depth = 0
    best_header_paths: dict[int, list[Any]] = {}
    best_reasons: list[str] = []

    max_title_rows = min(config.maximum_title_rows, candidate.height - 1)
    for title_rows in range(max_title_rows + 1):
        remaining = candidate.height - title_rows
        for header_depth in range(0, min(max_header_depth, remaining - 1) + 1):
            score, paths, reasons = _header_candidate_score(
                candidate,
                lookup,
                title_rows,
                header_depth,
            )
            if score > best_score:
                best_score = score
                best_title_rows = title_rows
                best_header_depth = header_depth
                best_header_paths = paths
                best_reasons = reasons

    data_start_row = candidate.min_row + best_title_rows + best_header_depth
    row_label_columns = _detect_row_label_columns(
        candidate,
        lookup,
        data_start_row,
        config.maximum_row_label_columns,
    )
    block_class, column_roles, class_reasons = _classify_block(
        candidate,
        best_header_depth,
        row_label_columns,
        best_header_paths,
        lookup,
        data_start_row,
    )
    reasons = best_reasons + class_reasons

    confidence = max(
        0.0,
        min(
            1.0,
            0.55 * best_score
            + 0.25 * candidate.detection_confidence
            + (0.15 if block_class != "unclassified" else 0.0)
            + (0.05 if candidate.official_table_name else 0.0),
        ),
    )
    if block_class in {"notes_block", "form_like_block"}:
        status = "classified_non_analytical"
    elif confidence >= config.accept_confidence and block_class != "unclassified":
        status = "accepted"
    elif confidence >= config.review_confidence:
        status = "manual_review"
    else:
        status = "unresolved"

    warnings: list[ExtractionWarning] = []
    header_keys = [
        tuple(normalize_label(value) for value in path)
        for path in best_header_paths.values()
    ]
    duplicate_count = sum(
        count - 1 for count in Counter(header_keys).values() if count > 1
    )
    if duplicate_count:
        warnings.append(
            ExtractionWarning(
                warning_code="BL003",
                severity=WarningSeverity.WARNING,
                message=(
                    f"{duplicate_count} duplicate resolved header path(s) "
                    "were detected."
                ),
                workbook_id=workbook_id,
                sheet_name=candidate.sheet_name,
                block_id=candidate.block_id,
                source_range=candidate.source_range,
                suggested_action="Review multi-row header interpretation.",
            )
        )
    if status in {"manual_review", "unresolved"}:
        warnings.append(
            ExtractionWarning(
                warning_code="BL002",
                severity=(
                    WarningSeverity.WARNING
                    if status == "manual_review"
                    else WarningSeverity.ERROR
                ),
                message=(
                    "Block could not be interpreted with sufficient confidence "
                    f"for automatic normalization ({confidence:.3f})."
                ),
                workbook_id=workbook_id,
                sheet_name=candidate.sheet_name,
                block_id=candidate.block_id,
                source_range=candidate.source_range,
                suggested_action="Review the detected range and header split.",
            )
        )

    return BlockInterpretation(
        candidate=candidate,
        block_class=block_class,
        title_rows=best_title_rows,
        header_depth=best_header_depth,
        row_label_columns=row_label_columns,
        confidence=confidence,
        status=status,
        reasons=reasons,
        header_paths=best_header_paths,
        column_roles=column_roles,
        row_types=_row_types(
            candidate,
            lookup,
            data_start_row,
            row_label_columns,
            sheet_rule,
        ),
        warnings=warnings,
    )


def block_record(
    interpretation: BlockInterpretation,
    workbook_id: str,
) -> dict[str, Any]:
    """Convert an interpretation into a detected-block record."""

    candidate = interpretation.candidate
    return {
        "workbook_id": workbook_id,
        "block_id": candidate.block_id,
        "sheet_name": candidate.sheet_name,
        "source_range": candidate.source_range,
        "block_class": interpretation.block_class,
        "detection_method": candidate.detection_method,
        "detection_confidence": candidate.detection_confidence,
        "interpretation_confidence": interpretation.confidence,
        "interpretation_status": interpretation.status,
        "official_table_name": candidate.official_table_name,
        "occupied_cell_count": candidate.occupied_cells,
        "min_row": candidate.min_row,
        "max_row": candidate.max_row,
        "min_column": candidate.min_col,
        "max_column": candidate.max_col,
        "title_row_count": interpretation.title_rows,
        "header_row_start": interpretation.header_start_row,
        "header_row_end": interpretation.header_end_row,
        "header_depth": interpretation.header_depth,
        "row_label_column_count": interpretation.row_label_columns,
        "data_row_start": interpretation.data_start_row,
        "data_row_end": candidate.max_row,
        "data_column_start": candidate.min_col,
        "data_column_end": candidate.max_col,
        "interpretation_reasons_json": str(interpretation.reasons),
    }
