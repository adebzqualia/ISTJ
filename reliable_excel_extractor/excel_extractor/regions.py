"""Candidate region and block detection."""

from __future__ import annotations

from collections import defaultdict
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from .config import ExtractionConfig, SheetRule
from .models import BlockCandidate
from .utils import a1_range, is_missing, parse_a1_range


class _UnionFind:
    """Minimal union-find implementation for occupied-cell components."""

    def __init__(self, items: set[tuple[int, int]]) -> None:
        self.parent = {item: item for item in items}
        self.rank = {item: 0 for item in items}

    def find(self, item: tuple[int, int]) -> tuple[int, int]:
        """Return the representative for an item."""

        parent = self.parent[item]
        if parent != item:
            self.parent[item] = self.find(parent)
        return self.parent[item]

    def union(self, left: tuple[int, int], right: tuple[int, int]) -> None:
        """Join two components."""

        left_root = self.find(left)
        right_root = self.find(right)
        if left_root == right_root:
            return
        if self.rank[left_root] < self.rank[right_root]:
            left_root, right_root = right_root, left_root
        self.parent[right_root] = left_root
        if self.rank[left_root] == self.rank[right_root]:
            self.rank[left_root] += 1


def _occupied_coordinates(
    raw_cells: list[dict[str, Any]],
) -> set[tuple[int, int]]:
    """Return coordinates that visually contain meaningful information."""

    occupied: set[tuple[int, int]] = set()
    for record in raw_cells:
        value = record["raw_value"]
        if record["formula"] is not None:
            occupied.add((record["row_index"], record["column_index"]))
        elif not is_missing(value):
            occupied.add((record["row_index"], record["column_index"]))
        elif not is_missing(record["resolved_merged_value"]):
            occupied.add((record["row_index"], record["column_index"]))
    return occupied


def _generic_components(
    occupied: set[tuple[int, int]],
    config: ExtractionConfig,
) -> list[set[tuple[int, int]]]:
    """Join occupied cells across small horizontal or vertical gaps."""

    if not occupied:
        return []

    union_find = _UnionFind(occupied)
    rows: dict[int, list[int]] = defaultdict(list)
    columns: dict[int, list[int]] = defaultdict(list)
    for row, column in occupied:
        rows[row].append(column)
        columns[column].append(row)

    horizontal_limit = config.max_blank_column_gap + 1
    vertical_limit = config.max_blank_row_gap + 1

    for row, row_columns in rows.items():
        ordered = sorted(row_columns)
        for left, right in zip(ordered, ordered[1:]):
            if right - left <= horizontal_limit:
                union_find.union((row, left), (row, right))

    for column, column_rows in columns.items():
        ordered = sorted(column_rows)
        for top, bottom in zip(ordered, ordered[1:]):
            if bottom - top <= vertical_limit:
                union_find.union((top, column), (bottom, column))

    components: dict[tuple[int, int], set[tuple[int, int]]] = defaultdict(set)
    for item in occupied:
        components[union_find.find(item)].add(item)
    return list(components.values())




def _split_component_on_blank_axes(
    component: set[tuple[int, int]],
    minimum_cells: int,
) -> list[set[tuple[int, int]]]:
    """Split a component when a fully blank internal row or column exists.

    Small blank gaps are allowed during connected-component construction so
    sparse tables remain connected. This second pass prevents that tolerance
    from merging independent side-by-side or vertically stacked tables.
    """

    if len(component) < 2 * minimum_cells:
        return [component]

    rows = [row for row, _ in component]
    columns = [column for _, column in component]
    min_row, max_row = min(rows), max(rows)
    min_col, max_col = min(columns), max(columns)
    occupied_rows = set(rows)
    occupied_columns = set(columns)

    blank_columns = [
        column
        for column in range(min_col + 1, max_col)
        if column not in occupied_columns
    ]
    for split_column in blank_columns:
        left = {item for item in component if item[1] < split_column}
        right = {item for item in component if item[1] > split_column}
        if len(left) >= minimum_cells and len(right) >= minimum_cells:
            return _split_component_on_blank_axes(
                left, minimum_cells
            ) + _split_component_on_blank_axes(right, minimum_cells)

    blank_rows = [
        row
        for row in range(min_row + 1, max_row)
        if row not in occupied_rows
    ]
    for split_row in blank_rows:
        top = {item for item in component if item[0] < split_row}
        bottom = {item for item in component if item[0] > split_row}
        if len(top) >= minimum_cells and len(bottom) >= minimum_cells:
            return _split_component_on_blank_axes(
                top, minimum_cells
            ) + _split_component_on_blank_axes(bottom, minimum_cells)

    return [component]


def _candidate_from_bounds(
    *,
    sheet_name: str,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
    method: str,
    occupied_cells: int,
    confidence: float,
    official_table_name: str | None = None,
) -> BlockCandidate:
    """Create a stable candidate block from numeric bounds."""

    source_range = a1_range(min_row, max_row, min_col, max_col)
    safe_sheet = sheet_name.replace(" ", "_")
    block_id = f"{safe_sheet}__{source_range.replace(':', '_')}"
    if official_table_name:
        block_id = f"{block_id}__{official_table_name}"
    return BlockCandidate(
        block_id=block_id,
        sheet_name=sheet_name,
        min_row=min_row,
        max_row=max_row,
        min_col=min_col,
        max_col=max_col,
        detection_method=method,
        source_range=source_range,
        detection_confidence=confidence,
        occupied_cells=occupied_cells,
        official_table_name=official_table_name,
    )


def detect_candidate_blocks(
    worksheet: Worksheet,
    raw_cells: list[dict[str, Any]],
    config: ExtractionConfig,
    sheet_rule: SheetRule | None = None,
) -> list[BlockCandidate]:
    """Detect explicit and generic candidate blocks in a worksheet."""

    occupied = _occupied_coordinates(raw_cells)
    candidates: list[BlockCandidate] = []
    explicit_ranges: set[tuple[int, int, int, int]] = set()

    for name in worksheet.tables:
        table = worksheet.tables[name]
        min_row, max_row, min_col, max_col = parse_a1_range(table.ref)
        explicit_ranges.add((min_row, max_row, min_col, max_col))
        count = sum(
            min_row <= row <= max_row and min_col <= column <= max_col
            for row, column in occupied
        )
        candidates.append(
            _candidate_from_bounds(
                sheet_name=worksheet.title,
                min_row=min_row,
                max_row=max_row,
                min_col=min_col,
                max_col=max_col,
                method="official_excel_table",
                occupied_cells=count,
                confidence=1.0,
                official_table_name=name,
            )
        )

    if sheet_rule:
        for known_range in sheet_rule.known_ranges:
            min_row, max_row, min_col, max_col = parse_a1_range(known_range)
            explicit_ranges.add((min_row, max_row, min_col, max_col))
            count = sum(
                min_row <= row <= max_row and min_col <= column <= max_col
                for row, column in occupied
            )
            candidates.append(
                _candidate_from_bounds(
                    sheet_name=worksheet.title,
                    min_row=min_row,
                    max_row=max_row,
                    min_col=min_col,
                    max_col=max_col,
                    method="template_known_range",
                    occupied_cells=count,
                    confidence=0.98,
                )
            )

    generic_components: list[set[tuple[int, int]]] = []
    for component in _generic_components(occupied, config):
        generic_components.extend(
            _split_component_on_blank_axes(
                component,
                config.minimum_region_cells,
            )
        )

    for component in generic_components:
        rows = [item[0] for item in component]
        columns = [item[1] for item in component]
        min_row, max_row = min(rows), max(rows)
        min_col, max_col = min(columns), max(columns)
        bounds = (min_row, max_row, min_col, max_col)

        if bounds in explicit_ranges:
            continue
        generic_area = (max_row - min_row + 1) * (max_col - min_col + 1)
        overlaps_explicit = False
        for exp_min_row, exp_max_row, exp_min_col, exp_max_col in explicit_ranges:
            intersection_rows = max(
                0,
                min(max_row, exp_max_row) - max(min_row, exp_min_row) + 1,
            )
            intersection_columns = max(
                0,
                min(max_col, exp_max_col) - max(min_col, exp_min_col) + 1,
            )
            intersection = intersection_rows * intersection_columns
            explicit_area = (
                (exp_max_row - exp_min_row + 1)
                * (exp_max_col - exp_min_col + 1)
            )
            if intersection and (
                intersection / generic_area >= 0.50
                or intersection / explicit_area >= 0.80
            ):
                overlaps_explicit = True
                break
        if overlaps_explicit:
            continue
        height = max_row - min_row + 1
        width = max_col - min_col + 1
        if len(component) < config.minimum_region_cells:
            continue
        if height < config.minimum_region_rows:
            continue
        if width < config.minimum_region_columns:
            continue

        density = len(component) / (height * width)
        confidence = min(0.95, 0.55 + 0.40 * density)
        candidates.append(
            _candidate_from_bounds(
                sheet_name=worksheet.title,
                min_row=min_row,
                max_row=max_row,
                min_col=min_col,
                max_col=max_col,
                method="connected_region",
                occupied_cells=len(component),
                confidence=confidence,
            )
        )

    candidates.sort(
        key=lambda block: (
            block.min_row,
            block.min_col,
            -block.detection_confidence,
        )
    )
    return candidates
