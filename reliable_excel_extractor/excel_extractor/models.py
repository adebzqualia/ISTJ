"""Shared data models used by the extraction pipeline."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path
from typing import Any

import pandas as pd


class WarningSeverity(StrEnum):
    """Supported extraction-warning severities."""

    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    BLOCKING = "blocking"


@dataclass(frozen=True)
class ExtractionWarning:
    """Represent an auditable extraction warning."""

    warning_code: str
    severity: WarningSeverity
    message: str
    workbook_id: str | None = None
    sheet_name: str | None = None
    block_id: str | None = None
    cell_address: str | None = None
    source_range: str | None = None
    suggested_action: str | None = None

    def to_record(self) -> dict[str, Any]:
        """Convert the warning to a tabular record."""

        return {
            "warning_code": self.warning_code,
            "severity": self.severity.value,
            "message": self.message,
            "workbook_id": self.workbook_id,
            "sheet_name": self.sheet_name,
            "block_id": self.block_id,
            "cell_address": self.cell_address,
            "source_range": self.source_range,
            "suggested_action": self.suggested_action,
        }


@dataclass(frozen=True)
class CellRef:
    """Identify one cell in one worksheet."""

    sheet_name: str
    row: int
    column: int
    address: str


@dataclass(frozen=True)
class BlockCandidate:
    """Represent a rectangular candidate data block."""

    block_id: str
    sheet_name: str
    min_row: int
    max_row: int
    min_col: int
    max_col: int
    detection_method: str
    source_range: str
    detection_confidence: float
    occupied_cells: int
    official_table_name: str | None = None

    @property
    def height(self) -> int:
        """Return block height in rows."""

        return self.max_row - self.min_row + 1

    @property
    def width(self) -> int:
        """Return block width in columns."""

        return self.max_col - self.min_col + 1


@dataclass
class BlockInterpretation:
    """Contain the interpreted structure of one candidate block."""

    candidate: BlockCandidate
    block_class: str
    title_rows: int
    header_depth: int
    row_label_columns: int
    confidence: float
    status: str
    reasons: list[str] = field(default_factory=list)
    header_paths: dict[int, list[Any]] = field(default_factory=dict)
    column_roles: dict[int, str] = field(default_factory=dict)
    row_types: dict[int, str] = field(default_factory=dict)
    warnings: list[ExtractionWarning] = field(default_factory=list)

    @property
    def header_start_row(self) -> int | None:
        """Return the first header row, when a header exists."""

        if self.header_depth == 0:
            return None
        return self.candidate.min_row + self.title_rows

    @property
    def header_end_row(self) -> int | None:
        """Return the final header row, when a header exists."""

        start = self.header_start_row
        if start is None:
            return None
        return start + self.header_depth - 1

    @property
    def data_start_row(self) -> int:
        """Return the first body row."""

        return self.candidate.min_row + self.title_rows + self.header_depth


@dataclass
class ExtractionResult:
    """Hold all tabular outputs from an extraction run."""

    tables: dict[str, pd.DataFrame]
    metadata: dict[str, Any] = field(default_factory=dict)

    def __getitem__(self, name: str) -> pd.DataFrame:
        """Return a named output table."""

        return self.tables[name]

    def to_directory(
        self,
        output_dir: str | Path,
        *,
        file_format: str = "csv",
    ) -> Path:
        """Write all outputs to a directory.

        :param output_dir: Destination directory.
        :param file_format: Either ``csv`` or ``jsonl``.
        :return: Resolved output directory.
        """

        destination = Path(output_dir)
        destination.mkdir(parents=True, exist_ok=True)

        for name, frame in self.tables.items():
            if file_format == "csv":
                frame.to_csv(destination / f"{name}.csv", index=False)
            elif file_format == "jsonl":
                frame.to_json(
                    destination / f"{name}.jsonl",
                    orient="records",
                    lines=True,
                    date_format="iso",
                )
            else:
                raise ValueError(
                    f"Unsupported output format: {file_format!r}"
                )

        with (destination / "extraction_metadata.json").open(
            "w", encoding="utf-8"
        ) as stream:
            json.dump(self.metadata, stream, indent=2, default=str)

        return destination.resolve()
