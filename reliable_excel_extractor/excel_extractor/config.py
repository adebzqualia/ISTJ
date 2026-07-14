"""Configuration models for workbook extraction."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass(frozen=True)
class SheetRule:
    """Describe deterministic expectations for one worksheet.

    :param expected_anchor_labels: Labels that should exist on the sheet.
    :param known_ranges: Explicit ranges to inspect before generic detection.
    :param maximum_header_depth: Maximum number of header rows to evaluate.
    :param known_total_labels: Exact labels interpreted as totals.
    :param known_subtotal_labels: Exact labels interpreted as subtotals.
    :param known_grand_total_labels: Exact labels interpreted as grand totals.
    :param required: Whether absence of the sheet is blocking.
    """

    expected_anchor_labels: tuple[str, ...] = ()
    known_ranges: tuple[str, ...] = ()
    maximum_header_depth: int = 5
    known_total_labels: tuple[str, ...] = ("total",)
    known_subtotal_labels: tuple[str, ...] = ("subtotal",)
    known_grand_total_labels: tuple[str, ...] = ("grand total",)
    required: bool = False


@dataclass(frozen=True)
class TemplateConfig:
    """Describe workbook-level template rules.

    :param required_sheets: Sheet names that must be present.
    :param optional_sheets: Known but optional sheet names.
    :param sheet_rules: Per-sheet extraction rules.
    :param known_kpi_aliases: Mapping from normalized aliases to canonical KPIs.
    """

    required_sheets: tuple[str, ...] = ()
    optional_sheets: tuple[str, ...] = ()
    sheet_rules: dict[str, SheetRule] = field(default_factory=dict)
    known_kpi_aliases: dict[str, str] = field(default_factory=dict)


@dataclass(frozen=True)
class ExtractionConfig:
    """Control generic extraction and interpretation behavior.

    :param include_hidden_sheets: Whether hidden sheets are inspected.
    :param include_styled_empty_cells: Whether style-only cells enter raw output.
    :param include_comments: Whether comment text is retained.
    :param include_style_details: Whether simplified style metadata is retained.
    :param max_blank_row_gap: Blank rows allowed when joining occupied cells.
    :param max_blank_column_gap: Blank columns allowed when joining occupied cells.
    :param minimum_region_cells: Minimum occupied cells in a generic region.
    :param minimum_region_rows: Minimum height of a generic region.
    :param minimum_region_columns: Minimum width of a generic region.
    :param maximum_header_depth: Generic maximum header depth.
    :param maximum_title_rows: Maximum sparse title rows before a header.
    :param maximum_row_label_columns: Maximum leading row-label columns.
    :param accept_confidence: Threshold for automatic normalization.
    :param review_confidence: Threshold below which a block is unresolved.
    :param propagate_blank_hierarchy_labels: Permit cautious non-merge filling.
    :param evaluate_simple_sum_formulas: Reconcile simple SUM formulas.
    :param template: Optional deterministic template configuration.
    """

    include_hidden_sheets: bool = True
    include_styled_empty_cells: bool = False
    include_comments: bool = True
    include_style_details: bool = True
    max_blank_row_gap: int = 1
    max_blank_column_gap: int = 1
    minimum_region_cells: int = 4
    minimum_region_rows: int = 2
    minimum_region_columns: int = 2
    maximum_header_depth: int = 5
    maximum_title_rows: int = 3
    maximum_row_label_columns: int = 5
    accept_confidence: float = 0.75
    review_confidence: float = 0.50
    propagate_blank_hierarchy_labels: bool = False
    evaluate_simple_sum_formulas: bool = True
    template: TemplateConfig = field(default_factory=TemplateConfig)

    def to_dict(self) -> dict[str, Any]:
        """Return the configuration as a JSON-serializable dictionary."""

        return asdict(self)

    @classmethod
    def from_dict(cls, payload: dict[str, Any]) -> "ExtractionConfig":
        """Construct a configuration from a dictionary.

        :param payload: Parsed configuration dictionary.
        :return: Validated extraction configuration.
        """

        data = dict(payload)
        template_payload = data.pop("template", {}) or {}
        rule_payloads = template_payload.pop("sheet_rules", {}) or {}
        sheet_rules = {
            name: SheetRule(
                **{
                    key: tuple(value) if isinstance(value, list) else value
                    for key, value in rule.items()
                }
            )
            for name, rule in rule_payloads.items()
        }
        template = TemplateConfig(
            required_sheets=tuple(template_payload.pop("required_sheets", ())),
            optional_sheets=tuple(template_payload.pop("optional_sheets", ())),
            sheet_rules=sheet_rules,
            known_kpi_aliases=dict(
                template_payload.pop("known_kpi_aliases", {})
            ),
            **template_payload,
        )
        return cls(template=template, **data)
