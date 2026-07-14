# Reliable Excel Extractor

A small Python package for inspecting one complex Excel workbook and extracting
source-faithful raw cells, detected blocks, interpreted structures, normalized
observations, and explicit warnings.

## Design guarantees

- The source workbook is never modified.
- Formula text and cached values are stored separately.
- Merged ranges are preserved and resolved through their anchor cells.
- Every normalized value retains source sheet, range, cell, header cells, and
  row-label cells.
- Low-confidence blocks are returned for review instead of silently normalized.
- Raw extraction and analytical normalization are separate output layers.

## Supported files

- `.xlsx`
- `.xlsm`
- `.xltx`
- `.xltm`

`openpyxl` does not calculate formulas. Formula cells without cached values are
reported with warning code `CL001`.

## Installation

```bash
python -m pip install -e .
```

## Command line

```bash
extract-excel input.xlsx --output extracted_output
```

With a template configuration:

```bash
extract-excel input.xlsx \
  --config examples/config.json \
  --output extracted_output
```

The output directory contains:

```text
workbook_inventory.csv
sheet_inventory.csv
raw_cells.csv
detected_blocks.csv
block_columns.csv
block_rows.csv
analytical_observations.csv
extraction_warnings.csv
extraction_metadata.json
```

## Programmatic use

```python
from excel_extractor import ExcelExtractor, ExtractionConfig

config = ExtractionConfig(
    accept_confidence=0.75,
    propagate_blank_hierarchy_labels=False,
)

result = ExcelExtractor(config).extract("input.xlsx")
result.to_directory("extracted_output")

observations = result["analytical_observations"]
warnings = result["extraction_warnings"]
```

## Template-specific rules

Use a configuration for known sheets, anchor labels, expected ranges, canonical
KPI names, and exact total labels. Explicit ranges are evaluated before generic
connected-region detection.

```json
{
  "template": {
    "required_sheets": ["KPI Summary"],
    "known_kpi_aliases": {
      "rev": "Revenue"
    },
    "sheet_rules": {
      "KPI Summary": {
        "expected_anchor_labels": ["KPI Performance"],
        "known_ranges": ["B2:F20"],
        "maximum_header_depth": 3
      }
    }
  }
}
```

## Interpretation policy

Blocks can have these statuses:

- `accepted`: normalized automatically.
- `manual_review`: structure is plausible but not sufficiently confident.
- `unresolved`: no safe interpretation was found.
- `classified_non_analytical`: notes or form-like content retained as metadata.

The initial generic heuristics are intentionally conservative. For a stable
business template, add deterministic `known_ranges`, anchor labels, total
labels, and KPI aliases before lowering confidence thresholds.

## Tests

```bash
pytest -q
```
