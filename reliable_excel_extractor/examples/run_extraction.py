"""Minimal programmatic extraction example."""

from pathlib import Path

from excel_extractor import ExcelExtractor, ExtractionConfig


workbook_path = Path("input.xlsx")
output_path = Path("extracted_output")

result = ExcelExtractor(ExtractionConfig()).extract(workbook_path)
result.to_directory(output_path)

print(result["detected_blocks"])
print(result["analytical_observations"].head())
print(result["extraction_warnings"])
