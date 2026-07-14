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
