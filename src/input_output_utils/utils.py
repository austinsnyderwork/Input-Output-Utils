
import pandas as pd
from pathlib import Path

def easy_export(output_path: str | Path, df: pd.DataFrame, sheet_name: str = None):
    output_path = Path(output_path) if isinstance(output_path, str) else output_path

    if output_path.suffix == '.csv':
        df.to_csv(output_path)
    elif output_path.suffix == '.xlsx':
        df.to_excel(
            output_path,
            sheet_name=sheet_name
        )