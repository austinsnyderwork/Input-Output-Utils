
import pandas as pd
from pathlib import Path


def easy_import(import_path: str | Path, sheet_name: str = None) -> pd.DataFrame:
    import_path = Path(import_path) if isinstance(import_path, str) else import_path

    if import_path.suffix == '.csv':
        return pd.read_csv(import_path)
    elif import_path.suffix == '.xlsx':
        if sheet_name:
            return pd.read_excel(import_path, sheet_name=sheet_name)
        else:
            return pd.read_excel(import_path)
    else:
        raise ValueError(f"Unsupported file extension: {import_path.suffix}")

def easy_export(output_path: str | Path, df: pd.DataFrame, sheet_name: str = None):
    output_path = Path(output_path) if isinstance(output_path, str) else output_path

    if output_path.suffix == '.csv':
        df.to_csv(output_path, index=False)
    elif output_path.suffix == '.xlsx':
        if sheet_name:
            df.to_excel(
                output_path,
                sheet_name=sheet_name,
                index=False
            )
        else:
            df.to_excel(output_path, index=False)