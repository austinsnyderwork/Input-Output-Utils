
import os
from typing import Union

import pandas as pd


def excel_only(func):
    def wrapper(self, *args, **kwargs):
        if not isinstance(self._manager, _ExcelManager):
            raise TypeError(f"{func.__name__} can only be used for Excel output files.\n"
                            f"Output path {self._output_path} is not an Excel path.")
        return func(self, *args, **kwargs)
    return wrapper


class _ExcelManager:

    def __init__(self, output_path: str):
        self._writer = pd.ExcelWriter(output_path, engine='openpyxl')
        self._sheet_data = dict()

    def add_data(self, data: pd.DataFrame, sheet_name: str):
        self._sheet_data[sheet_name] = data
        data.to_excel(self._writer, sheet_name=sheet_name, index=False)

    def _fetch_worksheet(self, sheet_name: str):
        return self._writer.sheets[sheet_name]

    def autofit_column_widths(self, sheet_name: str=None):
        """

        :param sheet_name: If provided, only autofits column widths for the singular sheet. Otherwise, autofit column widths for all sheets.
        :return:
        """
        from openpyxl.utils import get_column_letter

        sheet_names = [sheet_name] if sheet_name else list(self._sheet_data.keys())

        for sheet_name in sheet_names:
            worksheet = self._fetch_worksheet(sheet_name=sheet_name)
            data = self._sheet_data[sheet_name]

            for i in range(data.shape[1]):
                col_series = data.iloc[:, i]
                col_name = data.columns[i]
                max_len = max(col_series.astype(str).map(len).max(), len(str(col_name)))
                worksheet.column_dimensions[get_column_letter(i + 1)].width = max_len + 2

    def export(self):
        self._writer.close()



class _CsvManager:

    def __init__(self,
                 output_path: str,
                 data: pd.DataFrame = None):
        self._output_path = output_path
        self._data = data

    def add_data(self, data: pd.DataFrame):
        print("Overwriting previous data in CsvManager.")
        self._data = data

    def export(self):
        if not self._data:
            raise ValueError(f"CsvManager has no data to export.")

        self._data.to_csv(self._output_path)


class OutputManager:
    _manager: Union[_ExcelManager, _CsvManager]

    def __init__(self,
                 output_path: str,
                 data: pd.DataFrame = None,
                 sheet_name: str = None):
        self.output_path = output_path

        ext = os.path.splitext(output_path)[1].lower()
        if ext in ('.xlsx', '.xls'):
            self._manager = _ExcelManager(output_path)
        elif ext == '.csv':
            self._manager = _CsvManager(output_path)
        else:
            raise ValueError(f"Output path {output_path} is not supported.")

        if data:
            self.add_data(
                data=data,
                sheet_name=sheet_name
            )

    @property
    def is_excel(self):
        return isinstance(self._manager, _ExcelManager)

    @property
    def is_csv(self):
        return isinstance(self._manager, _CsvManager)

    @excel_only
    def autofit_column_widths(self, sheet_name: str = None):
        self._manager.autofit_column_widths(sheet_name=sheet_name)

    def add_data(self, data: pd.DataFrame | dict, sheet_name: str = None):
        if self.is_excel:
            if isinstance(data, pd.DataFrame) and not sheet_name:
                raise ValueError(f"Sheet name must be included as an argument when the a singular DataFrame is supplied "
                                 f"and the output is Excel.")

            if isinstance(data, dict):
                non_dataframes = [v for v in data.values() if not isinstance(v, pd.DataFrame)]
                if non_dataframes:
                    raise TypeError(f"Non-DataFrames supplied as values in dict when calling add_data.\n"
                                    f"Invalid types: {[type(non_df) for non_df in non_dataframes]}")

                for sheet_name, df in data.items():
                    self._manager.add_data(
                        data=df,
                        sheet_name=sheet_name
                    )

            return

        if self.is_csv:
            if sheet_name:
                raise ValueError(f"Sheet name is an invalid argument when the output is a CSV file.")

            if isinstance(data, dict):
                raise TypeError(f"Supplying a dict to add_data provides sheet names for each DataFrame.\n"
                                f"This is invalid when the output is CSV.")

            self._manager.add_data(
                data=data
            )

            return

    def export(self):
        self._manager.export()
