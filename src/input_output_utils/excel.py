
import os
import xlsxwriter
from typing import Union
from enum import Enum
from pathlib import Path
import numpy as np

import pandas as pd


def excel_only(func):
    def wrapper(self, *args, **kwargs):
        if not isinstance(self._manager, ExcelManager):
            raise TypeError(f"{func.__name__} can only be used for Excel output files.\n"
                            f"Output path {self._output_path} is not an Excel path.")
        return func(self, *args, **kwargs)
    return wrapper


class HexColor(Enum):
    DARK_BLUE = 'ced8ff'
    BLUE = 'd8e0ff'
    LIGHT_BLUE = 'f1f4ff'
    DARK_GREEN = '2da15d'
    PASTEL_GREEN = 'a9c37d'
    LIGHT_PINK = 'FDD9FF'


class ExcelFormat:

    def __init__(self,
                 font_color: HexColor = None,
                 font_size: int = 14,
                 is_bold: bool = False,
                 fill_color: HexColor = None):
        self._font_color = font_color
        self._font_size = font_size
        self._is_bold = is_bold
        self._fill_color = fill_color

    def create_xlsx_writer_format(self) -> dict:
        f = {
            "font_size": self._font_size,
            "bold": self._is_bold
        }

        if self._font_color:
            f.update({"font_color": self._font_color.value})

        if self._fill_color:
            f.update({"bg_color": self._fill_color.value})

        return f

    def __hash__(self):
        return hash(
            (self._font_color, self._font_size, self._is_bold, self._fill_color)
        )


class CellFormatMap:

    def __init__(self):
        self._cell_map = dict()

    def format_cell(self, row_idx: int, col_idx: int, excel_format: ExcelFormat, exist_ok=False):
        t = (row_idx, col_idx)
        if t in self._cell_map and not exist_ok:
            raise KeyError(f"Cell at (row, col) ({row_idx}, {col_idx}) already exists in {self.__class__.__name__}.")

        self._cell_map[(row_idx, col_idx)] = excel_format

    def iter_cells(self):
        for (row_idx, col_idx), excel_format in self._cell_map.items():
            yield row_idx, col_idx, excel_format


class DataTable:

    def __init__(self,
                 df: pd.DataFrame):
        self.df = df

        self.format_map = CellFormatMap()

    @property
    def total_rows(self):
        # We increase the rows by 1 to include the title
        return self.df.shape[0] + 1

    @property
    def total_columns(self):
        return self.df.shape[1]

    def insert_empty_rows(self, how_many: int = 1):
        for _ in range(how_many):
            self.df.loc[len(self.df)] = None

    def insert_empty_columns(self, how_many: int = 1):
        num_rows = self.total_rows
        for _ in range(how_many):
            self.df.insert(self.total_columns, None, [None for _ in range(num_rows)])

    def format_row(self, row_idx: int, excel_format: ExcelFormat):
        for col_idx in range(self.total_columns):
            self.format_map.format_cell(row_idx=row_idx, col_idx=col_idx, excel_format=excel_format)

    def format_column(self, col_idx: int, excel_format: ExcelFormat):
        for row_idx in range(self.total_rows):
            self.format_map.format_cell(row_idx=row_idx, col_idx=col_idx, excel_format=excel_format)


class DataSheet:

    def __init__(self):
        self._tables: list[DataTable] = []

    @property
    def shape(self) -> tuple:
        rows = 0
        cols = 0

        for (row_start_idx, col_start_idx), table in self._tables:
            rows = max(
                rows,
                row_start_idx + table.total_rows
            )
            cols = max(
                cols,
                col_start_idx + table.total_columns
            )

        return rows, cols

    def insert_data_table(self,
                          data_table: DataTable,
                          row_start_idx: int,
                          col_start_idx: int):
        self._tables.append(
            ((row_start_idx, col_start_idx), data_table)
        )


    def create_dataframe(self) -> pd.DataFrame:
        canvas = np.empty(self.shape, dtype=object)

        for (row_start_idx, col_start_idx), table in self._tables:
            df = table.df
            table_values = np.vstack([df.columns.to_numpy(), df.to_numpy()])

            row_end_idx = row_start_idx + table_values.shape[0]
            col_end_idx = col_start_idx + table_values.shape[1]

            canvas[row_start_idx:row_end_idx, col_start_idx:col_end_idx] = table_values

        return pd.DataFrame(canvas)

    def create_format_map(self) -> CellFormatMap:
        master_map = CellFormatMap()
        for (row_start_idx, col_start_idx), table in self._tables:
            table_map = table.format_map
            for (row_idx, col_idx), excel_format in table_map._cell_map.items():
                master_map.format_cell(
                    row_idx=row_start_idx + row_idx,
                    col_idx=col_start_idx + col_idx,
                    excel_format=excel_format
                )
        return master_map

    def iter_tables(self):
        for (row_start_idx, col_start_idx), table in self._tables:
            yield row_start_idx, col_start_idx, table

class ExcelManager:

    def __init__(self,
                 output_path: Path | str):
        self.output_path = Path(output_path) if isinstance(output_path, str) else output_path
        if self.output_path.suffix != '.xlsx':
            raise ValueError(f"ExcelManager expects extension '.xlsx'. Received '{output_path.suffix}'")

        self._data_sheets = dict()

    def write_dataframe(self, df: pd.DataFrame, sheet_name: str):
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    def add_data_sheet(self, data_sheet: DataSheet, sheet_name: str):
        self._data_sheets[sheet_name] = data_sheet

    def export(self,
               autofit_column_widths: bool = True):
        formats_map = dict()
        with pd.ExcelWriter(self.output_path, engine='xlsxwriter') as writer:
            wb = writer.book

            for sheet_name, sheet_data in self._data_sheets.items():
                df = sheet_data.create_dataframe()
                df.to_excel(writer, sheet_name=sheet_name, index=False)

                ws = writer.sheets[sheet_name]
                format_map = sheet_data.create_format_map()

                for row_idx, col_idx, excel_format in format_map.iter_cells():
                    header_adjusted_row_idx = row_idx + 1
                    if excel_format not in formats_map:
                        formats_map[excel_format] = wb.add_format(excel_format.create_xlsx_writer_format())

                    excel_format = formats_map[excel_format]

                    value = df.iloc[row_idx, col_idx]
                    if pd.isna(value) or value in (float('inf'), float('-inf')):
                        value = None

                    # to_excel function writes a header in row 0 of the Excel file, so we have to shift our own
                    # data down 1 row to align with that
                    header_adjusted_row_idx = row_idx + 1
                    ws.write(header_adjusted_row_idx, col_idx, value, excel_format)

                if autofit_column_widths:
                    for i, col in enumerate(df.columns):
                        max_len = max(
                            df[col].astype(str).map(len).max(),
                            len(str(col))
                        )
                        ws.set_column(i, i, max_len + 4)


