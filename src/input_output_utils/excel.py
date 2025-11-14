
import copy
import os
import xlsxwriter
from typing import Union
from enum import Enum
from pathlib import Path
import numpy as np

from .enums import HexColor
import pandas as pd


class ExcelFormat:

    def __init__(self,
                 font_color: HexColor | None = None,
                 font_size: int | None = 14,
                 is_bold: bool | None = False,
                 fill_color: HexColor | None= None,
                 text_align: str | None = 'left'):
        self.font_color = font_color
        self.font_size = font_size
        self.is_bold = is_bold
        self.fill_color = fill_color
        self.text_align = text_align

    def create_xlsx_writer_format(self) -> dict:
        f = {
            "font_size": self.font_size,
            "bold": self.is_bold,
            "align": self.text_align
        }

        if self.font_color:
            f.update({"font_color": self.font_color.value})

        if self.fill_color:
            f.update({"bg_color": self.fill_color.value})

        return f

    def __hash__(self):
        return hash(
            (self.font_color, self.font_size, self.is_bold, self.fill_color)
        )

    def update(self, other: "ExcelFormat"):
        if not isinstance(other, ExcelFormat):
            raise TypeError(f"Type {type(other)} not applicable to function update of ExcelFormat.")

        for k, v in other.__dict__.items():
            if v is not None:
                setattr(self, k, v)

class CellFormatMap:

    def __init__(self):
        self._cell_map = dict()

    def format_cell(self, row_idx: int, col_idx: int, excel_format: ExcelFormat, exist_ok=False):
        format_copy = copy.deepcopy(excel_format)

        loc = (row_idx, col_idx)

        if loc not in self._cell_map:
            self._cell_map[loc] = format_copy
            return

        if exist_ok:
            existing_format = self._cell_map[loc]
            existing_format.update(format_copy)
        else:
            raise KeyError(f"Cell at ({row_idx}, {col_idx}) already exists in {self.__class__.__name__}.")

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

    def format_cell(self, row_idx: int, col_idx: int, excel_format: ExcelFormat, exist_ok=False):
        self.format_map.format_cell(row_idx=row_idx,
                                    col_idx=col_idx,
                                    excel_format=excel_format,
                                    exist_ok=exist_ok)

    def format_rows(self, row_indices: list[int], excel_format: ExcelFormat, exist_ok=False):
        for row_idx in row_indices:
            for col_idx in range(self.total_columns):
                self.format_map.format_cell(row_idx=row_idx,
                                            col_idx=col_idx,
                                            excel_format=excel_format,
                                            exist_ok=exist_ok)

    def format_columns(self, col_indices: list[int], excel_format: ExcelFormat, exist_ok=False):
        for col_idx in col_indices:
            for row_idx in range(self.total_rows):
                self.format_map.format_cell(row_idx=row_idx,
                                            col_idx=col_idx,
                                            excel_format=excel_format,
                                            exist_ok=exist_ok)


class DataSheet:

    def __init__(self):
        self._grid = []
        self.master_format_map = CellFormatMap()

        self._num_rows = 0
        self._num_cols = 0

    @property
    def shape(self) -> tuple:
        return self._num_rows, self._num_cols

    def _ensure_size(self, required_num_rows: int, required_num_cols: int):
        rows = max(self._num_rows, required_num_rows)
        cols = max(self._num_cols, required_num_cols)

        # Ensure existing rows have enough columns
        if cols > self._num_cols:
            col_diff = cols - self._num_cols
            for row in self._grid:
                row.extend([''] * col_diff)

        # Ensure enough rows exist
        if rows > self._num_rows:
            row_diff = rows - self._num_rows
            for _ in range(row_diff):
                self._grid.append([''] * cols)

        self._num_rows = rows
        self._num_cols = cols

    def insert_data_table(self,
                          data_table: DataTable,
                          row_start_idx: int,
                          col_start_idx: int):
        df = data_table.df
        table_values = np.vstack([df.columns.to_numpy(), df.to_numpy()])

        num_rows, num_cols = table_values.shape

        row_end_idx = row_start_idx + table_values.shape[0]
        col_end_idx = col_start_idx + table_values.shape[1]

        self._ensure_size(
            required_num_rows=row_end_idx,
            required_num_cols=col_end_idx
        )
        for row_idx in range(num_rows):
            for col_idx in range(num_cols):
                self._grid[row_start_idx + row_idx][col_start_idx + col_idx] = table_values[row_idx][col_idx]

        for row_idx, col_idx, excel_format in data_table.format_map.iter_cells():
            self.master_format_map.format_cell(
                row_idx=row_start_idx + row_idx,
                col_idx=col_start_idx + col_idx,
                excel_format=excel_format
            )

    def create_dataframe(self) -> pd.DataFrame:
        return pd.DataFrame(self._grid)

class ExcelManager:

    def __init__(self,
                 output_path: Path | str):
        self.output_path = Path(output_path) if isinstance(output_path, str) else output_path
        if self.output_path.suffix != '.xlsx':
            raise ValueError(f"ExcelManager expects extension '.xlsx'. Received '{output_path.suffix}'")

        self._data_sheets = dict()

    def add_data_sheet(self, sheet_name: str, data_sheet: DataSheet):
        if sheet_name in self._data_sheets:
            raise KeyError(f"Sheet {sheet_name} already in data sheets dict.")

        self._data_sheets[sheet_name] = data_sheet

    def write_dataframe(self, df: pd.DataFrame, sheet_name: str):
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    def export(self,
               autofit_column_widths: bool = True):
        formats_map = dict()
        with pd.ExcelWriter(self.output_path, engine='xlsxwriter') as writer:
            wb = writer.book

            for sheet_name, data_sheet in self._data_sheets.items():
                df = data_sheet.create_dataframe()
                df.to_excel(writer, sheet_name=sheet_name, index=False)

                ws = writer.sheets[sheet_name]

                for row_idx, col_idx, excel_format in data_sheet.master_format_map.iter_cells():
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
                        ws.set_column(i, i, max_len + 8)


