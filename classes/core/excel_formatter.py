import os
import pandas as pd
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet
from excel_automation.classes.formats.colors import Color
from excel_automation.classes.formats.formats import Formats
from typing import Tuple, Literal
import numpy as np

script_dir = os.path.abspath(os.path.dirname(__file__))
save_dir = os.path.join(script_dir, "..", "..", "charts")


class ExcelFormatter:
    def __init__(self, df_list: list[pd.DataFrame], output_name: str = "ExcelAutoChart"):
        """Class to write to Excel files from DataFrames and applying format. Engine: xlsxwriter

        Parameters
        ----------
        df_list : list(pd.DataFrame):
            Data that will be written to Excel
        output_name : str, optional: 
            File name for the output file. Defaults to "ExcelAutoChart".
        """
        self.output_name = output_name
        self.writer = pd.ExcelWriter(os.path.join(save_dir, f'{output_name}.xlsx'), engine='xlsxwriter')
        self.workbook: Workbook = self.writer.book
        self.df_list = df_list
        self.sheet_dfs = {}
        self.format = Formats()
    
    def _write_to_excel(self, df: pd.DataFrame, sheet_name: str, num_format: str, format: Literal["database", "index"] | None = "database") -> Tuple[pd.DataFrame, Worksheet]:
        df.to_excel(self.writer, sheet_name=sheet_name, index=False)
        worksheet = self.writer.sheets[sheet_name]
        self.sheet_dfs[sheet_name] = df
        if format == "database":
            self.apply_database_format(worksheet, df, num_format)
        if format == "index":
            self.apply_index_format(worksheet, df, num_format)
        
        return df, worksheet

    def save_workbook(self):
        self.writer.close()
        print(f'✅ Excel guardado como "{self.output_name}"')
        

    def apply_database_format(self, worksheet: Worksheet, df: pd.DataFrame, num_format: str):
        """Applies formatting only to cells with data."""

        # Set column widths
        worksheet.set_column('A:A', 15)
        if len(df.columns) > 1:
            if len(str(df.iloc[0,1])) > 11:
                worksheet.set_column(1, len(df.columns) - 1, 14)
            else:
                worksheet.set_column(1, len(df.columns) - 1, 10)

        # Hide gridlines
        worksheet.hide_gridlines(2)

        # Determine format for the first column and adjust for datetime
        first_col = df.columns[0]
        first_col_fmt = self.workbook.add_format(self.format.cells['first_column'])
        if pd.api.types.is_datetime64_any_dtype(df[first_col]):
            first_col_fmt.set_num_format('mmm-yy')

        # Write headers with header format
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, self.workbook.add_format(self.format.cells["header"]))

        # Define format for numeric data columns once
        fmt = self.workbook.add_format(self.format.cells['data'])
        fmt.set_num_format(num_format)

        # Write data cells with appropriate formats
        for row_idx in range(df.shape[0]):
            # First column (e.g., dates or text)
            cell_value = df.iloc[row_idx, 0]
            worksheet.write(row_idx + 1, 0, cell_value, first_col_fmt)

            # Other columns (numeric data)
            for col_idx in range(1, df.shape[1]):
                cell_value = df.iloc[row_idx, col_idx]

                # Skip NaN/Inf values by checking if the value is NaN or Inf
                if pd.isna(cell_value) or (isinstance(cell_value, (int, float)) and np.isinf(cell_value)):
                    worksheet.write(row_idx + 1, col_idx, '')  # Write an empty cell
                else:
                    worksheet.write(row_idx + 1, col_idx, cell_value, fmt)

    def apply_index_format(self, worksheet: Worksheet, df: pd.DataFrame, num_format: str):
        """NOT IMPLEMENTED"""
        # Set column widths
        worksheet.set_column('A:A', 15)
        if len(df.columns) > 1:
            if len(str(df.iloc[0,1])) > 11:
                worksheet.set_column(1, len(df.columns) - 1, 14)
            else:
                worksheet.set_column(1, len(df.columns) - 1, 10)

        # Hide gridlines
        worksheet.hide_gridlines(2)

        # Determine format for the first column and adjust for datetime
        first_col = df.columns[0]
        first_col_fmt = self.workbook.add_format(self.format_cells['first_column'])
        if pd.api.types.is_datetime64_any_dtype(df[first_col]):
            first_col_fmt.set_num_format('mmm-yy')

        # Write headers with header format
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, self.workbook.add_format(self.format_cells["header"]))

        # Define format for numeric data columns once
        fmt = self.workbook.add_format(self.format_cells['data'])
        fmt.set_num_format(num_format)

        # Write data cells with appropriate formats
        for row_idx in range(df.shape[0]):
            # First column (e.g., dates or text)
            cell_value = df.iloc[row_idx, 0]
            worksheet.write(row_idx + 1, 0, cell_value, first_col_fmt)

            # Other columns (numeric data)
            for col_idx in range(1, df.shape[1]):
                cell_value = df.iloc[row_idx, col_idx]

                # Skip NaN/Inf values by checking if the value is NaN or Inf
                if pd.isna(cell_value) or np.isinf(cell_value):
                    worksheet.write(row_idx + 1, col_idx, '')  # Write an empty cell
                else:
                    worksheet.write(row_idx + 1, col_idx, cell_value, fmt)
    

