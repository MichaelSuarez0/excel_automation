import os
import pandas as pd
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet
from excel_automation.classes.utils.formats import Formats
from excel_automation.classes.core.excel_writer import ExcelWriterXL
from typing import Tuple, Literal
import numpy as np

script_dir = os.path.abspath(os.path.dirname(__file__))
save_dir = os.path.join(script_dir, "..", "..", "products")

# Ideally should only receive Wb, Ws and Formats
class ExcelFormatter:
    def __init__(self, df_list: list[pd.DataFrame], writer: pd.ExcelWriter):
        """
        Class for applying custom formatting to Excel worksheets created with pandas.
        Depends from ExcelWriterXL class.
        
        Parameters
        ----------
        df_list : list[pd.DataFrame]
            List of pandas DataFrames that will be formatted
        writer : pd.ExcelWriter
            ExcelWriter object with xlsxwriter engine
            
        Attributes
        ----------
        df_list : list[pd.DataFrame]
            The list of DataFrames to be formatted
        writer : pd.ExcelWriter
            The Excel writer object injected from ExcelWriterXL
        format : Formats
            Object containing predefined format configurations
        """
        self.df_list = df_list
        self.writer = writer
        self.format = Formats()
        

    def apply_database_format(self, worksheet: Worksheet, df: pd.DataFrame, num_format: str):
        """Applies formatting only to cells with data."""

        ### Widths and heights
        worksheet.set_column('A:A', 15)
        if len(df.columns) > 1:
            if len(str(df.columns[1])) > 11:
                worksheet.set_column(1, len(df.columns) - 1, 14)
            else:
                worksheet.set_column(1, len(df.columns) - 1, 10)

        ### Basic configurations
        fmt = self.format.cells['database']
        fmt['data']['num_format'] = num_format
        worksheet.hide_gridlines(2)

        ### Writing
        # Write data cells with appropriate formats
        for row_idx in range(df.shape[0]):
            cell_value = df.iloc[row_idx, 0]
            # First column (e.g., dates or text)
            if isinstance(cell_value, str) and ("/" in cell_value or "-" in cell_value):
                try: 
                    date_value = pd.to_datetime(cell_value)
                    excel_date = (date_value - pd.Timestamp("1899-12-30")).days
                    date_fmt = fmt['first_column']
                    date_fmt['num_format'] = 'mmm-yy'
                    worksheet.write(row_idx + 1, 0, excel_date, self.workbook.add_format(date_fmt))
                except ValueError:
                    worksheet.write(row_idx + 1, 0, cell_value, self.workbook.add_format(fmt['first_column']))
            else:
                worksheet.write(row_idx + 1, 0, cell_value, self.workbook.add_format(fmt['first_column']))

            # Rest of columns (numeric data)
            for col_idx in range(1, df.shape[1]):
                cell_value = df.iloc[row_idx, col_idx]

                # Skip NaN/Inf values by checking if the value is NaN or Inf
                if pd.isna(cell_value) or (isinstance(cell_value, (int, float)) and np.isinf(cell_value)):
                    worksheet.write(row_idx + 1, col_idx, '')  # Write an empty cell
                else:
                    worksheet.write(row_idx + 1, col_idx, cell_value, self.workbook.add_format(fmt['data']))
        
        # Headers
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, self.workbook.add_format(fmt["header"]))
    
     # TODO: Try if df.iloc[0,1] has a '-' 
    def apply_text_table_format(self, worksheet: Worksheet, df: pd.DataFrame, num_format: str):
        """Applies formatting only to cells with data."""

        ### Widths and heights
        worksheet.set_column('A:A', 26)
        worksheet.set_column('B:B', 54)
        worksheet.set_row(0, 20)

        ### Basic configurations
        worksheet.hide_gridlines(2)
        fmt = self.format.cells['text_table']

        ### Writing
        # Write headers with header format
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, self.workbook.add_format(fmt['header']))

        # Modify base formats
        gray_format = {**fmt['first_column']}
        gray_bold_format = {**gray_format, 'bold': True, 'align': 'left'}
        white_format = {**fmt['data'], 'right': 0}
        white_bold_format = {**white_format, 'bold': True, 'align': 'left'}
        
        # Write table contents with alternating colors and bold for first column
        for row_idx in range(df.shape[0]):
            for col_idx in range(df.shape[1]):
                cell_value = df.iloc[row_idx, col_idx]

                # Select format based on column and row index
                if col_idx == 0:
                    cell_format = self.workbook.add_format(gray_bold_format) if row_idx % 2 == 0 else self.workbook.add_format(white_bold_format)
                else:
                    cell_format = self.workbook.add_format(gray_format) if row_idx % 2 == 0 else self.workbook.add_format(white_format)

                worksheet.write(row_idx + 1, col_idx, cell_value, cell_format)
    
    # TODO: Set row heights dinamically
    def apply_data_table_format(self, worksheet: Worksheet, df: pd.DataFrame, num_format: str):
        """Applies formatting to data tables"""

        ### Widths and heights
        # First column width
        worksheet.set_column('A:A', 13)
        if len(df.columns) > 1:
            if len(str(df.iloc[0,1])) > 11:
                worksheet.set_column(1, len(df.columns) - 1, 14)
            else:
                worksheet.set_column(1, len(df.columns) - 1, 10)

        # Rest of columns widths
        for col_idx in range(1, df.shape[1]):
            worksheet.set_column(col_idx, col_idx, 5.15)

        # Row heights
        for row_idx in range(df.shape[0]+1):
            worksheet.set_row(row_idx, 15)

        ### Basic configurations
        worksheet.hide_gridlines(2)
        fmt = self.format.cells['data_table']
        fmt['data']['num_format'] = num_format

        ### Writing data
        for row_idx in range(df.shape[0]):
            cell_value = df.iloc[row_idx, 0]
            # First column (e.g., dates or text)
            worksheet.write(row_idx + 1, 0, cell_value, self.workbook.add_format(fmt['first_column']))

            # Rest of columns (numeric data)
            for col_idx in range(1, df.shape[1]):
                cell_value = df.iloc[row_idx, col_idx]

                # Skip NaN/Inf values by checking if the value is NaN or Inf
                if pd.isna(cell_value) or (isinstance(cell_value, (int, float)) and np.isinf(cell_value)):
                    worksheet.write(row_idx + 1, col_idx, '')  # Write an empty cell
                else:
                    worksheet.write(row_idx + 1, col_idx, cell_value, self.workbook.add_format(fmt['data']))
        
        # Headers
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, self.workbook.add_format(fmt["header"]))


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
    

