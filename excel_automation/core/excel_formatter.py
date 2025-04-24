import pandas as pd
from ..utils import Color, Formats
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet
import numpy as np
import copy


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
        self.workbook: Workbook = self.writer.book
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
            if isinstance(cell_value, pd.Timestamp): 
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
                fmt_modified = copy.deepcopy(fmt["data"])
                try:
                    if int(cell_value) > 9999:
                        fmt_modified['num_format'] = "# ### ##" + num_format
                except ValueError:
                    pass

                # Skip NaN/Inf values by checking if the value is NaN or Inf
                if pd.isna(cell_value) or (isinstance(cell_value, (int, float)) and np.isinf(cell_value)):
                    worksheet.write(row_idx + 1, col_idx, '')  # Write an empty cell
                else:
                    worksheet.write(row_idx + 1, col_idx, cell_value, self.workbook.add_format(fmt_modified))
        
        # Headers
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, self.workbook.add_format(fmt["header"]))
    
     # TODO: Try if df.iloc[0,1] has a '-' 
    def apply_text_table_format(self, worksheet: Worksheet, df: pd.DataFrame, num_format: str):
        """Applies formatting only to cells with data."""

        ### Widths and heights
        worksheet.set_column('A:A', 27)
        worksheet.set_column('B:B', 57)
        worksheet.set_row(0, 29)

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
    

    def apply_data_table_format(self, worksheet: Worksheet, df: pd.DataFrame, num_format: str, highlighted_categories: str | list = ""):
        """Applies formatting to data tables"""

        if isinstance(highlighted_categories, str):
            highlighted_categories = [highlighted_categories]
        ### Widths and heights
        # First column width
        worksheet.set_column('A:A', 12)  # Ancho fijo para nombres
        
        # Columnas restantes (años y Var (%))
        num_columns = df.shape[1] - 1 
        base_width = 8.0 if num_columns <= 5 else (6.0 if num_columns <= 8 else 5.2)
        
        handicap = 0 if num_format == 0 else (2.3 if num_format in ('0.0', '0,0%') else 3.3)
        for col_idx in range(1, df.shape[1]):
            # Longitud máxima considerando solo parte entera (evita decimales inflados)
            max_len = df.iloc[:, col_idx].astype(str).apply(
                lambda x: len(x.split('.')[0])
            ).max()
                
            dynamic_width = max(base_width, min(10, max_len + handicap))
            dynamic_width = round(float(dynamic_width), 2)  # Convertir a Python float
            
            # Aplicar ancho + formato numérico
            worksheet.set_column(col_idx, col_idx, dynamic_width)

        # Row heights
        for row_idx in range(1, df.shape[0]+1):
            if df.shape[0] > 10:
                worksheet.set_row(row_idx, 18)
            else:
                worksheet.set_row(row_idx, 26) # consider 30

        ### Basic configurations
        worksheet.hide_gridlines(2)
        fmt = self.format.cells['data_table']
        fmt['data']['num_format'] = num_format

        ### Writing data
        highlighted_row = False
        for row_idx in range(df.shape[0]):
            cell_value = df.iloc[row_idx, 0]

            # First column (e.g., dates or text)
            if highlighted_categories and cell_value in highlighted_categories:
                highlighted_row = True
                fmt_modified = copy.deepcopy(fmt["first_column"])
                fmt_modified['bg_color'] = Color.BLUE_LIGHT
                fmt_modified['bold'] = True,
                worksheet.write(row_idx + 1, 0, cell_value, self.workbook.add_format(fmt_modified))
            else:
                worksheet.write(row_idx + 1, 0, cell_value, self.workbook.add_format(fmt['first_column']))

            # Rest of columns (numeric data)
            for col_idx in range(1, df.shape[1]):
                cell_value = df.iloc[row_idx, col_idx]
                fmt_modified = copy.deepcopy(fmt["data"])
                try:
                    if int(cell_value) > 9999:
                        fmt_modified['num_format'] = "# ### ##" + num_format
                except ValueError:
                    pass

                if highlighted_row:
                    fmt_modified['bg_color'] = Color.BLUE_LIGHT
                    fmt_modified['bold'] = True
                
                # Skip NaN/Inf values by checking if the value is NaN or Inf
                if pd.isna(cell_value) or (isinstance(cell_value, (int, float)) and np.isinf(cell_value)):
                    worksheet.write(row_idx + 1, col_idx, '', self.workbook.add_format(fmt_modified))  # Write an empty cell
                else:
                    worksheet.write(row_idx + 1, col_idx, cell_value, self.workbook.add_format(fmt_modified))
            highlighted_row = False

        # Headers
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, self.workbook.add_format(fmt["header"]))


    def apply_index_format(self, worksheet: Worksheet, df: pd.DataFrame, num_format: str = ""):
        # Set column widths
        worksheet.set_column('A:A', 10)
        worksheet.set_column('B:C', 40)
        worksheet.set_column('D:D', 15)
        worksheet.set_column('E:G', 25)
        #worksheet.set_column('G:G', 15)

        ### Basic configurations
        worksheet.hide_gridlines(2)
        fmt = self.format.cells['index']

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
    
    

