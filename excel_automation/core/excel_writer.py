import os
import pandas as pd
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet
from .excel_formatter import ExcelFormatter
from ..utils import Formats
from typing import Tuple, Literal


class ExcelWriterXL:
    def __init__(self, df_list: list[pd.DataFrame], output_name: str, output_folder: str):
        """
        A class for writing multiple pandas DataFrames to Excel files with customized formatting.
        Uses xlsxwriter as the underlying engine and works with the ExcelFormatter class for styling.
        
        Parameters
        ----------
        df_list : list[pd.DataFrame]
            List of pandas DataFrames to be written to the Excel file
        output_name : str: 
            Name for the output file (extension already provided)
        output_folder : str:
            Folder path to save the file in.
            
        Attributes
        ----------
        writer : pd.ExcelWriter
            ExcelWriter object with xlsxwriter engine
        formatter : ExcelFormatter
            Formatter object that handles the styling of the Excel worksheets
        format : Formats
            Object containing format definitions
        df_list : list[pd.DataFrame]
            The list of DataFrames to be written
        """
        self.output_path = os.path.join(output_folder, f"{output_name}.xlsx")
        os.makedirs(os.path.dirname(self.output_path), exist_ok=True)
        self.writer = pd.ExcelWriter(self.output_path, engine='xlsxwriter')
        self.formatter = ExcelFormatter(df_list, self.writer)
        self.format = Formats()
        self.workbook: Workbook = self.writer.book
        self.sheet_list = []
        self.df_list = df_list
    
    def _ensure_worksheet_exists(self, sheet_name: str) -> Worksheet:
        if sheet_name in self.writer.sheets:
            worksheet: Worksheet = self.writer.sheets[sheet_name]
        else:
            # Crea una nueva hoja y la añade al diccionario de hojas
            worksheet: Worksheet = self.workbook.add_worksheet(sheet_name)
            self.writer.sheets[sheet_name] = worksheet
        return worksheet

    def write_from_df(
        self, 
        df: pd.DataFrame, 
        sheet_name: str, 
        num_format: str, 
        format_template: Literal["database", "index", "data_table", "text_table", "report"] | None = "database",
        highlighted_categories: str | list = "",
        **kwargs
    ) -> Tuple[pd.DataFrame, Worksheet]:
        """
        Write a DataFrame to a specific worksheet with the specified formatting template.
        
        Parameters
        ----------
        df : pd.DataFrame
            The DataFrame to write to the Excel worksheet
        sheet_name : str
            Name of the worksheet to write to
        num_format : str
            Number format string for numeric cells (e.g., "#,##0.00", "0.0%")
        format_template : Literal["database", "index", "data_table", "text_table"] | None, optional
            Template style to apply to the worksheet, defaults to "database"
            - "database": Standard format for database-like data
            - "index": Format with special handling for index columns
            - "data_table": Format optimized for numeric data tables
            - "text_table": Format optimized for text-heavy tables
            - None: No formatting applied, uses pandas default
        **kwargs : dict, optional
            Additional arguments for specific templates:
            - For "report": config_dict (dict): Report configuration settings
        
            
        Returns
        -------
        Tuple[pd.DataFrame, Worksheet]
            A tuple containing the written DataFrame and the xlsxwriter Worksheet object
        """
        worksheet = self._ensure_worksheet_exists(sheet_name)
        df = df.infer_objects(copy=False).fillna("")

        if format_template == "database":
            self.formatter.apply_database_format(worksheet, df, num_format)
        elif format_template == "data_table":
            self.formatter.apply_data_table_format(worksheet, df, num_format, highlighted_categories)
        elif format_template == "text_table":
            self.formatter.apply_text_table_format(worksheet, df, num_format)    
        elif format_template == "index":
            self.formatter.apply_index_format(worksheet, df, num_format)
        elif format_template == "index":
            self.formatter.apply_report_format(worksheet, df, num_format, **kwargs) 
        else:
            df.to_excel(self.writer, sheet_name=sheet_name, index=False)
        
        return df, worksheet
    

    def write_to_excel(self, sheet_name: str, row_num: int, column_num: int, value: str, header: bool = False) -> Worksheet:
        worksheet = self._ensure_worksheet_exists(sheet_name)
        if header:
            worksheet.write_string(row_num, column_num, value, cell_format=self.format.cells["report"]["header"])
        else:
            worksheet.write_string(row_num, column_num, value, cell_format=self.format.cells["report"]["data"])

        return worksheet


    def write_to_all_sheets(self, row_num: int, column_num: int, value: str, header: bool = False) -> None:
        for sheet_name in self.sheet_list:
            self.write_to_excel(sheet_name, row_num, column_num, value, header)
            
        
    def save_workbook(self):
        self.writer.close()
        print(f'✅ Excel guardado en "{self.output_path}"')
