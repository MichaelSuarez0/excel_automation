import re
import os
import pandas as pd
import openpyxl
from icecream import ic
import logging
import datetime
import pandas as pd
from typing import Tuple, Optional


script_dir = os.path.abspath(os.path.dirname(__file__))
#macros_folder = os.path.join(script_dir, "..", "macros", "excel")
save_dir = os.path.join(script_dir, "..", "charts")

class ExcelDataExtractor():
    def __init__(self, file_name: str, output_name: str):
        """Class to obtain data from an Excel file, convert to DataFrame, apply transformations, and export it. 
        Engine: mostly pandas

        Parameters
        ----------
        file_name : str
                The name of the Excel file to be loaded (from databases folder)
        """
        self.file_path = os.path.join(script_dir, "..", "databases", f'{file_name}.xlsx')
        self.output_path = os.path.join(script_dir, "..", "charts", f'{output_name}.xlsx')
        self.wb = None
        self.ws = None
        self.load_workbook()
    
    def load_workbook(self):
        self.wb = openpyxl.load_workbook(self.file_path)
        # Access different worksheets with self.wb.sheetnames[int]
        self.ws= self.wb.active
    
    def save_workbook(self)-> None:
        """Save your workbook. Automatically includes extension in the name if not declared.

        Args:
            name (str, optional): Choose a name for your Excel file. Defaults to "excel_test".
        """
        self.wb.save(self.output_path)
        print(f'✅ Excel guardado como "{self.output_path}"')

    # def open_new_workbook(self, ws_name: str = None) -> Tuple[Workbook, Worksheet]:
    #     """Dynamically create new workbooks and name them wb2, wb3, etc."""        
    #     self.wb_count += 1  
        
    #     # Create new workbook and assign it dynamically
    #     new_wb_name = f"wb{self.wb_count}"
    #     self.workbooks[(self.wb_count)] = Workbook()
        
    #     # Create new variables dynamically (starting with .self)
    #     setattr(self, new_wb_name, self.workbooks[self.wb_count])
    #     setattr(self, f"ws{self.wb_count}", self.workbooks[self.wb_count].active)
    #     # Get the active worksheet or create a new one with the specified name
    #     if ws_name:
    #         new_ws = self.workbooks[self.wb_count].create_sheet(title=ws_name)
    #     else:
    #         new_ws = self.workbooks[self.wb_count].active

    #     print(f"✅ Created new workbook: {new_wb_name}")
    #     return self.workbooks[self.wb_count], new_ws
    
    @property
    def sheet_names(self) -> list: 
        """Devuelve una lista de los nombres de las hojas."""
        print("Sheet names:")
        for sheet_name in self.wb.sheetnames:
            print(f"- {sheet_name}")
        return self.wb.sheetnames

    @property
    def count_sheets(self) -> int:
        count = len(self.wb.sheetnames)
        print(f'The workbook has {count} sheets.')
        return count
    
    # Opening methods
    def worksheet_to_dataframe(self, sheet_index: int = None) -> pd.DataFrame:
        """Reads sheet and return a DataFrame, may specify worksheet index"""
        sheet_name = self.wb.sheetnames[sheet_index] if sheet_index else self.wb.sheetnames[0]
        df = pd.read_excel(self.file_path, sheet_name)
        return df
    
    def worksheets_to_dataframes(self, include_first = False) -> list[pd.DataFrame]:
        """Reads all sheets at once and returns a list of DataFrames, may specify to skip first"""
        dfs_dict = pd.read_excel(self.file_path, sheet_name=None) # This method reads all sheets at once a returns a dictionary of DataFrames
        sheet_names = list(dfs_dict.keys())[1:] if not include_first else list(dfs_dict.keys())
        dfs = [dfs_dict[name] for name in sheet_names]
        return dfs
    
    # Transformation methods
    def normalize_orientation(self, dfs: pd.DataFrame | list[pd.DataFrame]) -> list[pd.DataFrame]:
        """Normalizes the orientation of all DataFrames. Converts to list if a single df is provided"""
        if not isinstance(dfs, (pd.DataFrame, list)):
            raise ValueError("Must provide either a DataFrame or a list of DataFrames")
        if isinstance(dfs, pd.DataFrame):
            dfs= [dfs]
        normalized_dfs = []
        for df in dfs:
            # Check if the first row contains categories. If it does, it will transpose the df.
            if isinstance(df.iloc[0, 1], str) and isinstance(df.iloc[1, 0], str):
                continue
            if not isinstance(df.iloc[0, 1], str):
                index_name = df.columns[0]
                df = df.set_index(df.columns[0]).transpose() # Manually set another index, or else the default index stays on top
                df.reset_index(inplace=True)
                df.columns = [index_name] + df.columns[1:].tolist()  # I loathe pandas indexes
            normalized_dfs.append(df)
        
        return normalized_dfs
    
    def filter_data(
        self,
        df: pd.DataFrame,
        selected_categories: Optional[list[str]] = None,
    ) -> pd.DataFrame:
        """Filters data based on selected_categories"""
        if selected_categories:
            cols = [df.columns[0]] # Start with the first column, remember labels are in Row 1
            
            # Loop through the remaining columns and add them if they are in selected_labels
            for col in df.columns[1:]:
                if col in selected_categories:
                    cols.append(col)
            filtered_df = df[cols]
        else:
            filtered_df = df
        #ic(filtered_df)

        return filtered_df
    
    # Writing methods (simple)
    def dataframe_to_worksheet(self, df: pd.DataFrame, sheet_name: str = 'Hoja1', mode: str = 'w') -> None:
        """Writes a DataFrame to a worksheet in the Excel file.

        Parameters
        ----------
        df : pd.DataFrame
            The DataFrame to write to the worksheet.
        sheet_name : str, optional
            The name of the worksheet. Defaults to 'Hoja1'.
        mode : str, optional
            The mode to open the Excel file ('w' for write, 'a' for append). Defaults to 'w'.
        """
        with pd.ExcelWriter(self.output_path, engine='openpyxl', mode=mode) as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
    def dataframes_to_worksheets(self, dfs: list[pd.DataFrame], sheet_names: list[str] = None, mode: str = 'w', skip_first: bool = True) -> None:
        """Writes multiple DataFrames to multiple worksheets in the Excel file.

        Parameters
        ----------
        dfs : List[pd.DataFrame]
            A list of DataFrames to write to the worksheets.
        sheet_names : list[str], optional
            A list of worksheet names. If not provided, default names will be used.
        mode : str, optional
            The mode to open the Excel file ('w' for write, 'a' for append). Defaults to 'w'.
        skip_first : bool, optional
            Whether to start writing from Worksheet 2 onward. Defaults to True.
        """
        if sheet_names is None:
            sheet_names = [f'Hoja{i+1}' for i in range(len(dfs))]  # Default sheet names: Hoja1, Hoja2, etc.

        if len(dfs) != len(sheet_names):
            raise ValueError("The number of DataFrames must match the number of sheet names.")

        # If skip_first is True, add a blank worksheet as the first one
        if skip_first:
            with pd.ExcelWriter(self.output_path, engine='openpyxl', mode=mode) as writer:
                pd.DataFrame().to_excel(writer, sheet_name='Índice') 

        # Write DataFrames to subsequent sheets
        for i, (df, sheet_name) in enumerate(zip(dfs, sheet_names), start=1 if skip_first else 0):
            self.dataframe_to_worksheet(df, sheet_name=sheet_name, mode=mode)
   
    # TODO: All that is missing is FUENTE and URL
    # TODO: Use sheet index instead of name




# class ExcelAutomation:
#     def __init__(self, file_name: str):
#         """Class for automating Excel-related tasks.

#         Parameters
#         ----------
#         file_name : str
#             The name of the Excel file to be created or loaded.
#         """
#         self.handler = ExcelHandler(file_name)  # Initialize ExcelHandler
#         self.formatter = ExcelFormatter(workbook= self.handler.wb)  # Pass the workbook to ExcelFormatter

#     def save_workbook(self, name: str = "excel_test") -> None:
#         """Saves the workbook using ExcelHandler."""
#         self.handler.save_workbook(name)

#     def apply_database_format(self, sheet_name: str = 'Hoja1', decimals: bool = True) -> None:
#         """Applies database formatting using ExcelFormatter."""
#         self.formatter.apply_database_format(sheet_name, decimals)

#     def get_sheet_names(self) -> list[str]:
#         """Returns the sheet names using ExcelHandler."""
#         return self.handler.sheet_names

#     def get_count_sheets(self) -> int:
#         """Returns the number of sheets using ExcelHandler."""
#         return self.handler.count_sheets
