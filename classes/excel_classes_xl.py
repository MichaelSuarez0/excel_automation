import re
import time
import os
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from openpyxl.styles.numbers import FORMAT_DATE_XLSX17
from enum import Enum
from icecream import ic
import logging
import datetime
import pandas as pd
from typing import Tuple, Optional
import xlsxwriter
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet
from xlsxwriter.format import Format
from xlsxwriter.utility import xl_range

# Hex colors
class Color(Enum):
    RED: str = "#d81326"
    RED_LIGHT: str = "#FFABAB"
    BLUE_LIGHT: str = "#A6CAEC"
    BLUE: str = "#2E70D0"
    BLUE_DARK: str = "#12213b"
    GREEN_DARK: str = "#008E2C"
    GRAY_LIGHT: str = "#ebebeb"
    GRAY: str = "#ebebe0"
    YELLOW: str = "#FFC000"
    WHITE: str = '#FFFFFF'
    ORANGE: str = "#DD6909"


script_dir = os.path.abspath(os.path.dirname(__file__))
macros_folder = os.path.join(script_dir, "..", "macros", "excel")
save_dir = os.path.join(script_dir, "..", "charts")

class ExcelReader:
    def __init__(self, file_name: str):
        self.file_name = os.path.join(script_dir, "..", "databases", f'{file_name}.xlsx')
        self.wb = None
        self.ws = None
        self.load_workbook()
    
    def load_workbook(self):
        self.wb = openpyxl.load_workbook(self.file_name)
        # Access different worksheets with self.wb.sheetnames[int]
        self.ws= self.wb.active
    
    def save_workbook(self, name: str = "pyexcel")-> None:
        """Save your workbook. Automatically includes extension in the name if not declared.

        Args:
            name (str, optional): Choose a name for your Excel file. Defaults to "pyexcel".
        """
        if not name.endswith('xlsx'):
            name = f'{name}.xlsx'
        self.wb.save(os.path.join(save_dir, name))
        print(f'✅ Excel guardado como "{name}"')
    
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
    
    def worksheet_to_dataframe(self, sheet_number: int = None) -> pd.DataFrame:
        sheet_name = self.wb.sheetnames[sheet_number] if sheet_number else self.wb.sheetnames[0]
        df = pd.read_excel(self.file_name, sheet_name)
        return df
    
    def worksheets_to_dataframes(self, include_first = False) -> list[pd.DataFrame]:
        dfs_dict = pd.read_excel(self.file_name, sheet_name=None) # Read all sheets at once into a dictionary of DataFrames
        sheet_names = list(dfs_dict.keys())[1:] if not include_first else list(dfs_dict.keys())
        dfs = [dfs_dict[name] for name in sheet_names]
        return dfs
        
    def dataframe_to_excel(self, df: pd.DataFrame, sheet_name='Hoja1', mode='w'):
        with pd.ExcelWriter(self.file_name, engine='openpyxl', mode=mode) as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False) # Pandas automatically saves
    
    def normalize_orientation(self, dfs: pd.DataFrame | list[pd.DataFrame]) -> list[pd.DataFrame]:
        """Normalizes the orientation of all DataFrames. Converts to list first if a single df is provided"""
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
   
    # TODO: All that is missing is FUENTE and URL
    # TODO: Use sheet index instead of name
    def apply_database_format(self, sheet_name='Hoja1', decimals = True)-> None:
        self.ws.column_dimensions['A'].width = 15
        self.ws.sheet_view.showGridLines = False
        w = Color.WHITE.value
        g = Color.GRAY_LIGHT.value

        # Data iteration: number formatting to cells containing data
        vanilla_border = Border(left=Side(style='thin', color=g), right=Side(style='thin', color=g), top=Side(style='thin', color=g), bottom=Side(style='thin', color=g))
        for row in self.ws.iter_rows(min_row=2, max_row=self.ws.max_row, min_col=2, max_col=self.ws.max_column):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "0.00" if decimals else "0"
                    cell.border = vanilla_border

        # First column (years)
        fill = PatternFill(start_color=Color.GRAY_LIGHT.value, fill_type="solid")
        custom_border = Border(left=Side(style='thin', color=w), right=Side(style='thin', color=w), top=Side(style='thin', color=w), bottom=Side(style='thin', color=w))
        for row in self.ws.iter_rows(min_row=1, max_row=self.ws.max_row, min_col=1, max_col=1):
            for cell in row:
                cell.fill = fill
                cell.border = custom_border
                # Check if the cell contains a datetime value
                if isinstance(cell.value, (datetime.date, datetime.datetime)):
                    cell.number_format = FORMAT_DATE_XLSX17  # Apply mmm-yy format

        # First row (categories)
        fill = PatternFill(start_color=Color.BLUE_DARK.value, fill_type="solid")
        for row in self.ws.iter_rows(min_row=1, max_row=1, min_col=1, max_col=self.ws.max_column):
            for cell in row:
                cell.fill = fill
                cell.font = Font(color=w, bold=True)
                cell.border = custom_border
                cell.alignment = Alignment(horizontal= "center", vertical="center")  # Ajustar texto. TODO: no funciona en excel en línea

# TODO: Handle orientation appropiately
# TODO: Set axis max and min range dynamically
# TODO: Processes like iterating over selected rows can be modularized further    
# TODO: Parámetro para especificar si leyenda o no, si no, que se haga más largo el cuadro del gráfico (porque ya no habría leyenda)


class ExcelAutoChart:
    def __init__(self, df_list: list[pd.DataFrame], output_name: str = "ExcelAutoChart"):
        self.writer = pd.ExcelWriter(os.path.join(save_dir, f'{output_name}.xlsx'), engine='xlsxwriter')
        self.workbook: Workbook = self.writer.book
        self.df_list = df_list
        self._initialize_chart_formats()

    def _initialize_chart_formats(self):
        """Predefine chart-specific formats using the Color enum"""
        color_list = [Color.RED.value, Color.BLUE.value, Color.GREEN_DARK.value, Color.ORANGE.value]
        self.chart_formats = {
            'line': {
                'colors': color_list,
                'width': 2.5,
                'dash_types': ['solid', 'dash', 'dot', 'solid']
            },
            'marker': {
                'size': 6,
                'colors': color_list
            }
        }

    def prepare_chart_data(
        self,
        df: pd.DataFrame,
        selected_labels: Optional[list[str]] = None,
        sheet_name: str = "ChartData"
    ) -> Tuple[pd.DataFrame, Worksheet]:
        """Process source data and prepare worksheet for charting"""
        if selected_labels:
            cols = [df.columns[0]] # Start with the first column
            
            # Loop through the remaining columns and add them if they are in selected_labels
            for col in df.columns[1:]:
                if col in selected_labels:
                    cols.append(col)
            filtered_df = df[cols]
        else:
            filtered_df = df
        #ic(filtered_df)

        # Write the filtered DataFrame to Excel
        filtered_df.to_excel(self.writer, sheet_name=sheet_name, index=False)

        # Get the worksheet and close the writer
        worksheet = self.writer.sheets[sheet_name]
        return filtered_df, worksheet

    def create_line_chart(
        self,
        index: int = 0,
        selected_labels: Optional[list[str]] = None,
        sheet_name: str = "LineChart",
        marker: bool = True
    ) -> Worksheet:
        """Generate line chart with color scheme from Color enum"""
        data_df, worksheet = self.prepare_chart_data(self.df_list[index], selected_labels, sheet_name)
        
        # Check if the DataFrame is empty
        if data_df.empty:
            raise ValueError("DataFrame is empty. No data to plot.")

        chart = self.workbook.add_chart({'type': 'line'})
        
        # Configure chart appearance
        chart.set_title({'name': 'Performance Over Time', 'name_font': {'size': 14}})
        chart.set_size({'width': 600, 'height': 400})
        chart.set_legend({'position': 'bottom'})

        # Add data series with color scheme
        for idx, col in enumerate(data_df.columns[1:]):
            col_letter = chr(66 + idx)  # Get column letter (e.g., B, C, D, ...)
            #print(f"Adding series for column {col}: {col_letter}")  # Debug

            series_params = {
                'name': f"={sheet_name}!${col_letter}$1",  # Use column letter dynamically
                'categories': f"={sheet_name}!$A$2:$A${len(data_df)+1}",
                'values': f"={sheet_name}!${col_letter}$2:${col_letter}${len(data_df)+1}",
                'line': {
                    'color': self.chart_formats['line']['colors'][idx % len(self.chart_formats['line']['colors'])],
                    'width': self.chart_formats['line']['width'],
                    'dash_type': self.chart_formats['line']['dash_types'][idx % len(self.chart_formats['line']['dash_types'])],
                }
            }

            if marker:
                marker_color = self.chart_formats['marker']['colors'][idx % len(self.chart_formats['marker']['colors'])]
                series_params['marker'] = {
                    'type': 'circle',
                    'size': self.chart_formats['marker']['size'],
                    'fill': {'color': marker_color},
                    'line': {'color': marker_color},
                }

            chart.add_series(series_params)


        # Axis configuration
        chart.set_y_axis({
            'name': 'Percentage (%)',
            'num_format': '0.00',
            'major_gridlines': {
                'visible': True,
                'line': {'color': Color.GRAY_LIGHT.value}
            }
        })
        
        chart.set_x_axis({
            'name': '',
            'num_format': '0',
            'text_axis': True,
        })

        # Insert chart with proper positioning
        worksheet.insert_chart('E2', chart, {'x_offset': 25, 'y_offset': 10})
        
        self.writer.close()  # Make sure to save the workbook
        return worksheet


