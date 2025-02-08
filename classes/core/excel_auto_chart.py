import os
import pandas as pd
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet
from xlsxwriter.format import Format
from xlsxwriter.utility import xl_range
from excel_automation.classes.formats.colors import Color
from excel_automation.classes.formats.formats import ExcelFormats
from typing import Tuple, Optional, Literal
import numpy as np

script_dir = os.path.abspath(os.path.dirname(__file__))
save_dir = os.path.join(script_dir, "..", "..", "charts")


class ExcelAutoChart:
    def __init__(self, df_list: list[pd.DataFrame], output_name: str = "ExcelAutoChart"):
        """Class to write to Excel files from DataFrames and creating charts. Engine: xlsxwriter

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
        format = ExcelFormats()
        self.format_cells = format.format_cells
        self.format_charts = format.format_charts
        self.numeric_types = format.numeric_types
        
        
    def _apply_formatting_to_worksheet(self, worksheet: Worksheet, df: pd.DataFrame, num_format: str):
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
                
    def _write_to_excel(self, df: pd.DataFrame, num_format:str, sheet_name: str = "ChartData", apply_format = True) -> Tuple[pd.DataFrame, Worksheet]:
        df.to_excel(self.writer, sheet_name=sheet_name, index=False)
        worksheet = self.writer.sheets[sheet_name]
        self.sheet_dfs[sheet_name] = df
        if apply_format:
            self._apply_formatting_to_worksheet(worksheet, df, num_format)
        return df, worksheet
    
    # TODO: Add in formats
    # TODO: Here you can define base chart configs for bar charts
    # TODO: Consider discussing chart font being Aptos Narrow
    def _create_base_chart(self, worksheet: Worksheet, chart_type: str, chart_subtype: str = ""):
        """Default settings for all chart types"""
        chart = self.workbook.add_chart({'type': chart_type}) if not chart_subtype else self.workbook.add_chart({'type': chart_type, 'subtype': chart_subtype})

        chart.set_title({'name': ''})
        chart.set_size({'width': 600, 'height': 420})
        chart.set_legend({'position': 'bottom'})
        chart.set_plotarea({
            'layout': {
                'x':      0.11,
                'y':      0.09,
                'width':  0.83,
                'height': 0.75,
            }
        })
        chart.set_chartarea({'border': {'none': True}})

        return chart
 
    # TODO: Add axis in formats
    # TODO: Implement manual logic for specific series (i.e. Peru series) if column.name == Peru
    def create_line_chart(
            self,
            index: int = 0,
            sheet_name: str = "LineChart",
            markers_add: bool = True,
            numeric_type: Literal['integer', 'decimal_1', 'decimal_2', 'percentage'] = "decimal_2",
            axis_title: str = ""
        ) -> Worksheet:
        """
        Creates and inserts a line chart into an Excel worksheet using data from a DataFrame.

        Parameters
        -------
        index : int, optional
            Index of the DataFrame in df_list to use (default 0).
        sheet_name : str, optional
            Name of the worksheet (default is "FigX").
        markers_add : bool, optional
            Whether to add markers for series (default True).
        numeric_type : str, optional
            Defines the number format for the series. Options are:
            'integer', 'decimal_1', 'decimal_2', 'percentage'. (default is 'decimal_2')

        Returns
        -------
        Worksheet
            The worksheet with the inserted chart.
        """
        # Definir el formato numérico según 'numeric_type'
        num_format: str = self.numeric_types[numeric_type]
        data_df, worksheet = self._write_to_excel(self.df_list[index], num_format, sheet_name)

        # Check if the DataFrame is empty
        if data_df.empty:
            raise ValueError("DataFrame is empty. No data to plot.")
        
        chart = self._create_base_chart(worksheet, 'line')

        if len(data_df.columns) < 5:
            format_line = self.format_charts["line_simple"]
            colors = self.format_charts["line_simple"].get("colors", {})
        else:
            format_line = self.format_charts["line"]
            colors = self.format_charts["line"].get("colors", {})

        # Add data series with color scheme
        for idx, col in enumerate(data_df.columns[1:]):
            col_letter = chr(66 + idx)  # Get column letter (e.g., B, C, D, ...)
            #print(f"Adding series for column {col}: {col_letter}")  # Debug

            series_params = {
                **format_line,
                'name': f"={sheet_name}!${col_letter}$1",  # Use column letter dynamically
                'categories': f"={sheet_name}!$A$2:$A${len(data_df)+1}",
                'values': f"={sheet_name}!${col_letter}$2:${col_letter}${len(data_df)+1}",
                'fill': {'color': colors[(idx-1) % len(colors)]},
                'data_labels': {
                    'value': True if idx in (1,2) else False,
                    'position': 'above' if idx == 1 else 'below',
                    'num_format': num_format,
                    'font':{
                            'color': colors[(idx-1) % len(colors)],
                            }
                        },
            }

            if markers_add:
                series_params['marker'] = {
                    **self.format_charts["marker"]
                }

            chart.add_series(series_params)

        # Axis configuration
        chart.set_y_axis({
            **self.format_charts["y_axis"],
            'name': axis_title,
            'num_format': '0%' if num_format=='0.0%' else num_format,
        })
        
        chart.set_x_axis({
            **self.format_charts["x_axis"],
            'num_format': '0',
        })

        # Insert chart with proper positioning
        position = 'E3' if len(data_df.columns[1:]) < 4 else 'J3'
        worksheet.insert_chart(position, chart, {'x_offset': 25, 'y_offset': 10})
        
        #self.writer.close()  # Automatically saves
        print(f"✅ Gráfico de líneas agregado a la hoja {index + 1}")
        return worksheet
    
    # TODO: Ordenar por secciones
    # TODO: If bar, sort ascending     
    def create_bar_chart(
            self,
            index: int = 0,
            sheet_name: str = "FigX",
            grouping: Literal['standard', 'stacked', 'percentStacked'] = "standard",
            chart_type: Literal['bar', 'column'] = 'column',
            numeric_type: Literal['decimal_1', 'decimal_2', 'integer', 'percentage'] = "decimal_1",
            axis_title: str = ""
        ) -> Worksheet:
        """Generate a bar or column chart in Excel from data in a DataFrame list.

        Parameters
        ----------
        index : int, optional
            Index of the DataFrame in df_list to use (default is 0).
        sheet_name : str, optional
            Name of the worksheet (default is "FigX").
        grouping : str, optional
            Grouping type: "standard", "stacked", or "percentStacked" (default is "standard").
        chart_type : str, optional
            Type of chart to create: "column" for vertical or "bar" for horizontal (default is "column").
        numeric_type : str, optional
            Defines the number format for the series. Options are:
            'integer', 'decimal_1', 'decimal_2', 'percentage'. (default is 'decimal_2')

        Returns
        -------
        Worksheet
            The worksheet with the inserted chart.

        Raises
        ------
        ValueError
            If the DataFrame is empty or if an invalid chart_type is provided.
        """
        num_format = self.numeric_types[numeric_type]
        data_df, worksheet = self._write_to_excel(self.df_list[index], num_format, sheet_name)
        
        # Check if DataFrame is empty
        if data_df.empty:
            raise ValueError("DataFrame is empty. No data to plot.")

        # Map grouping types to xlsxwriter subtypes
        subtype_map = {
            'standard': 'clustered',
            'stacked': 'stacked',
            'percentStacked': 'percent_stacked'
        }
        subtype = subtype_map.get(grouping, 'clustered')

        # Validate chart type
        if chart_type not in {"column", "bar"}:
            raise ValueError("Invalid chart_type. Use 'column' (vertical) or 'bar' (horizontal).")

        # Predefined formats
        chart = self._create_base_chart(worksheet, chart_type, subtype)

        if grouping == "standard" or chart_type == "bar" or len(data_df.columns) < 4:
            format_column = self.format_charts["column_simple"]
            colors = self.format_charts["column_simple"].get("colors", {})
        else: 
            format_column = self.format_charts["column"]
            colors = self.format_charts["column"].get("colors", {})

        # Add data series with color scheme
        if chart_type == "column":
            for idx, col in enumerate(data_df.columns[1:]): # Saltamos la primera columna (categorías), recorre las columnas
                col_idx = idx + 1  
                value_data = (data_df[col] != 0).all()
                
                series_params = {
                    **format_column,
                    'name': [sheet_name, 0, col_idx],
                    'categories': [sheet_name, 1, 0, len(data_df), 0],  # Categorías en la primera columna 
                    'fill': {'color': colors[(idx-1) % len(colors)]},
                    'values': [sheet_name, 1, col_idx, len(data_df), col_idx],  
                    'data_labels': {**self.format_charts["bar"].get("data_labels", {}), 'num_format': num_format, 'value': value_data},
                }
                chart.add_series(series_params)
                
        elif chart_type == "bar":
            for row_idx in range(1, len(data_df) + 1):  
                print(colors[row_idx % len(colors)])

                series_params = {
                    **self.format_charts["bar"],
                    'name': [sheet_name, row_idx, 0],  
                    'categories': [sheet_name, 0, 1, 0, data_df.shape[1] - 1],  # Categorías en la primera fila
                    'fill': {'color': colors[(row_idx-1) % len(colors)]},
                    'values': [sheet_name, row_idx, 1, row_idx, data_df.shape[1] - 1], 
                    'data_labels': {**self.format_charts["bar"].get("data_labels", {}), 'num_format': num_format},
                }
                chart.add_series(series_params)
            
        # TODO: Move to initialize
        # Configure axes
        if chart_type == "column":
            chart.set_y_axis({
                **self.format_charts["y_axis"],
                'name': axis_title,
                'num_format': '0%' if num_format=='0.0%' else num_format,
                'min': 0,
            })
            chart.set_x_axis({
                **self.format_charts["x_axis"],
                'num_format': '@',
                })

        # TODO: Config properly
        elif chart_type == "bar":
            chart.set_legend({'none': True})
            chart.set_x_axis({
                'name': axis_title,
                'num_format': '0%' if num_format=='0.0%' else num_format,
                'min': 0,
                'minor_tick_mark': 'outside',
                'major_tick_mark': 'none',
                'major_gridlines': {'visible': True, 'line': {'color': Color.GRAY_LIGHT.value}}
            })
            chart.set_y_axis({
                **self.format_charts["x_axis"], # Inverted
                'num_format': '@',
                })

        # Insert chart with proper positioning
        position = 'E3' if len(data_df.columns[1:]) < 4 else 'J3'
        worksheet.insert_chart(position, chart, {'x_offset': 25, 'y_offset': 10})
        
        #self.writer.close()  # Save and close workbook
        print(f"✅ Gráfico de barras agregado a la hoja {index + 1}")
        return worksheet
    
    def create_table(
            self,
            index: int = 0,
            sheet_name: str = "TabX",
        ) -> Worksheet:
        """Generate a bar or column chart in Excel from a DataFrame.

        Parameters
        ----------
        index : int, optional
            Index of the DataFrame in df_list to use (default is 0).
        sheet_name : str, optional
            Name of the worksheet (default is "TabX").

        Returns
        -------
        Worksheet
            The worksheet with the inserted chart.

        """
        # Definir el formato numérico según 'numeric_type'
        data_df, worksheet = self._write_to_excel(self.df_list[index], "", sheet_name, apply_format=False)

        # Check if the DataFrame is empty
        if data_df.empty:
            raise ValueError("DataFrame is empty. No data to plot.")
        
        # Set column widths
        worksheet.set_column('A:A', 24)
        worksheet.set_column('B:B', 60)

        # Hide gridlines
        worksheet.hide_gridlines(2)

        # Write headers with header format
        for col_num, col_name in enumerate(data_df.columns):
            worksheet.write(0, col_num, col_name, self.workbook.add_format(self.format_cells['header']))

        
        # Modify base formats
        gray_format = {**self.format_cells['first_column'], 'valign': 'vcenter'}
        gray_bold_format = {**gray_format, 'bold': True}
        default_format = {**self.format_cells['data'], 'text_wrap': True, 'valign': 'vcenter'}
        bold_format = {**default_format, 'bold': True}
        

        # Write table contents with alternating colors and bold for first column
        for row_idx in range(data_df.shape[0]):
            for col_idx in range(data_df.shape[1]):
                cell_value = data_df.iloc[row_idx, col_idx]

                # Select format based on column and row index
                if col_idx == 0:
                    cell_format = self.workbook.add_format(gray_bold_format) if row_idx % 2 == 0 else self.workbook.add_format(bold_format)
                else:
                    cell_format = self.workbook.add_format(gray_format) if row_idx % 2 == 0 else self.workbook.add_format(default_format)

                worksheet.write(row_idx + 1, col_idx, cell_value, cell_format)

        print(f"✅ Tabla agregada en la hoja {index + 1}")
        return worksheet

    def save_workbook(self):
        self.writer.close()
        print(f'✅ Excel guardado como "{self.output_name}"')