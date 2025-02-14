import os
import pandas as pd
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet
from excel_automation.classes.formats.colors import Color
from excel_automation.classes.formats.formats import Formats
from excel_automation.classes.core.excel_formatter import ExcelFormatter
from typing import Tuple, Literal
import numpy as np
from icecream import ic

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
        self.df_list = df_list
        self.output_name = output_name
        self.writer = ExcelFormatter(df_list, output_name)
        self.workbook: Workbook = self.writer.workbook
        self.format = Formats()
        
    # TODO: Add in formats
    # TODO: Here you can define base chart configs for bar charts
    # TODO: Consider discussing chart font being Aptos Narrow
    def _create_base_chart(self, chart_type: str, chart_subtype: str = ""):
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
 
    # TODO: Data labels position
    # TODO: Add axis in formats
    # TODO: Implement manual logic for specific series (i.e. Peru series) if column.name == Peru
    def create_line_chart(
        self,
        index: int = 0,
        sheet_name: str = "LineChart",
        numeric_type: Literal['integer', 'decimal_1', 'decimal_2', 'percentage'] = "decimal_2",
        chart_template: Literal['line', 'line_simple', 'line_monthly'] = "line",
        axis_title: str = ""
    ) -> Worksheet:
        """
        Creates and inserts a line chart into an Excel worksheet using data from a DataFrame.

        Parameters
        ----------
        index : int, optional
            Index of the DataFrame in df_list to use (default is 0).
        sheet_name : str, optional
            Name of the worksheet (default is "LineChart").
        numeric_type : str, optional
            Defines the number format for the series. Options are:
            'integer', 'decimal_1', 'decimal_2', 'percentage'. (default is 'decimal_2').
        chart_template : str, optional
            Template for the chart configuration: 'line', 'line_simple', 'line_monthly' (default is "line").
        axis_title : str, optional
            Title for the axis (default is an empty string).

        Returns
        -------
        Worksheet
            The worksheet with the inserted chart.

        Raises
        ------
        ValueError
            If the DataFrame is empty or if an invalid chart_type is provided.
        """
        # Definir el formato numérico según 'numeric_type'
        num_format: str = self.format.numeric_types[numeric_type]
        data_df, worksheet = self.writer._write_to_excel(self.df_list[index], sheet_name, num_format, "database")

        # Check if the DataFrame is empty
        if data_df.empty:
            raise ValueError("DataFrame is empty. No data to plot.")
        
        chart = self._create_base_chart('line')

        configs = self.format.charts[chart_template]

        if len(data_df.columns) < 5:
            colors = self.format.charts["line_simple"].get("colors", {})
        else:
            colors = self.format.charts["line"].get("colors", {})
        
        colors = configs["colors"]

        # Base chart configurations
        chart.set_title(configs["title"])
        chart.set_size(configs["size"])
        chart.set_legend(configs["legend"])
        chart.set_plotarea(configs["plotarea"])
        chart.set_chartarea(configs["chartarea"])

        # Add data series with color scheme
        for idx, col in enumerate(data_df.columns[1:]):
            col_letter = chr(66 + idx)  # Get column letter (e.g., B, C, D, ...)
            #print(f"Adding series for column {col}: {col_letter}")  # Debug

            if configs["series"]["marker"].get("none", False):
                marker_config = {"none": True, "type": "none"}
            else:
                marker_config = {
                    **configs["series"]["marker"], 
                    'fill': {'color': colors[(idx) % len(colors)]},
                    'line': {'color': colors[(idx) % len(colors)]},
                    'type': configs["series"]["marker"].get("type", "circle")  
                }

            series_params = {
                **configs["series"],
                'name': f"={sheet_name}!${col_letter}$1",  # Use column letter dynamically
                'categories': f"={sheet_name}!$A$2:$A${len(data_df)+1}",
                'values': f"={sheet_name}!${col_letter}$2:${col_letter}${len(data_df)+1}",
                'line': {
                        'width': 1.75,
                        'dash_type': configs["dash_type"][(idx) % len(configs["dash_type"])],
                        'color': colors[(idx) % len(colors)],
                        },
                'data_labels': {
                    'value': True if chart_template in ("line", "line_simple") else False,
                    'position': 'above' if idx == 1 else 'below',
                    'num_format': num_format,
                    'font':{
                            'color': colors[(idx) % len(colors)],
                            }
                        },
                'marker': marker_config
            }

            chart.add_series(series_params)

        # Axis configuration
        chart.set_y_axis({
            **self.format.charts["y_axis"],
            'name': axis_title,
            'num_format': '0%' if num_format=='0.0%' else num_format,
        })
        
        chart.set_x_axis({
            **self.format.charts["x_axis"],
            'num_format': '0',
        })

        # Insert chart with proper positioning
        position = 'E3' if len(data_df.columns[1:]) < 4 else 'J3'
        worksheet.insert_chart(position, chart, {'x_offset': 25, 'y_offset': 10})
        
        #self.writer.close()  # Automatically saves
        print(f"✅ Gráfico de líneas agregado a la hoja {index + 1}")
        return worksheet
    
    # TODO: Verificar si  col_idx = idx + 1  causa problemas
    # TODO: Ordenar por secciones
    def create_bar_chart(
        self,
        index: int = 0,
        sheet_name: str = "FigX",
        grouping: Literal['standard', 'stacked', 'percentStacked'] = "standard",
        chart_type: Literal['bar', 'column'] = 'column',
        numeric_type: Literal['decimal_1', 'decimal_2', 'integer', 'percentage'] = "decimal_1",
        chart_template: Literal['column', 'column_simple', 'bar', 'bar_single'] = "column",
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
        chart_template : str, optional
            Template for the chart configuration: 'column', 'column_simple', 'bar', or 'bar_single' (default is "column").
        axis_title : str, optional
            Title for the axis (default is an empty string).

        Returns
        -------
        Worksheet
            The worksheet with the inserted chart.

        Raises
        ------
        ValueError
            If the DataFrame is empty or if an invalid chart_type is provided.
        """
        num_format = self.format.numeric_types[numeric_type]
        data_df, worksheet = self.writer._write_to_excel(self.df_list[index], sheet_name, num_format, "database")
        
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

        configs = self.format.charts[chart_template]

        # Validate chart type
        if chart_type not in {"column", "bar"}:
            raise ValueError("Invalid chart_type. Use 'column' (vertical) or 'bar' (horizontal).")
        
        # Predefined formats
        chart = self._create_base_chart(chart_type, subtype)

        # Base chart configurations
        chart.set_title(configs["title"])
        chart.set_size(configs["size"])
        chart.set_legend(configs["legend"])
        chart.set_plotarea(configs["plotarea"])
        chart.set_chartarea(configs["chartarea"])

        colors = configs['colors']

        # Add data series with color scheme
        for idx, col in enumerate(data_df.columns[1:]): # Saltamos la primera columna (categorías), recorre las columnas
            col_idx = idx + 1  
            value_data = (data_df[col] != 0).all()
            
            series_params = {
                **configs['series'],
                'name': [sheet_name, 0, col_idx],
                'values': [sheet_name, 1, col_idx, len(data_df), col_idx],  
                'categories': [sheet_name, 1, 0, len(data_df), 0],  # Categorías en la primera columna 
                'fill': {'color': colors[(col_idx-1) % len(colors)]},
                'data_labels': {**configs['series']['data_labels'], 'num_format': num_format, 'value': value_data},
            }
            chart.add_series(series_params)

        # elif chart_type == "bar":
        #     for row_idx in range(1, len(data_df) + 1):  

        #         series_params = {
        #             **configs['series'],
        #             'name': [sheet_name, row_idx, 0],  
        #             'categories': [sheet_name, 0, 1, 0, data_df.shape[1] - 1],  # Categorías en la primera fila
        #             'fill': {'color': colors[(row_idx-1) % len(colors)]},
        #             'values': [sheet_name, row_idx, 1, row_idx, data_df.shape[1] - 1], 
        #             'data_labels': {**configs['series']['data_labels'], 'num_format': num_format},
        #         }
        #         chart.add_series(series_params)
            
        # TODO: Move to initialize
        # Configure axes
        if chart_type == "column":
            chart.set_y_axis({
                **self.format.charts["y_axis"],
                'name': axis_title,
                'num_format': '0%' if num_format=='0.0%' else num_format,
                'min': 0,
            })
            chart.set_x_axis({
                **self.format.charts["x_axis"],
                'num_format': '@',
                })

        # TODO: Add in formats
        elif chart_type in ("bar", "bar_single"):
            chart.set_legend(configs['legend'])
            chart.set_x_axis({
                'name': axis_title,
                'visible': False,
                'line': {'none': True},
                'major_tick_mark': 'none',
                'major_gridlines': {'visible': False}
            })
            chart.set_y_axis({
                **self.format.charts["x_axis"], # Inverted
                'num_format': '@',
                'reverse': True
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
        chart_template: Literal["index", "data_table", "text_table"] = "text_table",
        numeric_type: Literal['decimal_1', 'decimal_2', 'integer', 'percentage'] = "decimal_1",
    ) -> Worksheet:
        """
        Generates a table in an Excel worksheet based on a DataFrame from the given list of DataFrames.
        Optionally formats the numeric values and applies a specific chart template.

        Parameters
        ----------
        index : int, optional
            The index of the DataFrame in `df_list` to use (default is 0).
        sheet_name : str, optional
            The name of the worksheet where the table will be inserted (default is "TabX").
        chart_template : {'index', 'data_table', 'text_table'}, optional
            The type of chart template to apply to the table (default is 'text_table').
        numeric_type : {'decimal_1', 'decimal_2', 'integer', 'percentage'}, optional
            The numeric format for the values in the table (default is 'decimal_1').

        Returns
        -------
        Worksheet
            The worksheet with the inserted table and chart.

        Raises
        ------
        ValueError
            If the selected DataFrame is empty, a ValueError is raised indicating that there is no data to plot.
        """
        num_format = self.format.numeric_types[numeric_type]
        
        # Retrieve the DataFrame and the corresponding worksheet
        data_df, worksheet = self.writer._write_to_excel(
            df = self.df_list[index],
            sheet_name = sheet_name,
            num_format = num_format,
            format_template = chart_template
        )

        # Check if the DataFrame is empty
        if data_df.empty:
            raise ValueError("DataFrame is empty. No data to plot.")

        # Hide gridlines
        worksheet.hide_gridlines(2)

        print(f"✅ Tabla agregada en la hoja {index + 1}")
        return worksheet

    def save_workbook(self, output_name: str = "Excel_chart"):
        self.writer.save_workbook()
