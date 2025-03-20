import os
import pandas as pd
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet
from excel_automation.classes.core.excel_writer import ExcelWriterXL
from excel_automation.classes.utils.colors import Color
from excel_automation.classes.utils.formats import Formats
from typing import Tuple, Literal
from itertools import cycle
from icecream import ic
import copy

script_dir = os.path.abspath(os.path.dirname(__file__))
save_dir = os.path.join(script_dir, "..", "..", "charts")

# TODO: Use worksheet.dim_colmax
class ExcelAutoChart:
    def __init__(self, df_list: list[pd.DataFrame], output_name: str = "ExcelAutoChart", output_folder: str = "otros"):
        """Class to write to Excel files from DataFrames and creating charts. Engine: xlsxwriter

        Parameters
        ----------
        df_list : list(pd.DataFrame):
            Data that will be written to Excel
        output_name : str, optional: 
            File name for the output file. Defaults to "ExcelAutoChart".
        output_folder : str, optional:
            Folder name inside "products" to save the file in.
        """
        self.df_list = df_list
        self.writer = ExcelWriterXL(df_list, output_name, output_folder)
        #self.formatter = ExcelFormatter(df_list, output_name, output_folder)
        self.workbook: Workbook = self.writer.workbook
        self.format = Formats()
        self.sheet_count = 0
        self.sheet_list = []
        
    # TODO: Consider discussing chart font being Aptos Narrow
    # TODO: chart.set_y_axis({'crossing': 'min'}) if values < 0
    def _create_base_chart(self, chart_type: str, chart_subtype: str = ""):
        """Default settings for all chart types"""
        chart = self.workbook.add_chart({'type': chart_type}) if not chart_subtype else self.workbook.add_chart({'type': chart_type, 'subtype': chart_subtype})
        configs = self.format.charts["basic"]
        chart.set_title({'name': ''})
        chart.set_size(configs["size"])
        chart.set_chartarea(configs["chartarea"])
        chart.set_plotarea(configs["plotarea"])
        chart.set_legend(configs["legend"])

        return chart
    
    def _configure_axis(self, num_format: str):
        # Axis configuration
        if num_format in ('0.0', '0.00'):
            axis_type = '0'
        elif num_format in ('0.0%'):
            axis_type = '0%'
        else:
            axis_type = num_format
        return axis_type
 
    # TODO: Data labels position (automatically?)
    # TODO: Implement manual logic for specific series (i.e. Peru series solid) if column.name == Peru
    # TODO: Si es monthly, el tamaño del axis X debe ser 9 (no 10)
    # TODO: Menor height si hay muchas series de leyendas
    def create_line_chart(
        self,
        index: int = 0,
        sheet_name: str = "LineChart",
        numeric_type: Literal['integer', 'decimal_1', 'decimal_2', 'percentage'] = "decimal_2",
        chart_template: Literal['line', 'line_simple', 'line_single', 'line_monthly'] = "line",
        axis_title: str = "",
        custom_colors: list[Color] | None = None
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
            Template for the chart configuration: 'line', 'line_simple', 'line_single', 'line_monthly' (default is "line").
        axis_title : str, optional
            Title for the axis (default is an empty string).
        custom_colors : list of str or None, optional
            A list of custom colors to use for the chart series. If None, the default color cycle will be used.

        Returns
        -------
        Worksheet
            The worksheet with the inserted chart.

        Raises
        ------
        ValueError
            If the DataFrame is empty or if an invalid chart_type is provided.
        """
        # Initialize configurations
        configs = self.format.charts[chart_template]
        color_cycle = cycle(configs['colors']) if not custom_colors else cycle(custom_colors)
        num_format = self.format.numeric_types[numeric_type]
        
        # Writing to sheet
        df, worksheet = self.writer.write_from_df(self.df_list[index], sheet_name, num_format, "database")
        self.sheet_list.append(sheet_name)

        # Check if the DataFrame is empty
        if df.empty:
            raise ValueError("DataFrame is empty. No data to plot.")
        
        # Load predefined formats
        chart = self._create_base_chart('line')
        
        # Set legend based on number of columns (series) and modify plotarea
        plot_area = copy.deepcopy(self.format.charts["basic"]["plotarea"])
        if df.shape[1] < 3:
            chart.set_legend({'none': True})
            plot_area["layout"]["height"] += 0.10
        
        if axis_title:
            plot_area["layout"]["x"] += 0.03
            plot_area["layout"]["width"] -= 0.03

        # Override base chart configurations
        chart.set_plotarea(plot_area)
        if "size" in configs:
            chart.set_size(configs["size"])
    
        # Add data series with color scheme
        for idx, col in enumerate(df.columns[1:]):
            col_letter = chr(66 + idx)  # Get column letter (e.g., B, C, D, ...)
            #print(f"Adding series for column {col}: {col_letter}")  # Debug
            current_color = str(next(color_cycle))

            marker_config = {
                **configs["series"].get("markers", {}),
                'fill': {'color': current_color},
                'line': {'color': current_color},
                'type': configs["series"]["marker"].get("type", "circle"),
                'size': configs["series"]["marker"].get("size", 6)  
            } if not configs["series"].get("marker", {}).get("none", False) else {"none": True, "type": "none"}

            series_params = {
                **configs["series"],
                'name': f"={sheet_name}!${col_letter}$1",  # Use column letter dynamically
                'categories': f"={sheet_name}!$A$2:$A${len(df)+1}",
                'values': f"={sheet_name}!${col_letter}$2:${col_letter}${len(df)+1}",
                'line': {
                        'width': configs["series"]["line"].get("width", 1.75),
                        'dash_type': configs["dash_type"][(idx) % len(configs["dash_type"])],
                        'color': current_color,
                        },
                'data_labels': {
                    **configs['series'].get('data_labels', {}),
                    'value': configs['series']['data_labels'].get('value', True),
                    'position': configs['series']['data_labels'].get('position', True),
                    'num_format': num_format,
                    'fill': {'color': Color.BLUE_LIGHT,
                             'transparency': 100 if chart_template != "line_single" else 0},
                    'font':{'color': configs['series']['data_labels'].get('font', {}).get('color', current_color),
                            'size': configs['series']['data_labels'].get('font', {}).get('size', 10),
                            'bold': configs['series']['data_labels'].get('font', {}).get('bold', False)},
                    # **({'border': {
                    #         'width': configs['series']['data_labels']['border'].get('width', 1),  
                    #         'color': current_color
                    #     }} if 'border' in configs['series'].get('data_labels', {}) else {})  # Solo agrega 'border' si está definido
                    },
                'marker': marker_config
            }

            chart.add_series(series_params)

        # Axis configuration
        axis_format = self._configure_axis(num_format)

        chart.set_y_axis({
            **self.format.charts['basic']["y_axis"],
            **configs.get('y_axis', {}),
            'name': axis_title,
            'num_format': axis_format
        })
        
        chart.set_x_axis({
            **self.format.charts['basic']["x_axis"],
            **configs.get('x_axis', {}),
            'num_format': '0' if not isinstance(df.iloc[0,0], pd.Timestamp) else 'mmm-yy',
        })

        # Insert chart with proper positioning
        position = 'E3' if len(df.columns[1:]) < 4 else 'J3'
        worksheet.insert_chart(position, chart, {'x_offset': 90, 'y_offset': 10})
        
        #self.writer.close()  # Automatically saves
        print(f"✅ Gráfico de líneas agregado en la hoja {self.sheet_count + 1}")
        self.sheet_count += 1
        return worksheet
    
    
    def create_column_chart(
        self,
        index: int,
        sheet_name: str,
        grouping: Literal['standard', 'stacked', 'percentStacked'] = "standard",
        numeric_type: Literal['decimal_1', 'decimal_2', 'integer', 'percentage'] = "decimal_1",
        chart_template: Literal['column', 'column_simple', 'column_stacked', 'column_single'] = "column",
        axis_title: str = "",
        custom_colors: list[str] | None = None
    ) -> Worksheet:
        """Generate a column chart in Excel from data in a DataFrame list.

        Parameters
        ----------
        index : int, optional
            Index of the DataFrame in df_list to use (default is 0).
        sheet_name : str, optional
            Assign a name to the new worksheet.
        grouping : str, optional
            Grouping type: "standard", "stacked", or "percentStacked" (default is "standard").
        numeric_type : str, optional
            Defines the number format for the series. Options are:
            'integer', 'decimal_1', 'decimal_2', 'percentage'. (default is 'decimal_2')
        chart_template : str, optional
            Template for the chart configuration: 'column', 'column_simple', 'column_single' or 'column_stacked (default is "column").
        axis_title : str, optional
            Title for the axis (default is an empty string).
        custom_colors : list of str or None, optional
            A list of custom colors to use for the chart series. If None, the default color cycle will be used.

        Returns
        -------
        Worksheet
            The worksheet with the inserted column chart.

        Raises
        ------
        ValueError
            If the DataFrame is empty or if an invalid chart_template is provided.
        """
        # Initialize configurations
        configs = self.format.charts[chart_template]
        color_cycle = cycle(configs['colors']) if not custom_colors else cycle(custom_colors)
        num_format = self.format.numeric_types[numeric_type]

        # Writing to sheet
        df, worksheet = self.writer.write_from_df(self.df_list[index], sheet_name, num_format, "database")
        self.sheet_list.append(sheet_name)
        
        # Raising errors
        if df.empty:
            raise ValueError("DataFrame is empty. No data to plot.")

        # Map grouping types to xlsxwriter subtypes
        subtype_map = {
            'standard': 'clustered',
            'stacked': 'stacked',
            'percentStacked': 'percent_stacked'
        }
        subtype = subtype_map.get(grouping, 'clustered')

        # Predefined formats
        chart = self._create_base_chart('column', subtype)

        # Set legend based on number of columns (series) and modify plotarea
        plot_area = copy.deepcopy(self.format.charts["basic"]["plotarea"])
        if df.shape[1] < 3:
            chart.set_legend({'none': True})
            plot_area["layout"]["height"] += 0.10
        
        if axis_title:
            plot_area["layout"]["x"] += 0.03
            plot_area["layout"]["width"] -= 0.03

        # Override base chart configurations
        chart.set_plotarea(plot_area)

        if "size" in configs:
            chart.set_size(configs["size"])

        # Add data series with color scheme
        for idx, col in enumerate(df.columns[1:]): # Saltamos la primera columna (categorías), recorre las columnas
            col_idx = idx + 1  
            value_data = (df[col] != 0).all()

            current_color = next(color_cycle)
            
            series_params = {
                **configs['series'],
                'name': [sheet_name, 0, col_idx],
                'values': [sheet_name, 1, col_idx, len(df), col_idx],  
                'categories': [sheet_name, 1, 0, len(df), 0],  # Categorías en la primera columna 
                'fill': {'color': current_color},
                'data_labels': {
                    **configs['series']['data_labels'],
                    'num_format': num_format,
                    'value': value_data,
                    'font': {
                        **configs['series']['data_labels']['font'],
                        'color': Color.WHITE if grouping in ("stacked", "percentStacked") else Color.BLACK}
                    },
            }
            chart.add_series(series_params)

        # Axis configuration
        axis_format = self._configure_axis(num_format)

        chart.set_y_axis({
            **self.format.charts['basic']["y_axis"],
            **configs.get('y_axis', {}),
            'name': axis_title,
            'num_format': axis_format,
            'min': 0,
        })
        chart.set_x_axis({
            **self.format.charts['basic']["x_axis"],
            **configs.get('x_axis', {}),
            'num_format': '@',
        })

        # Insert chart with proper positioning
        position = 'E3' if len(df.columns[1:]) < 4 else 'J3'
        worksheet.insert_chart(position, chart, {'x_offset': 25, 'y_offset': 10})
        
        print(f"✅ Gráfico de columnas agregado en la hoja {self.sheet_count + 1}")
        self.sheet_count += 1
        return worksheet 
    

    def create_bar_chart(
        self,
        index: int,
        sheet_name: str = "",
        grouping: Literal['standard', 'stacked', 'percentStacked'] = "standard",
        numeric_type: Literal['decimal_1', 'decimal_2', 'integer', 'percentage'] = "decimal_1",
        highlighted_category: str = "",
        chart_template: Literal['bar', 'bar_single'] = "bar",
        axis_title: str = "",
        custom_colors: list[str] | None = None,
    ) -> Worksheet:
        """Generate a bar chart in Excel from data in a DataFrame list.

        Parameters
        ----------
        index : int, optional
            Index of the DataFrame in df_list to use (default is 0).
        sheet_name : str, optional
            Name of the worksheet (default is "FigX").
        grouping : str, optional
            Grouping type: "standard", "stacked", or "percentStacked" (default is "standard").
        numeric_type : str, optional
            Defines the number format for the series. Options are:
            'integer', 'decimal_1', 'decimal_2', 'percentage'. (default is 'decimal_2')
        highlighted_category : str, optional
            Category that will be highlighted with a different color (red).
        chart_template : str, optional
            Template for the chart configuration: 'bar', or 'bar_single' (default is "bar").
        axis_title : str, optional
            Title for the axis (default is an empty string).
        custom_colors : list of str or None, optional
            A list of custom colors to use for the chart series. If None, the default color cycle will be used.

        Returns
        -------
        Worksheet
            The worksheet with the inserted bar chart.

        Raises
        ------
        ValueError
            If the DataFrame is empty or if an invalid chart_type is provided.
        """
        # Initialize configurations
        configs = self.format.charts[chart_template]
        color_cycle = cycle(configs['colors']) if not custom_colors else cycle(custom_colors)
        num_format = self.format.numeric_types[numeric_type]

        # Writing to sheet
        df, worksheet = self.writer.write_from_df(self.df_list[index], sheet_name, num_format, "database")
        self.sheet_list.append(sheet_name)

        # Raising errors
        if df.empty:
            raise ValueError("DataFrame is empty. No data to plot.")
        if chart_template not in {"bar", "bar_single"}:
            raise ValueError(f"Invalid chart_template for bar chart: {chart_template}. Expected one of 'bar' or 'bar_single'")

        # Map grouping types to xlsxwriter subtypes
        subtype_map = {
            'standard': 'clustered',
            'stacked': 'stacked',
            'percentStacked': 'percent_stacked'
        }
        subtype = subtype_map.get(grouping, 'clustered')

        # Load predefined formats
        chart = self._create_base_chart('bar', subtype)

        # Set legend based on number of columns (series) and modify plotarea
        plot_area = copy.deepcopy(self.format.charts["basic"]["plotarea"])
        plot_area["layout"]["x"] += 0.10
        plot_area["layout"]["width"] -= 0.10
        plot_area["layout"]["height"] += 0.05
        if df.shape[1] < 3:
            chart.set_legend({'none': True})
            plot_area["layout"]["height"] += 0.10
        
        if axis_title:
            plot_area["layout"]["x"] += 0.03
            plot_area["layout"]["width"] -= 0.03

        # Override base chart configurations
        chart.set_plotarea(plot_area)

        if "size" in configs:
            chart.set_size(configs["size"])

        # Add data series with color scheme
        for idx, col in enumerate(df.columns[1:]): # Saltamos la primera columna (categorías), recorre las columnas
            col_idx = idx + 1  
            value_data = (df[col] != 0).all()
            current_color = next(color_cycle)

            points = []
            if highlighted_category is not None:
                for row_idx in range(df.shape[0]):
                    category_value = df.iloc[row_idx, 0]
                    if category_value == highlighted_category:
                        points.append({'fill': {'color': Color.RED}})
                    else:
                        points.append({'fill': {'color': current_color}})
                    
            series_params = {
                **configs['series'],
                'name': [sheet_name, 0, col_idx],
                'values': [sheet_name, 1, col_idx, len(df), col_idx],  
                'categories': [sheet_name, 1, 0, len(df), 0],  # Categorías en la primera columna 
                'fill': {'color': current_color},
                'data_labels': {**configs['series']['data_labels'], 'num_format': num_format, 'value': value_data},
                'points': points
            }
            chart.add_series(series_params)

        # Configure axes
        axis_format = self._configure_axis(num_format)

        chart.set_x_axis({
            **self.format.charts['basic']["y_axis"], # Inverted
            **configs.get('x_axis', {}),
            'name': axis_title,
        })
        chart.set_y_axis({
            **self.format.charts['basic']["x_axis"], # Inverted
            **configs.get('y_axis', {}),
            'num_format': axis_format,
            })

        # Insert chart with proper positioning
        position = 'E3' if len(df.columns[1:]) < 4 else 'J3'
        worksheet.insert_chart(position, chart, {'x_offset': 25, 'y_offset': 10})
        
        print(f"✅ Gráfico de barras agregado en la hoja {self.sheet_count + 1}")
        self.sheet_count += 1
        return worksheet 

    
    def create_table(
        self,
        index: int,
        sheet_name: str,
        chart_template: Literal["database", "index", "data_table", "text_table"] = "text_table",
        numeric_type: Literal['decimal_1', 'decimal_2', 'integer', 'percentage'] = "decimal_1",
        highlighted_category: str =""
    ) -> Worksheet:
        """
        Generates a table in an Excel worksheet based on a DataFrame from the given list of DataFrames.
        Optionally formats the numeric values and applies a specific chart template.

        Parameters
        ----------
        index : int, optional
            The index of the DataFrame in `df_list` to use.
        sheet_name : str, optional
            The name of the worksheet where the table will be inserted.
        chart_template : {'database', 'index', 'data_table', 'text_table'}, optional
            The type of chart template to apply to the table (default is 'text_table').
        numeric_type : {'decimal_1', 'decimal_2', 'integer', 'percentage'}, optional
            The numeric format for the values in the table (default is 'decimal_1').
        highlighted_category : str, optional
            A category value from the first column that will determine row highlighting.
            If a row's first column matches this value, the entire row will be formatted
            with a different color (default is an empty string, meaning no highlighting).

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
        data_df, worksheet = self.writer.write_from_df(self.df_list[index], sheet_name, num_format, chart_template, highlighted_category)
        self.sheet_list.append(sheet_name)

        # Check if the DataFrame is empty
        if data_df.empty:
            raise ValueError("DataFrame is empty. No data to plot.")

        # Hide gridlines
        worksheet.hide_gridlines(2)

        print(f"✅ Tabla agregada en la hoja {self.sheet_count + 1}")
        self.sheet_count += 1
        return worksheet

    def save_workbook(self):
        self.writer.save_workbook()
