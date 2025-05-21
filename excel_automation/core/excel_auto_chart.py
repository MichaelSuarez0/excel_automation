import pandas as pd
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet
from .excel_writer import ExcelWriterXL
from ..utils import Color, Formats
from typing import Literal, Tuple
from itertools import cycle
import copy

# TODO: Raise errors for invalid templates
# TODO: Use worksheet.dim_colmax
class ExcelAutoChart:
    def __init__(self, df_list: list[pd.DataFrame], output_name: str, output_folder: str):
        """Class to write to Excel files from DataFrames and creating charts. Engine: xlsxwriter

        Parameters
        ----------
        df_list : list(pd.DataFrame):
            Data that will be written to Excel
        output_name : str: 
            Name for the output file (extension already provided)
        output_folder : str:
            Folder path to save the file in.
        """
        self.df_list = df_list
        self.writer = ExcelWriterXL(df_list, output_name, output_folder)
        self.workbook: Workbook = self.writer.workbook
        self.format = Formats()
        self.tab_counter = 0
        self.fig_counter = 0
        
    # TODO: Consider discussing chart font being Aptos Narrow
    # TODO: chart.set_y_axis({'crossing': 'min'}) if values < 0
    def _create_base_chart(self, chart_type: str, chart_subtype: str = ""):
        """Default settings for all chart types"""
        chart = self.workbook.add_chart({'type': chart_type}) if not chart_subtype else self.workbook.add_chart({'type': chart_type, 'subtype': chart_subtype})
        configs = self.format.charts["basic"]
        chart.set_title({'name': ''})
        chart.set_size(configs["size"])
        chart.set_chartarea(configs["chartarea"])
        chart.set_legend(configs["legend"])

        return chart

    def _configure_dynamic_values(self, df: pd.DataFrame, bar: bool = False)-> Tuple[dict, dict, int]:
        "Returns legend, plot_area, sp_axis_num_format, num_font"
        plot_area = copy.deepcopy(self.format.charts["basic"]["plotarea"])
        legend = copy.deepcopy(self.format.charts["basic"]["legend"])
        first_col_max_len = max([len(str(value)) for value in df.iloc[:,0]])
        high_len = first_col_max_len > 5
        columns = df.shape[1] - 1

        if columns == 1:
            legend = {'none': True}
            plot_area["layout"]["height"] += 0.10
        elif columns > 5:
            plot_area["layout"]["height"] -= 0.11
            # legend = {
            #     "layout": {
            #         'x':      0.09,
            #         'y':      0.90,
            #         'width':  0.85,
            #         'height': 0.19
            #         }
            #     }

        if not bar:  # If not bar, reduce height based on len of categories
            if high_len:
                plot_area["layout"]["height"] -= 0.05
            if first_col_max_len >= 9:
                plot_area["layout"]["height"] -= 0.04
            if first_col_max_len >= 13:
                plot_area["layout"]["height"] -= 0.03
            if first_col_max_len >= 16:
                plot_area["layout"]["height"] -= 0.03
        else:   # If bar, reduce width and increase space for x axis based on len of categories
            if high_len:
                plot_area["layout"]["x"] += 0.05
                plot_area["layout"]["width"] -= 0.05
            if first_col_max_len >= 9:
                plot_area["layout"]["x"] += 0.04
                plot_area["layout"]["width"] -= 0.04
            if first_col_max_len >= 13:
                plot_area["layout"]["x"] += 0.03
                plot_area["layout"]["width"] -= 0.03
            if first_col_max_len >= 16:
                plot_area["layout"]["x"] -= 0.03
                plot_area["layout"]["width"] -= 0.03

        sp_axis_num_format = '0' if not isinstance(df.iloc[0,0], pd.Timestamp) else 'mmm-yy'
        num_font = {
            'size': 9 if high_len else 10,
            'rotation': -35 if high_len else 0
            }
        if bar:
            num_font["rotation"] = 0
        
        return legend, plot_area, sp_axis_num_format, num_font
            
    def _configure_axis(self, num_format: str):
        # Axis configuration
        if num_format in ('0', '0.0', '0.00'):
            axis_type = '0'
        elif num_format in ('0.0%'):
            axis_type = '0%'
        else:
            axis_type = num_format
        return axis_type
    
 
    # TODO: Data labels position (automatically?)
    # TODO: Implement manual logic for specific series (i.e. Peru series solid) if column.name == Peru
    # TODO: Si es monthly, el tamaño del axis X debe ser 9 (no 10)
    def create_line_chart(
        self,
        index: int = 0,
        sheet_name: str = "",
        numeric_type: Literal['integer', 'decimal_1', 'decimal_2', 'percentage'] = "decimal_2",
        template: Literal['line', 'line_simple', 'line_single', 'line_monthly'] = "line",
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
            Set a name for a worksheet, else name will be dynamically generated like 'Fig#'.
        numeric_type : str, optional
            Defines the number format for the series. Options are:
            'integer', 'decimal_1', 'decimal_2', 'percentage'. (default is 'decimal_2').
        template : str, optional
            Template for the chart configuration: 'line', 'line_simple', 'line_single', 'line_monthly' (default is "line").
        axis_title : str, optional
            Title for the axis (default is an empty string).
        custom_colors : list of str or None, optional
            A list of custom colors to use for the chart series. If None, the default color cycle will be used.

        Returns
        -------
        Worksheet
            The worksheet with the inserted chart.

        """
        # Initialize configurations
        configs = self.format.charts[template]
        color_cycle = cycle(configs['colors']) if not custom_colors else cycle(custom_colors)
        num_format = self.format.numeric_types[numeric_type]
        
        # Writing data to sheet
        self.fig_counter += 1
        sheet_name = sheet_name if sheet_name else f"Fig{self.fig_counter}"
        df, worksheet = self.writer.write_from_df(
            df = self.df_list[index], 
            sheet_name = sheet_name, 
            num_format = num_format, 
            format_template= "database")
        
        # Load predefined formats
        chart = self._create_base_chart('line')
        
        # Load dynamically modified formats
        legend, plot_area, sp_axis_num_format, num_font = self._configure_dynamic_values(df)
        if axis_title:
            plot_area["layout"]["x"] += 0.03
            plot_area["layout"]["width"] -= 0.03

        # Modify predefined formats
        chart.set_legend(legend)
        chart.set_plotarea(plot_area)
        if "size" in configs:
            chart.set_size(configs["size"])
    
        # Add data series with color scheme
        for idx, col in enumerate(df.columns[1:]):
            num_format = "# ### ##" + num_format if df.iloc[0, idx] else num_format
            col_letter = chr(66 + idx)  # Get column letter (e.g., B, C, D, ...)
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
                    'position': configs['series']['data_labels'].get('position', 'above'),
                    'num_format': num_format,
                    'fill': {'color': Color.BLUE_LIGHT,
                             'transparency': 100 if template != "line_single" else 0},
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
        axis_num_format = self._configure_axis(num_format)

        chart.set_y_axis({
            **self.format.charts['basic']["y_axis"],
            **configs.get('y_axis', {}),
            'name': axis_title,
            'num_format': axis_num_format
        })
        

        chart.set_x_axis({
            **self.format.charts['basic']["x_axis"],
            **configs.get('x_axis', {}),
            'num_format': sp_axis_num_format,
            'num_font': num_font
        })

        # Insert chart with proper positioning
        position = 'E3' if len(df.columns[1:]) < 4 else 'J3'
        worksheet.insert_chart(position, chart, {'x_offset': 90, 'y_offset': 10})
        
        print(f"✅ Gráfico de líneas agregado en la hoja {sheet_name}")
        return worksheet
    

    def create_column_chart(
        self,
        index: int,
        sheet_name: str,
        grouping: Literal['standard', 'stacked', 'percentStacked'] = "standard",
        numeric_type: Literal['decimal_1', 'decimal_2', 'integer', 'percentage'] = "decimal_1",
        template: Literal['column', 'column_simple', 'column_stacked', 'column_single'] = "column",
        axis_title: str = "",
        custom_colors: list[str] | None = None
    ) -> Worksheet:
        """Generate a column chart in Excel from data in a DataFrame list.

        Parameters
        ----------
        index : int, optional
            Index of the DataFrame in df_list to use (default is 0).
        sheet_name : str, optional
            Set a name for a worksheet, else name will be dynamically generated like 'Fig#'.
        grouping : str, optional
            Grouping type: "standard", "stacked", or "percentStacked" (default is "standard").
        numeric_type : str, optional
            Defines the number format for the series. Options are:
            'integer', 'decimal_1', 'decimal_2', 'percentage'. (default is 'decimal_2')
        template : str, optional
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
            If the DataFrame is empty or if an invalid template is provided.
        """
        # Initialize configurations
        configs = self.format.charts[template]
        color_cycle = cycle(configs['colors']) if not custom_colors else cycle(custom_colors)
        num_format = self.format.numeric_types[numeric_type]

        # Writing to sheet
        self.fig_counter += 1
        sheet_name = sheet_name if sheet_name else f"Fig{self.fig_counter}"
        df, worksheet = self.writer.write_from_df(
            df = self.df_list[index], 
            sheet_name = sheet_name, 
            num_format = num_format, 
            format_template= "database")
        
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

        # Load predefined formats
        chart = self._create_base_chart('column', subtype)

        # Load dynamically modified formats
        legend, plot_area, sp_axis_num_format, num_font = self._configure_dynamic_values(df)
        plot_area["layout"]["height"] += 0.05
        if axis_title:
            plot_area["layout"]["x"] += 0.03
            plot_area["layout"]["width"] -= 0.03

        # Modify predefined formats
        chart.set_legend(legend)
        chart.set_plotarea(plot_area)
        if "size" in configs:
            chart.set_size(configs["size"])

        # Add data series with color scheme
        for idx, col in enumerate(df.columns[1:]): # Saltamos la primera columna (categorías), recorre las columnas
            col_idx = idx + 1  
            value_data = (df[col] != 0).all()

            current_color = next(color_cycle)
            if grouping in ("stacked", "percentStacked"):
                if current_color in (Color.YELLOW, Color.GRAY_DARK):
                    label_color = Color.BLACK
                else:
                    label_color = Color.WHITE
            else:
                label_color = Color.WHITE
                
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
                        'color': label_color}
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
            'num_format': sp_axis_num_format,
            'num_font': num_font
        })

        # Insert chart with proper positioning
        position = 'E3' if len(df.columns[1:]) < 4 else 'J3'
        worksheet.insert_chart(position, chart, {'x_offset': 25, 'y_offset': 10})
        
        print(f"✅ Gráfico de columnas agregado en la hoja {sheet_name}")
        return worksheet 
    

    def create_bar_chart(
        self,
        index: int,
        sheet_name: str = "",
        grouping: Literal['standard', 'stacked', 'percentStacked'] = "standard",
        numeric_type: Literal['decimal_1', 'decimal_2', 'integer', 'percentage'] = "decimal_1",
        highlighted_category: str = "",
        template: Literal['bar', 'bar_single', 'bar_double'] = "bar",
        axis_title: str = "",
        custom_colors: list[str] | None = None,
    ) -> Worksheet:
        """Generate a bar chart in Excel from data in a DataFrame list.

        Parameters
        ----------
        index : int, optional
            Index of the DataFrame in df_list to use (default is 0).
        sheet_name : str, optional
            Set a name for a worksheet, else name will be dynamically generated like 'Fig#'.
        grouping : str, optional
            Grouping type: "standard", "stacked", or "percentStacked" (default is "standard").
        numeric_type : str, optional
            Defines the number format for the series. Options are:
            'integer', 'decimal_1', 'decimal_2', 'percentage'. (default is 'decimal_2')
        highlighted_category : str, optional
            Category that will be highlighted with a different color (red).
        template : str, optional
            Template for the chart configuration: 'bar', 'bar_single' or 'bar_double' (default is "bar").
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
        configs = self.format.charts[template]
        color_cycle = cycle(configs['colors']) if not custom_colors else cycle(custom_colors)
        num_format = self.format.numeric_types[numeric_type]

        # Writing to sheet
        self.fig_counter += 1
        sheet_name = sheet_name if sheet_name else f"Fig{self.fig_counter}"
        df, worksheet = self.writer.write_from_df(
            df = self.df_list[index], 
            sheet_name = sheet_name, 
            num_format = num_format, 
            format_template= "database")

        # Raising errors
        if df.empty:
            raise ValueError("DataFrame is empty. No data to plot.")
        if template not in {"bar", "bar_single", "bar_double"}:
            raise ValueError(f"Invalid template for bar chart: {template}. Expected one of 'bar' or 'bar_single'")

        # Map grouping types to xlsxwriter subtypes
        subtype_map = {
            'standard': 'clustered',
            'stacked': 'stacked',
            'percentStacked': 'percent_stacked'
        }
        subtype = subtype_map.get(grouping, 'clustered')

        # Load predefined formats
        chart = self._create_base_chart('bar', subtype)

        # Load dynamically modified formats
        legend, plot_area, sp_axis_num_format, num_font = self._configure_dynamic_values(df, bar = True)

        plot_area["layout"]["height"] += 0.12
        plot_area["layout"]["width"] += 0.04
        plot_area["layout"]["y"] -= 0.02
        
        if axis_title:
            plot_area["layout"]["x"] += 0.03
            plot_area["layout"]["width"] -= 0.03

        # Modify predefined formats
        chart.set_legend(legend)
        chart.set_plotarea(plot_area)
        if "size" in configs:
            chart.set_size(configs["size"])

        # Add data series with color scheme
        for idx, col in enumerate(df.columns[1:]):
            col_idx = idx + 1
            current_color = next(color_cycle)

            points = []
            for row_idx in range(df.shape[0]):
                cell_value = df.iloc[row_idx, col_idx]
                point_format = {}
                
                # Configuración del color (tu lógica original)
                if highlighted_category is not None:
                    category_value = df.iloc[row_idx, 0]
                    if category_value == highlighted_category:
                        color = Color.RED_DARK if col_idx == 1 else Color.RED_LIGHT
                    else:
                        color = current_color

                    point_format['fill'] = {
                        'color': color
                    }
                
                # Configuración especial para valores > 9999
                try:
                    if float(cell_value) > 9999:
                        point_format['data_labels'] = {
                            'num_format': '# ##0'
                        }
                except (ValueError, TypeError):
                    pass
                
                points.append(point_format)
            
            data_labels = {**configs['series']['data_labels'], 'num_format': num_format}
            data_labels.update({'font': {'color': Color.WHITE if col_idx == 1 else Color.BLACK}})
            
            series_params = {
                **configs['series'],
                'name': [sheet_name, 0, col_idx],
                'values': [sheet_name, 1, col_idx, len(df), col_idx],
                'categories': [sheet_name, 1, 0, len(df), 0],
                'fill': {'color': current_color},
                'data_labels': {data_labels},
                'points': points
            }
            chart.add_series(series_params)

        # Configure axes
        axis_format = self._configure_axis(num_format)

        chart.set_x_axis({
            **self.format.charts['basic']["y_axis"], # Inverted
            **configs.get('x_axis', {}),
            'name': axis_title,
            'num_font': num_font
        })
        chart.set_y_axis({
            **self.format.charts['basic']["x_axis"], # Inverted
            **configs.get('y_axis', {}),
            'num_format': axis_format,
            #'num_format': sp_axis_num_format,
            'num_font': num_font
            })
        

        # Insert chart with proper positioning
        position = 'E3' if len(df.columns[1:]) < 4 else 'J3'
        worksheet.insert_chart(position, chart, {'x_offset': 25, 'y_offset': 10})
        
        print(f"✅ Gráfico de barras agregado en la hoja '{sheet_name}'")
        return worksheet 

    # TODO: Padding around max and min values so that error bars do not cut
    # TODO: Not hardcoding delete_series for legend
    # TODO: Dynamically adjust plotarea x and width values based on first column length
    # TODO: Create small and large templates (with recommended max observations for each)
    # TODO: Alinear grids secundarias
    # TODO: Mover la sheet_name "_" al final siempre
    # TODO: Set transparency for mock data labels
    # TODO: For small templates, adjust data labels dynamically
    def create_dot_chart(
        self,
        index: int,
        sheet_name: str = "",
        numeric_type: Literal['decimal_1', 'decimal_2', 'integer', 'percentage'] = "decimal_1",
        highlighted_category: str = "",
        template: Literal['cleveland_dot'] = "cleveland_dot",
        axis_title: str = "",
        custom_colors: list[str] | None = None,
    ) -> Worksheet:
        """Generate a cleveland dot plot in Excel from data in a DataFrame list.

        Parameters
        ----------
        index : int, optional
            Index of the DataFrame in df_list to use (default is 0).
        sheet_name : str, optional
            Set a name for a worksheet, else name will be dynamically generated like 'Fig#'.
        numeric_type : str, optional
            Defines the number format for the series. Options are:
            'integer', 'decimal_1', 'decimal_2', 'percentage'. (default is 'decimal_2')
        highlighted_category : str, optional
            Category that will be highlighted with a different color (red).
        template : str, optional
            Template for the chart configuration: FILL
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
        configs = self.format.charts[template]
        color_cycle = cycle(configs['colors']) if not custom_colors else cycle(custom_colors)
        num_format = self.format.numeric_types[numeric_type]

        # Writing to sheet
        self.fig_counter += 1
        sheet_name = sheet_name if sheet_name else f"Fig{self.fig_counter}"
        df, worksheet = self.writer.write_from_df(
            df = self.df_list[index], 
            sheet_name = sheet_name, 
            num_format = num_format, 
            format_template= "database")

        # Raising errors
        if df.empty:
            raise ValueError("DataFrame is empty. No data to plot.")
        # if template not in {"bar", "bar_single"}:
        #     raise ValueError(f"Invalid template for bar chart: {template}. Expected one of 'bar' or 'bar_single'")
        
        # Create new columns
        n_rows = df.shape[0]
        max_value = int(df.iloc[:,1:].max().max())
        interval = max_value / (n_rows-1)
        y_values = []
        y_values.append(0)
        for number in range(n_rows):
            y_values.append(y_values[-1] + interval)
        
        y_values.pop()
        custom_values = [1] * (n_rows)

        mock_df = pd.DataFrame({
            'y_values': y_values,
            'custom_values': custom_values,
            'difference_values': (df.iloc[:,1] - df.iloc[:,2]).to_list()
            }
        )
        
        mock_sheet_name = "_"
        mock_df, mock_worksheet = self.writer.write_from_df(
            df = mock_df, 
            sheet_name = mock_sheet_name, 
            num_format = num_format, 
            format_template= "database")

        categories = [{'value': category} for category in df.iloc[:,0]]
        
        # Load predefined formats
        chart = self._create_base_chart('scatter')
        
        # Load dynamically modified formats
        legend, plot_area, sp_axis_num_format, num_font = self._configure_dynamic_values(df)
        plot_area["layout"]["height"] += 0.25
        plot_area["layout"]["x"] += 0.1
        plot_area["layout"]["width"] -= 0.03

        if axis_title:
            plot_area["layout"]["x"] += 0.03
            plot_area["layout"]["width"] -= 0.03

        # Modify predefined formats
        chart.set_legend({'delete_series': [-1, -2]})
        chart.set_plotarea(plot_area)
        if "size" in configs:
            chart.set_size(configs["size"])
    
        # Add data series with color scheme
        for idx, col in enumerate(df.columns[1:]):
            current_color = str(next(color_cycle))
            idx += 1

            marker_config = {
                #**configs["series"].get("markers", {}),
                'fill': {'color': current_color},
                'line': {'color': current_color},
                'type': 'circle',
                'size': 8,
            }
            error_bar_config = {
                'x_error_bars':{
                    **configs['x_error_bars'],
                    'plus_values': mock_df['difference_values'].to_list(),
                }
            }

            series_params = {
                **configs["series"],
                **(error_bar_config if idx == 2 else {}),
                'name': [sheet_name, 0, idx],
                'categories': [sheet_name, 1, idx, n_rows, idx], 
                'values': [mock_sheet_name, 1, 0, n_rows, 0], # List of y_values (intervals)
                'data_labels': {
                    **configs['series'].get('data_labels', {}),
                    'value': configs['series']['data_labels'].get('value', True),
                    'position': configs['series']['data_labels'].get('position', 'above'),
                    'num_format': num_format,
                    'font':{
                        'color': configs['series']['data_labels'].get('font', {}).get('color', current_color),
                        'size': configs['series']['data_labels'].get('font', {}).get('size', 10),
                        'bold': configs['series']['data_labels'].get('font', {}).get('bold', False)
                        },
                    },
                'marker': marker_config
            }
            
            chart.add_series(series_params)
        
        # Set false labels
        label_series = {
            **configs["series"],
            'name': [mock_sheet_name, 0, 0],
            'categories': [mock_sheet_name, 1, 1, n_rows, 1], 
            'values': [mock_sheet_name, 1, 0, n_rows, 0],
            'data_labels': {
                **configs['series'].get('data_labels', {}),
                'custom': categories,
                'position': 'left'
                },
            'marker': {
                'fill': {'color': Color.WHITE},
                'line': {'color': Color.WHITE},
                'type': 'circle',
                'size': 8,
            }
        }
            
        chart.add_series(label_series)

        # Configure axes
        axis_format = self._configure_axis(num_format)

        chart.set_x_axis({
            **self.format.charts['basic']["x_axis"],
            **configs.get('x_axis', {}),
            'name': axis_title,
            'num_font': num_font,
        })
        chart.set_y_axis({
            **self.format.charts['basic']["y_axis"],
            **configs.get('y_axis', {}),
            'num_format': axis_format,
            #'num_format': sp_axis_num_format,
            'num_font': num_font,
            'max': max_value,
            'min': 0
            })
        
        
        # Insert chart with proper positioning
        position = 'E3' if len(df.columns[1:]) < 4 else 'J3'
        worksheet.insert_chart(position, chart, {'x_offset': 25, 'y_offset': 10})
        
        print(f"✅ Gráfico de puntos (dot plot) agregado en la hoja '{sheet_name}'")
        self.fig_counter += 1
        return worksheet 
        

    def create_table(
        self,
        index: int,
        sheet_name: str = "",
        template: Literal["database", "index", "data_table", "text_table", "report"] = "text_table",
        numeric_type: Literal['decimal_1', 'decimal_2', 'integer', 'percentage'] = "decimal_1",
        highlighted_categories: str | list = "",
        **kwargs
    ) -> Worksheet:
        """
        Generates a table in an Excel worksheet based on a DataFrame from the given list of DataFrames.
        Optionally formats the numeric values and applies a specific chart template.

        Parameters
        ----------
        index : int, optional
            The index of the DataFrame in `df_list` to use.
        sheet_name : str, optional
            Set a name for a worksheet, else name will be dynamically generated like 'Tab#'.
        template : {'database', 'index', 'data_table', 'text_table'}, optional
            Template style to apply to the worksheet, defaults to "database"
            - "database": Standard format for database-like data (with numbers)
            - "index": Format optimized for index-like sheets or summaries.
            - "data_table": Format optimized for numeric data tables
            - "text_table": Format optimized for text-heavy tables
            - None: No formatting applied, uses pandas default
        numeric_type : {'decimal_1', 'decimal_2', 'integer', 'percentage'}, optional
            The numeric format for the values in the table (default is 'decimal_1').
        highlighted_category : str, optional
            A category value from the first column that will determine row highlighting.
            If a row's first column matches this value, the entire row will be formatted
            with a different color (default is an empty string, meaning no highlighting).
        **kwargs : dict, optional
            Additional arguments for specific templates:
            - For "report": config_dict (dict): Report configuration settings

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
        if not template == "index":
            self.tab_counter += 1
        sheet_name = sheet_name if sheet_name else f"Tab{self.tab_counter}"
        data_df, worksheet = self.writer.write_from_df(self.df_list[index], sheet_name, num_format, template, highlighted_categories, **kwargs)

        # Check if the DataFrame is empty
        if data_df.empty:
            raise ValueError("DataFrame is empty. No data to plot.")

        # Hide gridlines
        worksheet.hide_gridlines(2)

        print(f"✅ Tabla agregada en la hoja '{sheet_name}'")
        return worksheet


    def save_workbook(self):
        self.writer.save_workbook()
    