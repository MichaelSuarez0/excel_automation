import os
import pandas as pd
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet
from xlsxwriter.format import Format
from xlsxwriter.utility import xl_range
from excel_automation.classes.colors import Color
from typing import Tuple, Optional
import numpy as np

script_dir = os.path.abspath(os.path.dirname(__file__))
save_dir = os.path.join(script_dir, "..", "charts")


# TODO: Set axis max and min range dynamically
# TODO: Parámetro para especificar si leyenda o no, si no, que se haga más largo el cuadro del gráfico (porque ya no habría leyenda)
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
        self._initialize_chart_formats()
        self._initialize_cell_formats()

    def _initialize_cell_formats(self):
        """Crea formatos de celdas para encabezados, primera columna, números y fechas."""
        self.header_format = self.workbook.add_format({
            'bg_color': Color.BLUE_DARK.value,
            'font_color': Color.WHITE.value,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'border_color': Color.WHITE.value,
        })
        self.first_column_format = self.workbook.add_format({
            'bg_color': Color.GRAY_LIGHT.value,
            'border': 1,
            'border_color': Color.WHITE.value,
        })
        self.number_format = self.workbook.add_format({
            'num_format': '0.0',
            'border': 1,
            'border_color': Color.GRAY_LIGHT.value,
        })
        self.integer_format = self.workbook.add_format({
            'num_format': '0',
            'border': 1,
            'border_color': Color.GRAY_LIGHT.value,
        })
        self.date_format = self.workbook.add_format({
            'num_format': 'mmm-yy',
            'border': 1,
            'border_color': Color.WHITE.value,
            'bg_color': Color.GRAY_LIGHT.value,
        })

    # TODO: Test with column widths
    # TODO: Ajusta tamaño de la leyenda para gráficos no simples
    def _initialize_chart_formats(self):
        line_colors = [Color.BLUE_DARK.value, Color.RED.value, Color.ORANGE.value, Color.GREEN_DARK.value, Color.PURPLE.value, Color.GRAY.value]
        line_simple_colors = [Color.RED.value, Color.BLUE.value, Color.BLUE_DARK.value]
        column_colors = [Color.BLUE_DARK.value, Color.BLUE.value, Color.GREEN_DARK.value, Color.RED.value, Color.ORANGE.value, Color.YELLOW.value, Color.GRAY.value]
        column_simple_colors = [Color.BLUE_DARK.value, Color.RED.value, Color.GRAY.value]
        axis_base = {
            'minor_tick_mark': 'outside',
            'major_tick_mark': 'none',
            'major_gridlines': {'visible': True, 'line': {'color': Color.GRAY_LIGHT.value}}
        }

        self.chart_formats = {
            'line': {
                'colors': line_colors,
                'width': 1.75,
                'dash_types': ['solid', 'round_dot','round_dot','round_dot','round_dot','round_dot','round_dot','round_dot']
            },
            'line_simple': {
                'colors': line_simple_colors,
                'width': 1.75,
                'dash_types': ['round_dot', 'square_dot', 'solid']
            },
            'column': {
                'colors': column_colors,
                'width': 2.5,
            },
            'column_simple': {
                'colors': column_simple_colors,
                'width': 2.5,
                'dash_types': ['solid']
            },
            'marker': {
                'size': 6,
                'colors': line_colors
            },
            'marker_simple': {
                'size': 6,
                'colors': line_simple_colors
            },
            'axis_percentage':{
                **axis_base,
                'name': 'Porcentaje (%)',
                'num_format': '0',
                'max': 100,
                'min': 0,

            },
            'axis_number':{
                **axis_base,
                'name': 'Unidades',
                'num_format': '0',
            }
        }
        

    def _write_to_excel(self, df: pd.DataFrame, sheet_name: str = "ChartData", apply_format = True) -> Tuple[pd.DataFrame, Worksheet]:
        df.to_excel(self.writer, sheet_name=sheet_name, index=False)
        worksheet = self.writer.sheets[sheet_name]
        self.sheet_dfs[sheet_name] = df
        if apply_format:
            self._apply_formatting_to_worksheet(worksheet, df)
        return df, worksheet

    def _apply_formatting_to_worksheet(self, worksheet: Worksheet, df: pd.DataFrame):
        """Applies formatting only to cells with data."""
        # Set column widths
        worksheet.set_column('A:A', 15)
        if len(df.columns) > 1:
            worksheet.set_column(1, len(df.columns) - 1, 10)

        # Hide gridlines
        worksheet.hide_gridlines(2)

        # Determine format for the first column
        first_col = df.columns[0]
        if pd.api.types.is_datetime64_any_dtype(df[first_col]):
            first_col_fmt = self.date_format
        else:
            first_col_fmt = self.first_column_format

        # Write headers with header format
        for col_num, col_name in enumerate(df.columns):
            worksheet.write(0, col_num, col_name, self.header_format)

        # Write data cells with appropriate formats
        for row_idx in range(df.shape[0]):
            # First column (e.g., dates or text)
            cell_value = df.iloc[row_idx, 0]
            worksheet.write(row_idx + 1, 0, cell_value, first_col_fmt)

            # Other columns (numeric data)
            for col_idx in range(1, df.shape[1]):
                cell_value = df.iloc[row_idx, col_idx]
                dtype = df.dtypes.iloc[col_idx]  # Use .iloc for positional indexing

                # Skip NaN/Inf values by checking if the value is NaN or Inf
                if pd.isna(cell_value) or np.isinf(cell_value):
                    worksheet.write(row_idx + 1, col_idx, '')  # Write an empty cell

                else:
                    if np.issubdtype(dtype, np.floating):
                        fmt = self.number_format
                    else:
                        fmt = self.integer_format

                    worksheet.write(row_idx + 1, col_idx, cell_value, fmt)
    
    # TODO: Here you can define base chart configs for bar charts
    # TODO: Explicitly call title = ""
    # TODO: Add param for legend and adjust height if not legend
    # TODO: Chart font should be Aptos Narrow
    def _create_base_chart(self, worksheet: Worksheet, chart_type: str, chart_subtype: str = ""):
        """Default settings for all chart types"""
        chart = self.workbook.add_chart({'type': chart_type}) if not chart_subtype else self.workbook.add_chart({'type': chart_type, 'subtype': chart_subtype})

        chart.set_title({'name': ''})
        chart.set_size({'width': 600, 'height': 420})
        chart.set_legend({'position': 'bottom'})
        chart.set_plotarea({
            'layout': {
                'x':      0.11,
                'y':      0.10,
                'width':  0.83,
                'height': 0.75,
            }
        })
        chart.set_chartarea({'border': {'none': True}})

        return chart

    # TODO: Decidir si quedarse con un decimal o dos
    # TODO: Implement manual logic for specific series (i.e. Peru series) if column.name == Peru
    def create_line_chart(
        self,
        index: int = 0,
        sheet_name: str = "LineChart",
        markers_add: bool = True
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

        Returns
        -------
        Worksheet
            The worksheet with the inserted chart.
        """
        color_list = [Color.BLUE_DARK.value, Color.RED.value, Color.GREEN_DARK.value, Color.ORANGE.value, Color.GRAY.value]
        data_df, worksheet = self._write_to_excel(self.df_list[index], sheet_name)

        num_format = '# ##0.0' if isinstance(self.df_list[index].iloc[0,1], float) else '# ##0'
        
        # Check if the DataFrame is empty
        if data_df.empty:
            raise ValueError("DataFrame is empty. No data to plot.")
        
        chart = self._create_base_chart(worksheet, 'line')

        if len(data_df.columns) < 5:
            colors: list = self.chart_formats['line_simple']['colors']
            dashes: list = self.chart_formats['line_simple']['dash_types']
            markers: list = self.chart_formats['marker_simple']['colors']
        else:
            colors: list = self.chart_formats['line']['colors']
            dashes: list = self.chart_formats['line']['dash_types']
            markers: list = self.chart_formats['marker']['colors']

        # Add data series with color scheme
        for idx, col in enumerate(data_df.columns[1:]):
            col_letter = chr(66 + idx)  # Get column letter (e.g., B, C, D, ...)
            #print(f"Adding series for column {col}: {col_letter}")  # Debug

            series_params = {
                'name': f"={sheet_name}!${col_letter}$1",  # Use column letter dynamically
                'categories': f"={sheet_name}!$A$2:$A${len(data_df)+1}",
                'values': f"={sheet_name}!${col_letter}$2:${col_letter}${len(data_df)+1}",
                'smooth': True,
                'data_labels': {
                    'value': True if idx == 1 else False,
                    'position': 'below',
                    'num_format': num_format,
                    'font':{
                            'color': colors[idx % len(colors)],
                            }
                        },
                'line': {
                    'color': colors[idx % len(colors)],
                    'width': self.chart_formats['line']['width'],
                    'dash_type': dashes[idx % len(dashes)],
                }
            }

            if markers_add:
                marker_color = markers[idx % len(markers)]
                series_params['marker'] = {
                    'type': 'circle',
                    'size': self.chart_formats['marker']['size'],
                    'fill': {'color': marker_color},
                    'line': {'color': marker_color},
                }

            chart.add_series(series_params)

        # Axis configuration
        chart.set_y_axis({
            'name': 'Porcentaje (%)',
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
            'minor_tick_mark': 'outside',
            'major_tick_mark': 'none'
        })

        # Insert chart with proper positioning
        position = 'E3' if len(data_df.columns[1:]) < 4 else 'J3'
        worksheet.insert_chart(position, chart, {'x_offset': 25, 'y_offset': 10})
        
        #self.writer.close()  # Automatically saves
        print(f"✅ Gráfico de líneas agregado a la hoja {index + 1}")
        return worksheet
    
    # TODO: Ordenar por secciones     
    def create_bar_chart(
        self,
        index: int = 0,
        sheet_name: str = "FigX",
        grouping: str = "standard",
        chart_type: str = "column"  # "column" (vertical) o "bar" (horizontal)
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

        Returns
        -------
        Worksheet
            The worksheet with the inserted chart.

        Raises
        ------
        ValueError
            If the DataFrame is empty or if an invalid chart_type is provided.
        """
        data_df, worksheet = self._write_to_excel(self.df_list[index], sheet_name)

        num_format = '# ##0.00' if isinstance(self.df_list[index].iloc[0,1], float) else '# ##0'
        
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

        if grouping == "standard" or len(data_df.columns) < 4:
            colors: list = self.chart_formats['column_simple']['colors']
        else:
            colors: list = self.chart_formats['column']['colors']
        

        # Add data series with color scheme
        if chart_type == "column":
            for idx, col in enumerate(data_df.columns[1:]): # Saltamos la primera columna (categorías), recorre las columnas
                col_idx = idx + 1  
                color = colors[idx % len(colors)]
                value_data = (data_df[col] != 0).all()
                
                series_params = {
                    'name': [sheet_name, 0, col_idx],         
                    'categories': [sheet_name, 1, 0, len(data_df), 0],  # Categorías en la primera columna 
                    'values': [sheet_name, 1, col_idx, len(data_df), col_idx],  
                    'fill': {'color': color},
                    'gap': 100,
                    'data_labels': {
                        'value': value_data,
                        'position': 'outside_end',
                        'num_format': num_format,
                        'font':{
                            'bold': True,
                            'color': Color.WHITE.value if color not in (Color.YELLOW.value, Color.GRAY.value) else Color.BLACK.value,
                            'size': 10.5
                        }
                    },
                }
                chart.add_series(series_params)
                
        elif chart_type == "bar":
            for row_idx in range(1, len(data_df) + 1):  
                color = colors[(row_idx - 1) % len(colors)]
                
                series_params = {
                    'name': [sheet_name, row_idx, 0],  
                    'categories': [sheet_name, 0, 1, 0, data_df.shape[1] - 1],  # Categorías en la primera fila
                    'values': [sheet_name, row_idx, 1, row_idx, data_df.shape[1] - 1],  
                    'fill': {'color': color},
                    'gap': 50,
                    'data_labels': {
                        'value': True,
                        'position': 'outside_end',
                        'num_format': num_format
                    },
                }
                chart.add_series(series_params)
            
        # TODO: Move to initialize
        # Configure axes
        if chart_type == "column":
            chart.set_y_axis({
                'name': 'Porcentaje (%)',
                'num_format': '0',
                'max': 100,
                'min': 0,
                'major_gridlines': {'visible': True, 'line': {'color': Color.GRAY_LIGHT.value}}
            })
            chart.set_x_axis({
                'name': '',
                'text_axis': True,
                'num_format': '@',
                'minor_tick_mark': 'outside',
                'major_tick_mark': 'none'})

        elif chart_type == "bar":
            chart.set_legend({'none': True})
            chart.set_x_axis({
                'name': 'Porcentaje (%)',
                'num_format': '0',
                'max': 100,
                'min': 0,
                'minor_tick_mark': 'outside',
                'major_tick_mark': 'none',
                'major_gridlines': {'visible': True, 'line': {'color': Color.GRAY_LIGHT.value}}
            })
            chart.set_y_axis({
                'name': '',
                'text_axis': True,
                'num_format': '@',
                'minor_tick_mark': 'outside',
                'major_tick_mark': 'none',})

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
        data_df, worksheet = self._write_to_excel(self.df_list[index], sheet_name, apply_format=False)
        
        # Check if the DataFrame is empty
        if data_df.empty:
            raise ValueError("DataFrame is empty. No data to plot.")
        
        # Set column widths
        worksheet.set_column('A:A', 24)
        worksheet.set_column('B:B', 60)

        # Hide gridlines
        worksheet.hide_gridlines(2)

        # Define base formats
        gray_format = self.workbook.add_format({'bg_color': Color.GRAY_LIGHT.value, 'text_wrap': True, 'valign': 'vcenter'})
        default_format = self.workbook.add_format({'text_wrap': True, 'valign': 'vcenter'})
        bold_format = self.workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'vcenter'})
        gray_bold_format = self.workbook.add_format({'bg_color': Color.GRAY_LIGHT.value, 'bold': True, 'text_wrap': True, 'valign': 'vcenter'})

        # Write headers with header format
        for col_num, col_name in enumerate(data_df.columns):
            worksheet.write(0, col_num, col_name, self.header_format)

        # Write table contents with alternating colors and bold for first column
        for row_idx in range(data_df.shape[0]):
            for col_idx in range(data_df.shape[1]):
                cell_value = data_df.iloc[row_idx, col_idx]

                # Select format based on column and row index
                if col_idx == 0:
                    cell_format = gray_bold_format if row_idx % 2 == 0 else bold_format  
                else:
                    cell_format = gray_format if row_idx % 2 == 0 else default_format 

                worksheet.write(row_idx + 1, col_idx, cell_value, cell_format)

        print(f"✅ Tabla agregada en la hoja {index + 1}")
        return worksheet

    def save_workbook(self):
        self.writer.close()
        print(f'✅ Excel guardado como "{self.output_name}"')
