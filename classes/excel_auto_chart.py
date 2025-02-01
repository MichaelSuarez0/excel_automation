import os
import pandas as pd
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet
from xlsxwriter.format import Format
from xlsxwriter.utility import xl_range
from microsoft_office_automation.classes.colors import Color
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

    def _initialize_chart_formats(self):
        color_list = [Color.BLUE_DARK.value, Color.RED.value, Color.GREEN_DARK.value, Color.ORANGE.value, Color.GRAY.value]
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
            'num_format': '0.00',
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

    def _write_to_excel(self, df: pd.DataFrame, sheet_name: str = "ChartData") -> Tuple[pd.DataFrame, Worksheet]:
        df.to_excel(self.writer, sheet_name=sheet_name, index=False)
        worksheet = self.writer.sheets[sheet_name]
        self.sheet_dfs[sheet_name] = df
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
    
    # TODO: Explicitly call title = ""
    # TODO: Add param for legend and adjust height if not legend
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

        # # Agregar etiquetas de datos (doesn't work)
        # chart.set_series({
        #     0: {
        #         'data_labels': {
        #             'value': True,   
        #             'num_format': '00,00', 
        #             'position': 'outside_end'  # Posición de la etiqueta
        #         }
        #     }
        # })
        return chart

    def create_line_chart(
        self,
        index: int = 0,
        sheet_name: str = "LineChart",
        marker: bool = True
    ) -> Worksheet:
        """Generate line chart with color scheme from Color enum"""
        color_list = [Color.BLUE_DARK.value, Color.RED.value, Color.GREEN_DARK.value, Color.ORANGE.value, Color.GRAY.value]
        data_df, worksheet = self._write_to_excel(self.df_list[index], sheet_name)

        num_format = '# ##0.00' if isinstance(self.df_list[index].iloc[0,1], float) else '# ##0'
        
        # Check if the DataFrame is empty
        if data_df.empty:
            raise ValueError("DataFrame is empty. No data to plot.")
        
        chart = self._create_base_chart(worksheet, 'line')

        # Add data series with color scheme
        for idx, col in enumerate(data_df.columns[1:]):
            col_letter = chr(66 + idx)  # Get column letter (e.g., B, C, D, ...)
            #print(f"Adding series for column {col}: {col_letter}")  # Debug

            series_params = {
                'name': f"={sheet_name}!${col_letter}$1",  # Use column letter dynamically
                'categories': f"={sheet_name}!$A$2:$A${len(data_df)+1}",
                'values': f"={sheet_name}!${col_letter}$2:${col_letter}${len(data_df)+1}",
                'data_labels': {
                    'value': True if idx == 1 else False,
                    'position': 'above',
                    'num_format': num_format},
                'line': {
                    'color': self.chart_formats['line']['colors'][idx % len(self.chart_formats['line']['colors'])], # colors are cycled through the predefined list of colors (colors), even if there are more series than colors.
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
        position = 'E3' if len(data_df.columns[1:]) < 4 else 'J3'
        worksheet.insert_chart(position, chart, {'x_offset': 25, 'y_offset': 10})
        
        #self.writer.close()  # Automatically saves
        print(f"✅ Gráfico de líneas agregado a la hoja {index + 1}")
        return worksheet
    
    # TODO: Ordenar por secciones     
    def create_bar_chart(
        self,
        index: int = 0,
        sheet_name: str = "BarChart",
        grouping: str = "standard",
        chart_type: str = "column"  # "column" (vertical) o "bar" (horizontal)
    ) -> Worksheet:
        """Generate a bar or column chart in Excel from a DataFrame.

        Parameters
        ----------
        index : int, optional
            Index of the DataFrame in df_list to use (default is 0).
        sheet_name : str, optional
            Name of the worksheet (default is "BarChart").
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
        color_list = [Color.BLUE_DARK.value, Color.RED.value, Color.GREEN_DARK.value, Color.ORANGE.value, Color.GRAY.value]
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
        colors = self.chart_formats['line']['colors']

        # Add data series with color scheme
        if chart_type == "column":
            for idx, col in enumerate(data_df.columns[1:]):
                col_idx = idx + 1  # Saltamos la primera columna (categorías)
                color = colors[idx % len(colors)]
                
                series_params = {
                    'name': [sheet_name, 0, col_idx],         
                    'categories': [sheet_name, 1, 0, len(data_df), 0],  # Categorías en la primera columna 
                    'values': [sheet_name, 1, col_idx, len(data_df), col_idx],  
                    'fill': {'color': color},
                    'data_labels': {
                        'value': True,
                        'position': 'outside_end',
                        'num_format': num_format
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
                    'data_labels': {
                        'value': True,
                        'position': 'outside_end',
                        'num_format': num_format
                    },
                }
                chart.add_series(series_params)
            
        # TODO: Consider convertion to dataclasses
        # Configure axes
        if chart_type == "column":
            chart.set_y_axis({
                'name': 'Percentage (%)',
                'num_format': '0',
                'max': 100,
                'min': 0,
                'major_gridlines': {'visible': True, 'line': {'color': Color.GRAY_LIGHT.value}}
            })
            chart.set_x_axis({'name': '', 'text_axis': True, 'num_format': '@'})

        elif chart_type == "bar":
            chart.set_legend({'none': True})
            chart.set_x_axis({
                'name': 'Percentage (%)',
                'num_format': '0',
                'max': 100,
                'min': 0,
                'major_gridlines': {'visible': True, 'line': {'color': Color.GRAY_LIGHT.value}}
            })
            chart.set_y_axis({'name': '', 'text_axis': True, 'num_format': '@'})

        # Insert chart with proper positioning
        position = 'E3' if len(data_df.columns[1:]) < 4 else 'J3'
        worksheet.insert_chart(position, chart, {'x_offset': 25, 'y_offset': 10})
        
        #self.writer.close()  # Save and close workbook
        print(f"✅ Gráfico de barras agregado a la hoja {index + 1}")
        return worksheet
        
    def save_workbook(self):
        self.writer.close()
        print(f'✅ Excel guardado como "{self.output_name}"')
