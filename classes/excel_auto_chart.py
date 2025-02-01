import os
import pandas as pd
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet
from xlsxwriter.format import Format
from xlsxwriter.utility import xl_range
from microsoft_office_automation.classes.colors import Color
from typing import Tuple, Optional

script_dir = os.path.abspath(os.path.dirname(__file__))
save_dir = os.path.join(script_dir, "..", "charts")


# TODO: Set axis max and min range dynamically
# TODO: Parámetro para especificar si leyenda o no, si no, que se haga más largo el cuadro del gráfico (porque ya no habría leyenda)
class ExcelAutoChart:
    def __init__(self, df_list: list[pd.DataFrame], output_name: str = "ExcelAutoChart"):
        """Class to write to Excel files from DataFrames and creating charts. Engine: xlsxwriter

        Parameters
        ----------
            df_list (list[pd.DataFrame]): Data that will be written to Excel
            output_name (str, optional): File name for the output file. Defaults to "ExcelAutoChart".
        """
        self.writer = pd.ExcelWriter(os.path.join(save_dir, f'{output_name}.xlsx'), engine='xlsxwriter')
        self.workbook: Workbook = self.writer.book
        self.df_list = df_list
        self._initialize_chart_formats()

    # TODO: Merge with _create_base_chart
    def _initialize_chart_formats(self):
        """Predefine chart-specific formats using the Color enum"""
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

    def _write_to_excel(self, df: pd.DataFrame, sheet_name: str = "ChartData") -> Tuple[pd.DataFrame, Worksheet]:
        """Process source data and prepare worksheet for charting"""
        df.to_excel(self.writer, sheet_name=sheet_name, index=False)
        worksheet = self.writer.sheets[sheet_name]
        return df, worksheet
    
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
    
         
    def create_bar_chart(
        self,
        index: int = 0,
        sheet_name: str = "BarChart",
        grouping: str = "standard"
    ) -> Worksheet:
        """Generate vertical bar chart with color scheme from Color enum"""
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

        # Create column chart (vertical bars)
        chart = self._create_base_chart(worksheet, 'column', subtype)

        # Get colors from predefined formats
        colors = self.chart_formats['line']['colors']

        # Add data series with color scheme
        for idx, col in enumerate(data_df.columns[1:]):
            col_idx = idx + 1  # Skip first column (categories)
            color = colors[idx % len(colors)] 

            series_params = {
                'name': [sheet_name, 0, col_idx],  # Header row
                'categories': [sheet_name, 1, 0, len(data_df), 0],  
                'values': [sheet_name, 1, col_idx, len(data_df), col_idx], 
                'fill': {'color': color},
                'data_labels': {
                    'value': True,
                    'position': 'outside_end',
                    'num_format': num_format},
            }

            chart.add_series(series_params)

        # Y-axis configuration
        chart.set_y_axis({
            'name': 'Percentage (%)',
            'num_format': '0',  # No decimals
            'max': 100,
            'min': 0,
            'major_gridlines': {
                'visible': True,
                'line': {'color': Color.GRAY_LIGHT.value}
            }
        })

        # X-axis configuration
        chart.set_x_axis({
            'name': '',
            'text_axis': True,  # Treat as text categories
            'num_format': '@',  # Text format
        })

        # Insert chart with proper positioning
        position = 'E3' if len(data_df.columns[1:]) < 4 else 'J3'
        worksheet.insert_chart(position, chart, {'x_offset': 25, 'y_offset': 10})
        
        #self.writer.close()  # Save and close workbook
        print(f"✅ Gráfico de barras agregado a la hoja {index + 1}")
        return worksheet
    
    def save_workbook(self):
        return self.writer.close()
