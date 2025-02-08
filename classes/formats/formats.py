from numpy import True_
from excel_automation.classes.formats.colors import Color
from typing import TypedDict, Literal

class CellFormat(TypedDict):
    bg_color: str
    font_color: str
    bold: bool
    align: str
    valign: str
    num_format: str
    border: int
    border_color: str


class ExcelFormats:
    def __init__(self):
        self._initialize_chart_formats()
        self._initialize_cell_formats()

    def _initialize_cell_formats(self):
        """Crea formatos de celdas para encabezados, primera columna, números y fechas."""
        white_borders = {
            'border': 1,
            'border_color': Color.WHITE.value,
        }

        self.format_cells: dict[Literal['header', 'first_column', 'data'], CellFormat] = {
            'header' : {
                **white_borders,
                'bg_color': Color.BLUE_DARK.value,
                'font_color': Color.WHITE.value,
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True
            },

            'first_column': {
                **white_borders,
                'bg_color': Color.GRAY_LIGHT.value,
                'text_wrap': True
            },

            'data': {
                'border': 1,
                'border_color': Color.GRAY_LIGHT.value,
            },
        
        }
        
        self.numeric_types: dict[Literal['date', 'integer', 'decimal1', 'decimal2', 'percentage'], str] = {
            'date': 'mmm-yy',
            'integer': '0',
            'decimal1': '0.0',
            'decimal2': '0.00',
            'percentage': '0.0%'
        }


    # TODO: Test with column widths
    # TODO: Add param for legend and adjust height if not legend
    def _initialize_chart_formats(self):
        """Crea formatos de gráficos para líneas y columnas (verticales u horizontales)"""
        line_colors = [Color.BLUE_DARK.value, Color.RED.value, Color.ORANGE.value, Color.GREEN_DARK.value, Color.PURPLE.value, Color.GRAY.value]
        line_simple_colors = [Color.RED.value, Color.BLUE.value, Color.BLUE_DARK.value]
        column_colors = [Color.BLUE_DARK.value, Color.BLUE.value, Color.GREEN_DARK.value, Color.RED.value, Color.ORANGE.value, Color.YELLOW.value, Color.GRAY.value]
        column_simple_colors = [Color.BLUE_DARK.value, Color.RED.value, Color.GRAY.value]

        self.format_charts = {
            'line': {
                'colors': line_colors,
                'smooth': True,
                'line':{
                    'colors': line_colors,
                    'width': 1.75,
                    'dash_types': ['solid', 'round_dot','round_dot','round_dot','round_dot','round_dot','round_dot','round_dot'],
                }
            },
            'line_simple': {
                'colors': line_simple_colors,
                'smooth': True,
                'line':{
                    'colors': line_simple_colors,
                    'width': 1.75,
                    'dash_types': ['round_dot', 'square_dot', 'solid']
                }
            },
            'column': {
                'colors': column_colors,
                'fill':{'colors': column_colors},
                'gap': 100,
                'data_labels':{
                    'position': 'outside_end',
                    'font':{
                        'bold': True,
                        'color': Color.BLACK.value, #Color.WHITE.value if color not in (Color.YELLOW.value, Color.GRAY.value) else Color.BLACK.value,
                        'size': 10.5
                    }
                }
            },  
            'column_simple': {
                'colors': column_simple_colors,
                'fill':{'colors': column_simple_colors},
                'gap': 100,
                'data_labels':{
                    'position': 'outside_end',
                    'font':{
                        'bold': True,
                        'color': Color.BLACK.value, #Color.WHITE.value if color not in (Color.YELLOW.value, Color.GRAY.value) else Color.BLACK.value,
                        'size': 10.5
                    }
                }
            },
            'bar': {
                'colors': column_simple_colors,
                'fill':{'colors': column_simple_colors},    
                'gap': 50,
                'data_labels':{
                    'value': True,
                    'position': 'outside_end',
                    'font':{
                        'bold': True,
                        'color': Color.BLACK.value, #Color.WHITE.value if color not in (Color.YELLOW.value, Color.GRAY.value) else Color.BLACK.value,
                        'size': 10.5
                    }
                }
            },  
            'marker': {
                'type': 'circle',
                'size': 6,
                'fill': {'color': line_colors},
                'line':{'color': line_colors}
            },
            'marker_simple': {
                'type': 'circle',
                'size': 6,
                'fill': {'color': line_simple_colors},
                'line':{'color': line_simple_colors}
            },
            'y_axis': {
                'major_gridlines': {
                    'visible': True,
                    'line': {'color': Color.GRAY_LIGHT.value}
                    }
            },
           'x_axis': {
                'name': '',
                'text_axis': True,
                'minor_tick_mark': 'outside',
                'major_tick_mark': 'none'
            }
        }
