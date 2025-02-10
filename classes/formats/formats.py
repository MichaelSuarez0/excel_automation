from numpy import True_
from excel_automation.classes.formats.colors import Color
from typing import Any, TypedDict, Literal
from functools import cached_property


class CellConfigs(TypedDict):
    bg_color: str
    font_color: str
    bold: bool
    align: str
    valign: str
    num_format: str
    border: int
    border_color: str


class Formats:
    __slots__ = ('numeric_types', 'cells', 'charts')  

    @cached_property
    def numeric_types(self) -> dict[Literal['date', 'integer', 'decimal1', 'decimal2', 'percentage'], str]:
        return NumericTypes().numeric_types

    @cached_property
    def cells(self) -> dict[Literal['header', 'first_column', 'data'], CellConfigs]:
        return CellFormats().database

    @cached_property
    def charts(self) -> dict[
        Literal[
            'line', 'line_simple', 'column', 'column_simple', 'bar', 
            'marker', 'marker_simple', 'y_axis', 'x_axis'
        ], 
        Any
    ]:
        return ChartFormats().charts



class NumericTypes:
    @cached_property    
    def numeric_types(self) -> dict[Literal['date', 'integer', 'decimal1', 'decimal2', 'percentage'], Any]:
        return {
            'date': 'mmm-yy',
            'integer': '0',
            'decimal1': '0.0',
            'decimal2': '0.00',
            'percentage': '0.0%'
        }
    

class CellFormats:
    @cached_property
    def database(self) -> dict[Literal['header', 'first_column', 'data'], CellConfigs]:
        """Carga y almacena formatos de celdas para hojas que contienen datos"""
        white_borders = {
            'border': 1,
            'border_color': Color.WHITE.value,
        }

        return {
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
    

    @cached_property       
    def index(self) -> dict[Literal['header', 'first_column', 'data'], CellConfigs]:
        """Carga y almacena formatos de celdas para hojas de índices"""
        white_borders = {
            'border': 1,
            'border_color': Color.WHITE.value,
        }

        return {
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


# TODO: Test with column widths
# TODO: Add param for legend and adjust height if not legend
class ChartFormats:

    @cached_property
    def charts(self) -> dict[
        Literal[
            'line', 
            'line_simple', 
            'column', 
            'column_simple', 
            'bar', 
            'marker', 
            'marker_simple', 
            'y_axis', 
            'x_axis'
        ], Any]:
        """
        Carga y almacena formatos para gráficos usando lazy loading.
        Los formatos se definen solo cuando se accede a 'charts' por primera vez.
        """
        line_colors = [
            Color.BLUE_DARK.value, 
            Color.RED.value, 
            Color.ORANGE.value, 
            Color.GREEN_DARK.value, 
            Color.PURPLE.value, 
            Color.GRAY.value
        ]
        line_simple_colors = [
            Color.RED.value, 
            Color.BLUE.value, 
            Color.BLUE_DARK.value
        ]
        column_colors = [
            Color.BLUE_DARK.value, 
            Color.BLUE.value, 
            Color.GREEN_DARK.value, 
            Color.RED.value, 
            Color.ORANGE.value, 
            Color.YELLOW.value, 
            Color.GRAY.value
        ]
        column_simple_colors = [
            Color.BLUE_DARK.value, 
            Color.RED.value, 
            Color.GRAY.value
        ]

        return {
            'line': {
                'colors': line_colors,
                'smooth': True,
                'line': {
                    'colors': line_colors,
                    'width': 1.75,
                    'dash_types': [
                        'solid', 'round_dot', 'round_dot', 'round_dot',
                        'round_dot', 'round_dot', 'round_dot', 'round_dot'
                    ],
                }
            },
            'line_simple': {
                'colors': line_simple_colors,
                'smooth': True,
                'line': {
                    'colors': line_simple_colors,
                    'width': 1.75,
                    'dash_types': ['round_dot', 'square_dot', 'solid']
                }
            },
            'column': {
                'colors': column_colors,
                'fill': {'colors': column_colors},
                'gap': 100,
                'data_labels': {
                    'position': 'outside_end',
                    'font': {
                        'bold': True,
                        'color': Color.BLACK.value,
                        'size': 10.5
                    }
                }
            },
            'column_simple': {
                'colors': column_simple_colors,
                'fill': {'colors': column_simple_colors},
                'gap': 100,
                'data_labels': {
                    'position': 'outside_end',
                    'font': {
                        'bold': True,
                        'color': Color.BLACK.value,
                        'size': 10.5
                    }
                }
            },
            'bar': {
                'colors': column_simple_colors,
                'fill': {'colors': column_simple_colors},
                'gap': 50,
                'data_labels': {
                    'value': True,
                    'position': 'outside_end',
                    'font': {
                        'bold': True,
                        'color': Color.BLACK.value,
                        'size': 10.5
                    }
                }
            },
            'marker': {
                'type': 'circle',
                'size': 6,
                'fill': {'color': line_colors},
                'line': {'color': line_colors}
            },
            'marker_simple': {
                'type': 'circle',
                'size': 6,
                'fill': {'color': line_simple_colors},
                'line': {'color': line_simple_colors}
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