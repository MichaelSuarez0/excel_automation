from excel_automation.classes.formats.colors import Color
from typing import Any, TypedDict, Literal
from functools import cached_property


class CellConfigs(TypedDict):
    bg_color: str
    font_color: str
    font_size: int
    bold: bool
    align: str
    valign: str
    num_format: str
    border: int
    border_color: str


class Formats:
    @cached_property
    def numeric_types(self) -> dict[Literal['date', 'integer', 'decimal_1', 'decimal_2', 'percentage'], str]:
        return NumericTypes().numeric_types

    @cached_property
    def cells(self) -> dict[Literal['database', 'index', 'data_table', 'text_table'], dict[Literal['header', 'first_column', 'data'], CellConfigs]]:
        return CellFormats().cells

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
    def numeric_types(self) -> dict[Literal['date', 'integer', 'decimal_1', 'decimal_2', 'percentage'], Any]:
        return {
            'date': 'mmm-yy',
            'integer': '0',
            'decimal_1': '0.0',
            'decimal_2': '0.00',
            'percentage': '0.0%'
        }
    

class CellFormats:
    @cached_property
    def cells(self) -> dict[Literal['database', 'data_table', 'text_table', 'index'], dict[Literal['header', 'first_column', 'data'], CellConfigs]]:
        """Carga y almacena formatos de celdas para hojas que contienen datos (database e index)"""
        white_borders = {
            'border': 1,
            'border_color': Color.WHITE.value,
        }

        return { 
            'database': {
                'header': {
                    **white_borders,
                    'bg_color': Color.BLUE_DARK.value,
                    'font_color': Color.WHITE.value,
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True,
                    'font_size': 10
                },
                'first_column': {
                    **white_borders,
                    'bg_color': Color.GRAY_LIGHT.value,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'font_size': 10
                },
                'data': {
                    'border': 1,
                    'border_color': Color.GRAY_LIGHT.value,
                    'valign': 'vcenter',
                    'font_size': 10
                }
            },
            'text_table': {
                'header': {
                    **white_borders,
                    'bg_color': Color.BLUE_DARK.value,
                    'font_color': Color.WHITE.value,
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True,
                    'font_size': 10
                },
                'first_column': {
                    **white_borders,
                    'bg_color': Color.GRAY_LIGHT.value,
                    'valign': 'vcenter',
                    'align': 'justify',
                    'text_wrap': True,
                    'font_size': 10
                },
                'data': {
                    'border': 1,
                    'border_color': Color.GRAY_LIGHT.value,
                    'align': 'justify',
                    'valign': 'vcenter',
                    'font_size': 10
                }
            },
            'data_table': {
                'header': {
                    **white_borders,
                    'bg_color': Color.BLUE_DARK.value,
                    'font_color': Color.WHITE.value,
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True,
                    'font_size': 10
                },
                'first_column': {
                    **white_borders,
                    'bg_color': Color.GRAY_LIGHT.value,
                    'text_wrap': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'font_size': 10
                },
                'data': {
                    'border': 1,
                    'border_color': Color.GRAY_LIGHT.value,
                    'align': 'center',
                    'valign': 'vcenter',
                    'font_size': 10
                }
            },
            'index': {
                'header': {
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
                    'valign': 'center'
                }
            }
        }

# TODO: Test with column widths
# TODO: Add param for legend and adjust height if not legend
class ChartFormats:

    @cached_property
    def charts(self) -> dict[
        Literal[
            'line', 
            'line_simple',
            'line_monthly,' 
            'column', 
            'column_simple', 
            'bar',  
            'bar_single', 
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
            Color.BLUE_DARK.value,
            Color.BLUE.value, 
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
        bar_colors = [
            Color.RED.value, 
            Color.BLUE_DARK.value, 
            Color.YELLOW.value
        ]

        return {
            'line': {
                'title': {'name': ''},
                'size': {'width': 600, 'height': 420},
                'legend': {'position': 'bottom'},
                'chartarea': {'border': {'none': True}},
                'colors': line_colors,
                'dash_type': [
                            'solid', 'round_dot', 'round_dot', 'round_dot',
                            'round_dot', 'round_dot', 'round_dot', 'round_dot'
                        ],
                'plotarea': {
                    'layout':{
                        'x':      0.11,
                        'y':      0.08,
                        'width':  0.83,
                        'height': 0.75
                        }
                },
                'series': {
                    'smooth': True,
                    'line': {
                        'width': 1.75
                        },
                    'marker': {
                        'type': 'circle',
                        'size': 6
                    },
                    'data_labels': {'value': False}
                }
            },
            'line_simple': {
                'title': {'name': ''},
                'size': {'width': 580, 'height': 420},
                'legend': {'position': 'bottom'},
                'chartarea': {'border': {'none': True}},
                'colors': line_simple_colors,
                'dash_type': ['round_dot', 'square_dot', 'solid'],
                'plotarea': {
                    'layout':{
                        'x':      0.11,
                        'y':      0.08,
                        'width':  0.83,
                        'height': 0.75
                        }
                },    
                'series': {
                    'smooth': True,
                    'line': {
                        'width': 1.75
                    },
                    'marker': {
                        'type': 'circle',
                        'size': 6
                    }
                }
            },
            'line_monthly': {
                'title': {'name': ''},
                'size': {'width': 580, 'height': 360},
                'legend': {'position': 'bottom'},
                'chartarea': {'border': {'none': True}},
                'colors': line_simple_colors,
                'dash_type': [
                            'square_dot', 'square_dot', 'round_dot'
                        ],
                'plotarea': {
                    'layout':{
                        'x':      0.11,
                        'y':      0.06,
                        'width':  0.85,
                        'height': 0.75
                        }
                },
                'series': {
                    'smooth': True,
                    'line': {
                        'width': 1.75
                        },
                    'marker': {
                        'none': True
                    },
                    'data_labels': {'value': False}
                }
            },
            'column': {
                'title': {'name': ''},
                'size': {'width': 580, 'height': 420},
                'legend': {'position': 'bottom'},
                'chartarea': {'border': {'none': True}},
                'colors': column_colors,
                'plotarea': {
                    'layout':{
                        'x':      0.11,
                        'y':      0.08,
                        'width':  0.83,
                        'height': 0.75
                        }
                },
                'series': {
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
                }
            },
            'column_simple': {
                'title': {'name': ''},
                'size': {'width': 600, 'height': 420},
                'legend': {'position': 'bottom'},
                'chartarea': {'border': {'none': True}},
                'colors': column_simple_colors,
                'plotarea': {
                    'layout':{
                        'x':      0.11,
                        'y':      0.06,
                        'width':  0.85,
                        'height': 0.78
                        }
                },                
                'series':{
                    'gap': 100,
                    'data_labels': {
                        'position': 'outside_end',
                        'font': {
                            'bold': True,
                            'color': Color.BLACK.value,
                            'size': 10.5
                        }
                    }
                }
            },
            'bar': {
                'title': {'name': ''},
                'size': {'width': 570, 'height': 340},
                'legend': {'position': 'bottom'},
                'chartarea': {'border': {'none': True}},
                'colors': bar_colors,
                'plotarea': {
                    'layout':{
                        'x':      0.01,
                        'y':      0.03,
                        'width':  0.80,
                        'height': 0.89
                        }
                },
                'series':{
                    'gap': 60,
                    'data_labels': {
                        'value': True,
                        'position': 'outside_end',
                        'font': {
                            'bold': True,
                            'color': Color.BLACK.value,
                            'size': 10
                        }
                    }
                }
            },

            'bar_single': {
                'title': {'name': ''},
                'size': {'width': 570, 'height': 460},
                'legend': {'none': True},
                'chartarea': {'border': {'none': True}},
                'colors': column_simple_colors,
                'plotarea': {
                    'layout':{
                        'x':      0.01,
                        'y':      0.03,
                        'width':  0.80,
                        'height': 0.94
                        }
                },
                'series':{
                    'gap': 40,
                    'data_labels': {
                        'value': True,
                        'position': 'outside_end',
                        'font': {
                            'bold': True,
                            'color': Color.BLACK.value,
                            'size': 10
                        }
                    }
                }
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


        # chart.set_title({'name': ''})
        # chart.set_size({'width': 600, 'height': 420})
        # chart.set_legend({'position': 'bottom'})
        # chart.set_plotarea({
        #     'layout': {
        #         'x':      0.11,
        #         'y':      0.09,
        #         'width':  0.83,
        #         'height': 0.75,
        #     }
        # })
        # chart.set_chartarea({'border': {'none': True}})
