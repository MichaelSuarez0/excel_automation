from .colors import Color
from typing import Any, Literal
from functools import cached_property
from dataclasses import dataclass
from enum import StrEnum

class Alignment(StrEnum):
    left = 'left'
    right = 'right'
    center = 'center'
    justify = 'justify'


@dataclass
class CellConfig:
    bg_color: str
    font_color: str
    font_size: int
    bold: bool
    align: Alignment
    valign: Alignment
    num_format: str
    border: int
    border_color: str
    text_wrap: bool = False

@dataclass
class Formats():
    @cached_property
    def numeric_types(self) -> dict[Literal['date', 'integer', 'decimal_1', 'decimal_2', 'percentage'], str]:
        return NumericTypes().numeric_types

    @cached_property
    def cells(self) -> dict[Literal['database', 'index', 'data_table', 'text_table', 'report'], dict[Literal['header', 'first_column', 'data', 'column_widths'], CellConfig]]:
        return CellFormats().cells

    @cached_property
    def charts(self) -> dict[
        Literal[
            'basic', 'line', 'line_simple', 'line_single', 'line_monthly', 'column', 'column_simple', 'column_single', 'column_stacked', 
            'bar', 'bar_single', 'bar_double', 'cleveland_dot',
            # 'marker', 'marker_simple', 'y_axis', 'x_axis'
        ], 
        Any
    ]:
        return ChartFormats().charts


class NumericTypes:
    @cached_property    
    def numeric_types(self) -> dict[Literal['date', 'integer', 'decimal_1', 'decimal_2', 'percentage'], str]:
        return {
            'date': 'mmm-yy',
            'integer': '0',
            'decimal_1': '0.0',
            'decimal_2': '0.00',
            'percentage': '0.0%'
        }
    

class CellFormats:
    @cached_property
    def cells(self) -> dict[
        Literal['database', 'data_table', 'text_table', 'index', 'report'], 
        dict[Literal['header', 'first_column', 'data', 'column_widths'], CellConfig]]:
        """Carga y almacena formatos de celdas para hojas que contienen datos (database e index)"""
        white_borders = {
            'border': 1,
            'border_color': Color.WHITE,
        }

        gray_borders = {
            'border': 1,
            'border_color': Color.GRAY,
        }
        gray_light_borders = {
            'border': 1,
            'border_color': Color.GRAY_LIGHT,
        }

        first_column_basic = {
            **white_borders,
            'bg_color': Color.GRAY_LIGHT,
            'text_wrap': True,
            'valign': 'vcenter',
            'font_size': 10
        }

        header_basic = {
            **white_borders,
            'bg_color': Color.BLUE_DARK,
            'font_color': Color.WHITE,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
        }

        return { 
            'database': {
                'header': {
                    **header_basic,
                    'font_size': 10
                },
                'first_column': {
                    **first_column_basic
                },
                'data': {
                    **gray_light_borders,
                    'valign': 'vcenter',
                    'font_size': 10
                },
                'column_widths':{
                    'A:A': 12
                }
            },
            'text_table': {
                'header': {
                    **header_basic,
                    'font_size': 11
                },
                'first_column': {
                    **first_column_basic,
                    'align': 'justify',
                },
                'data': {
                    **gray_borders,
                    'align': 'justify',
                    'valign': 'vcenter',
                    'text_wrap': True,
                    'font_size': 10
                },
                'column_widths': {
                    'A:A': 27,
                    'B:B': 57
                }
            },
            'data_table': {
                'header': {
                    **header_basic,
                    'font_size': 10
                },
                'first_column': {
                    **first_column_basic,
                    'align': 'center'
                },
                'data': {
                    'border': 1,
                    'border_color': Color.GRAY_LIGHT,
                    'align': 'right',
                    'valign': 'vcenter',
                    'text_wrap': True,
                    'font_size': 10
                },
                'column_widths':{
                    'A:A': 12,

                }
            },
            'index': {
                'header': {
                    **header_basic,
                    'font_size': 10
                },
                'first_column': {
                    **first_column_basic,
                    'align': 'justify'
                },
                'data': {
                    **gray_borders,
                    'align': 'justify',
                    'valign': 'vcenter',
                    'text_wrap': True,
                    'font_size': 10
                },
                'column_widths':{
                    'A:A': 10,
                    'B:C': 40,
                    'D:D': 15,
                    'E:G': 35
                },
            },
            'report': {
                'header': {
                    **header_basic,
                    'font_size': 11
                },
                'data': {
                    **gray_light_borders,
                    'valign': 'vcenter',
                    'font_size': 10,
                    'text_wrap': True,
                },
                'first_column': {
                    **white_borders,
                    'bg_color': Color.GRAY_LIGHT,
                    'text_wrap': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'font_size': 10
                }
            }
        }

# TODO: Test with column widths
# TODO: Add param for legend and adjust height if not legend
class ChartFormats:
    def __init__(self):
        self._line_colors = [
            Color.RED, 
            Color.BLUE_DARK, 
            Color.GREEN_DARK,
            Color.ORANGE,
            Color.BLUE,
            Color.GRAY,
            Color.YELLOW,
            Color.PURPLE, 
        ]
        self._line_simple_colors = [
            Color.RED_DARK, 
            Color.BLUE_DARK,
            Color.BLUE, 
        ]
        self._line_single_colors = [
            Color.BLUE, 
        ]
        self._column_colors = [
            Color.BLUE_DARK, 
            Color.RED_DARK, 
            Color.BLUE, 
            Color.GREEN_DARK, 
            Color.ORANGE, 
            Color.YELLOW, 
            Color.GRAY
        ]
        self._column_simple_colors = [
            Color.BLUE_DARK, 
            Color.RED_DARK, 
            Color.GRAY
        ]
        self._column_percent_stacked_colors = [
            Color.GREEN_DARK, 
            Color.BLUE, 
            Color.BLUE_DARK,
            Color.RED_DARK,
            Color.GRAY
        ]
        self._bar_colors = [
            Color.RED_DARK, 
            Color.BLUE_DARK, 
            Color.YELLOW
        ]
        self._cleveland_dot_colors = [
            Color.BLUE_DARK,
            Color.BLUE_LIGHT,
        ]


    @cached_property
    def charts(self) -> dict[Literal['basic', 'line', 'line_simple', 'line_single', 'line_monthly',
                                     'column', 'column_simple', 'column_single', 'column_stacked', 'bar', 'bar_single', 'cleveland_dot'], Any]:
        """
        Accede a los formatos para gráficos en un solo diccionario.
        """
        return {
            'basic': self._basic(),
            'line': self._line(),
            'line_simple': self._line_simple(),
            'line_single': self._line_single(),
            'line_monthly': self._line_monthly(),
            'column': self._column(),
            'column_simple': self._column_simple(),
            'column_single': self._column_single(),
            'column_stacked': self._column_stacked(),
            'bar': self._bar(),
            'bar_single': self._bar_single(),
            'bar_double': self._bar_double(),
            'cleveland_dot': self._cleveland_dot()
        }

    def _basic(self):
        return {
            'title': {'name': ''},
            'size': {'width': 600, 'height': 300},
            'legend': {'position': 'bottom'}   
            #     'layout': {
            #         'x': 0.20,
            #         'y': 0.93,
            #         'width': 0.60,
            #         'height': 0.05
            #     }
            # }
            ,
            'chartarea': {'border': {'none': True}},
            'plotarea': {
                'layout': {
                    'x': 0.06,
                    'y': 0.04,
                    'width': 0.91,
                    'height': 0.77
                }
            },
            'x_axis': {
                'name': '',
                'text_axis': True,
                'minor_tick_mark': 'outside',
                'major_tick_mark': 'none',
            },
            'y_axis': {
                'major_gridlines': {
                    'visible': True,
                    'line': {'color': Color.GRAY_GRIDS}
                }
            }
        }

    def _line(self):
        return {
            'colors': self._line_colors,
            'dash_type': [
                'solid', 'round_dot', 'round_dot', 'round_dot',
                'round_dot', 'round_dot', 'round_dot'
            ],
            'series': {
                'smooth': True,
                'line': {'width': 1.75},
                'marker': {'type': 'circle', 'size': 6},
                'data_labels': {'value': False}
            },
            'x_axis': {
                'minor_gridlines': {
                    'visible': True,
                    'line': {'color': Color.GRAY_GRIDS},
                }
            },
            'y_axis': {
                'min': 0,
                'major_gridlines': {
                    'visible': True,
                    'line': {'color': Color.GRAY_GRIDS},
                }
            }
        }

    def _line_simple(self):
        return {
            'colors': self._line_simple_colors,
            'dash_type': ['round_dot', 'square_dot', 'solid'],
            'series': {
                'smooth': True,
                'line': {'width': 1.75},
                'marker': {'type': 'circle', 'size': 6},
                'data_labels': {
                    'value': True,
                    'position': 'above',
                    'font': {
                        'size': 10
                    }
                }
            },
            'y_axis': {
                'major_gridlines': {
                    'visible': True,
                    'line': {'color': Color.GRAY_GRIDS}
                }
            },
            'x_axis': {
                'major_gridlines': {
                    'visible': True,
                    'line': {'color': Color.GRAY_GRIDS}
                }
            }
        }

    def _line_single(self):
        return {
            'colors': self._line_single_colors,
            'legend': {'none': True},
            'dash_type': ['square_dot', 'solid'],
            'series': {
                'smooth': True,
                'line': {'width': 2.5},
                'marker': {'type': 'diamond', 'size': 12},
                'data_labels': {
                    'value': True,
                    'position': 'above',
                    'fill': {'color': Color.BLUE_LIGHT},
                    'font': {
                        'bold': True,
                        'color': Color.BLACK,
                        'size': 10.5
                    },
                    'border': {
                        'width': 1
                    }
                }
            },
           'y_axis': {
                'visible': False,
                'major_gridlines': {
                    'visible': False,
                    'line': {'color': Color.GRAY_GRIDS}
                }
            }
        }

    def _line_monthly(self):
        return {
            'colors': self._line_simple_colors,
            'dash_type': ['solid', 'square_dot', 'round_dot'],
            'series': {
                'smooth': True,
                'line': {'width': 1.75},
                'marker': {'none': True},
                'data_labels': {'value': False}
            },
           'x_axis': {
                'major_gridlines': {
                    'visible': True,
                    'line': {'color': Color.GRAY_GRIDS}
                }
            }
        }
        

    def _column(self):
        return {
            'colors': self._column_colors,
            'series': {
                'gap': 60,
                'data_labels': {
                    'position': 'outside_end',
                    'font': {
                        'bold': True,
                        'color': Color.BLACK,
                        'size': 10
                    }
                }
            }
        }

    def _column_simple(self):
        return {
            'colors': self._column_simple_colors,
            'series': {
                'gap': 60,
                'data_labels': {
                    'position': 'outside_end',
                    'font': {
                        'bold': True,
                        'color': Color.BLACK,
                        'size': 10.5
                    }
                }
            },
            'y_axis': {
                'visible': False,
                'reverse': False,
            },
            'x_axis': {
                'major_gridlines': {
                    'visible': False,
                    }
            }
        }
    
    def _column_single(self):
        return {
            'colors': self._column_simple_colors,
            'legend': {'none': True},
            'series': {
                'gap': 60,
                'data_labels': {
                    'position': 'outside_end',
                    'font': {
                        'bold': True,
                        'color': Color.BLACK,
                        'size': 10.5
                    }
                }
            },
            'y_axis': {
                'visible': False,
                'reverse': False,
                'major_gridlines': {
                    'visible': False,
                    },
            },
            'x_axis': {
                'major_gridlines': {
                    'visible': False,
                    }
            }
        }
    # Y max should be 100
    def _column_stacked(self):
        return {
            'size': {'width': 600, 'height': 320},
            'colors': self._column_percent_stacked_colors,
            'series': {
                'gap': 50,
                'data_labels': {
                    'position': 'outside_end',
                    'font': {
                        'bold': True,
                        'color': Color.WHITE,
                        'size': 9
                    }
                }
            },
            'x_axis': {
                'minor_tick_mark': 'none',
                'major_tick_mark': 'outside',
                'major_gridlines': {
                    'visible': False
                },
                'line': {
                    'width': 1,
                    'color': Color.GRAY
               }
            },
            'y_axis': {
                'minor_tick_mark': 'none',
                # 'major_gridlines': {
                #     'visible': False,
                # },
                'line': {
                    'none': True
                }
            }
        }

    def _bar(self):
        return {
            'size': {'width': 600, 'height': 310},
            'colors': self._bar_colors,
            'series': {
                'gap': 60,
                'data_labels': {
                    'value': True,
                    'position': 'outside_end',
                    'font': {
                        'bold': True,
                        'color': Color.BLACK,
                        'size': 10
                    }
                }
            },
            'x_axis': {
                'visible': False,
                'line': {'none': True},
                'major_tick_mark': 'none',
                'major_gridlines': {'visible': False}
            },
            # 'y_axis': {
            #     'reverse': True,
            # }
        }

    def _bar_single(self):
        return {
            'size': {'width': 600, 'height': 450},
            'legend': {'none': True},
            'colors': self._column_simple_colors,
            'series': {
                'gap': 40,
                'data_labels': {
                    'value': True,
                    'position': 'outside_end',
                    'font': {
                        'bold': True,
                        'color': Color.BLACK,
                        'size': 10
                    }
                }
            },
            'x_axis': {
                'visible': False,
                'line': {'none': True},
                'major_tick_mark': 'none',
                'major_gridlines': {'visible': False}
            }
        }
    
    def _bar_double(self):
        return {
            'size': {'width': 600, 'height': 450},
            'colors': self._cleveland_dot_colors,
            'series': {
                'gap': 40,
                'data_labels': {
                    'value': True,
                    'position': 'inside_end', # o inside_base
                    'font': {
                        'bold': False,
                        'color': Color.WHITE,
                        'size': 10
                    }
                }
            },
            'x_axis': {
                'visible': False,
                'line': {'none': True},
                'major_tick_mark': 'none',
                'major_gridlines': {'visible': False}
            }
        }

    def _cleveland_dot(self):
        return {
            'colors': self._cleveland_dot_colors,
            'size': {'width': 600, 'height': 450},
            'legend': {'delete_series': [-1]},
            'x_error_bars':{
                'end_style': 0,
                'direction': 'plus',
                'type': 'custom',
                'line': {'width': 4, 'color': Color.GRAY},
            },
            'series': {
                'marker': {'type': 'circle', 'size': 8},
                'data_labels': {'value': False}
            },
            'series': {
                'data_labels': {
                    'category': True,
                    'value': False,
                    'position': 'right',
                    'font': {
                        'bold': True,
                        'color': Color.BLACK,
                        'size': 10
                    }
                }
            },
            'y_axis': {
                'visible': False,
                'reverse': False,
                'major_gridlines': {'visible': False},
            },
            'x_axis': {
                'visible': False,
                'line': {'none': True},
                'major_tick_mark': 'none',
                'major_gridlines': {'visible': False}
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
