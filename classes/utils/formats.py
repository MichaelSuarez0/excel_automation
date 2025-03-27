from excel_automation.classes.utils.colors import Color
from typing import Any, Literal
from functools import cached_property
from pydantic import BaseModel, Field
from enum import Enum

class Alignment(str, Enum):
    left = 'left'
    right = 'right'
    center = 'center'
    justify = 'justify'


class CellConfig(BaseModel):
    bg_color: str
    font_color: str
    font_size: int = Field(..., gt=0, description="El tamaño de la fuente debe ser mayor que 0")
    bold: bool = False
    align: Alignment
    valign: Alignment
    num_format: str
    border: int = Field(..., ge=0, description="El grosor del borde debe ser 0 o mayor")
    border_color: str
    text_wrap: bool = False


class Formats(BaseModel):
    @cached_property
    def numeric_types(self) -> dict[Literal['date', 'integer', 'decimal_1', 'decimal_2', 'percentage'], str]:
        return NumericTypes().numeric_types

    @cached_property
    def cells(self) -> dict[Literal['database', 'index', 'data_table', 'text_table', 'report'], dict[Literal['header', 'first_column', 'data'], CellConfig]]:
        return CellFormats().cells

    @cached_property
    def charts(self) -> dict[
        Literal[
            'line', 'line_simple', 'line_single', 'line_monthly', 'column', 'column_simple', 'column_single', 'column_stacked', 'bar', 
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
    

# TODO: Return cycle para las listas (por ejemplo colores) y llamar al dict con next
class CellFormats:
    @cached_property
    def cells(self) -> dict[
        Literal['database', 'data_table', 'text_table', 'index', 'report'], 
        dict[Literal['header', 'first_column', 'data'], CellConfig]]:
        """Carga y almacena formatos de celdas para hojas que contienen datos (database e index)"""
        white_borders = {
            'border': 1,
            'border_color': Color.WHITE,
        }

        return { 
            'database': {
                'header': {
                    **white_borders,
                    'bg_color': Color.BLUE_DARK,
                    'font_color': Color.WHITE,
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True,
                    'font_size': 10
                },
                'first_column': {
                    **white_borders,
                    'bg_color': Color.GRAY_LIGHT,
                    'text_wrap': True,
                    'font_size': 10
                },
                'data': {
                    'border': 1,
                    'border_color': Color.GRAY_LIGHT,
                    'valign': 'vcenter',
                    'font_size': 10
                }
            },
            'text_table': {
                'header': {
                    **white_borders,
                    'bg_color': Color.BLUE_DARK,
                    'font_color': Color.WHITE,
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True,
                    'font_size': 10
                },
                'first_column': {
                    **white_borders,
                    'bg_color': Color.GRAY_LIGHT,
                    'valign': 'vcenter',
                    'align': 'justify',
                    'text_wrap': True,
                    'font_size': 10
                },
                'data': {
                    'border': 1,
                    'border_color': Color.GRAY_LIGHT,
                    'align': 'justify',
                    'valign': 'vcenter',
                    'text_wrap': True,
                    'font_size': 10
                }
            },
            'data_table': {
                'header': {
                    **white_borders,
                    'bg_color': Color.BLUE_DARK,
                    'font_color': Color.WHITE,
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True,
                    'font_size': 10
                },
                'first_column': {
                    **white_borders,
                    'bg_color': Color.GRAY_LIGHT,
                    'text_wrap': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'font_size': 10
                },
                'data': {
                    'border': 1,
                    'border_color': Color.GRAY_LIGHT,
                    'align': 'right',
                    'valign': 'vcenter',
                    'font_size': 10
                }
            },
            'index': {
                'header': {
                    **white_borders,
                    'bg_color': Color.BLUE_DARK,
                    'font_color': Color.WHITE,
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'text_wrap': True,
                    'font_size': 10
                },
                'first_column': {
                    **white_borders,
                    'bg_color': Color.GRAY_LIGHT,
                    'valign': 'vcenter',
                    'align': 'justify',
                    'text_wrap': True,
                    'font_size': 10
                },
                'data': {
                    'border': 1,
                    'border_color': Color.GRAY_LIGHT,
                    'align': 'justify',
                    'valign': 'vcenter',
                    'text_wrap': True,
                    'font_size': 10
                }
            },
            'report': {
                'header': {
                    'bold': True,
                    'valign': 'vcenter',
                    'text_wrap': False,
                    'font_size': 14
                },
                'data': {
                    'valign': 'vcenter',
                    'font_size': 10
                }
            },
        }

# TODO: Test with column widths
# TODO: Add param for legend and adjust height if not legend
class ChartFormats:
    def __init__(self):
        self._line_colors = [
            Color.BLUE_DARK, 
            Color.RED, 
            Color.ORANGE, 
            Color.GREEN_DARK, 
            Color.PURPLE, 
            Color.GRAY
        ]
        self._line_simple_colors = [
            Color.RED, 
            Color.BLUE_DARK,
            Color.BLUE, 
        ]
        self._line_single_colors = [
            Color.BLUE, 
        ]
        self._column_colors = [
            Color.BLUE_DARK, 
            Color.RED, 
            Color.BLUE, 
            Color.GREEN_DARK, 
            Color.ORANGE, 
            Color.YELLOW, 
            Color.GRAY
        ]
        self._column_simple_colors = [
            Color.BLUE_DARK, 
            Color.RED, 
            Color.GRAY
        ]
        self._column_percent_stacked_colors = [
            Color.GREEN_DARK, 
            Color.BLUE, 
            Color.BLUE_DARK,
            Color.RED,
            Color.GRAY
        ]
        self._bar_colors = [
            Color.RED, 
            Color.BLUE_DARK, 
            Color.YELLOW
        ]

    @cached_property
    def charts(self) -> dict[Literal['basic', 'line', 'line_simple', 'line_single', 'line_monthly', 'column', 'column_simple', 'column_single', 'column_stacked', 'bar', 'bar_single'], Any]:
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
        }

    def _basic(self):
        return {
            'title': {'name': ''},
            'size': {'width': 580, 'height': 300},
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
                    'line': {'color': Color.GRAY_LIGHT}
                }
            }
        }

    def _line(self):
        return {
            'colors': self._line_colors,
            'dash_type': [
                'square_dot', 'round_dot', 'round_dot', 'round_dot',
                'round_dot', 'round_dot', 'round_dot'
            ],
            'series': {
                'smooth': True,
                'line': {'width': 1.75},
                'marker': {'type': 'circle', 'size': 6},
                'data_labels': {'value': False}
            },
            'x_axis': {
                'mayor_gridlines': {
                    'visible': True,
                    'line': {'color': Color.GRAY_LIGHT}
                }
            },
            'y_axis': {'min': 0}
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
                    'line': {'color': Color.GRAY_LIGHT}
                }
            },
            'x_axis': {
                'major_gridlines': {
                    'visible': True,
                    'line': {'color': Color.GRAY_LIGHT}
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
                    'line': {'color': Color.GRAY_LIGHT}
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
                    'line': {'color': Color.GRAY_LIGHT}
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
            'size': {'width': 580, 'height': 320},
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
                'major_gridlines': {
                    'visible': False
                },
                'line': {
                    'width': 1,
                    'color': Color.GRAY_LIGHT
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
            'size': {'width': 570, 'height': 310},
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
            'y_axis': {
                'reverse': True,
            }
        }

    def _bar_single(self):
        return {
            'size': {'width': 570, 'height': 450},
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
