from enum import Enum

#TODO: Un poco más azul el blue_dark
class Color(Enum):
    RED: str = "#C80000"
    RED_LIGHT: str = "#FFABAB"
    BLUE_LIGHT: str = "#A6CAEC"
    BLUE: str = "#3B79D5"
    BLUE_DARK: str = "#152747"
    GREEN_DARK: str = "#008E2C"
    GRAY_LIGHT: str = "#ebebeb"
    GRAY: str = "#B8B8B8"
    YELLOW: str = "#FFC000"
    WHITE: str = '#FFFFFF'
    ORANGE: str = "#DD6909"
    PURPLE: str = "#7030A0"

    @property
    def no_hash(self):
        return self.value.lstrip("#")
    