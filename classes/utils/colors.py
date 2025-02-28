from enum import Enum

#TODO: Un poco más azul el blue_dark
class Color(Enum):
    RED: str = "#be0000"
    RED_LIGHT: str = "#FFABAB"
    BLUE_LIGHT: str = "#A6CAEC"
    BLUE: str = "#0060a4"
    BLUE_DARK: str = "#00194b"
    GREEN_DARK: str = "#008236"
    GRAY_LIGHT: str = "#ebebeb"
    GRAY: str = "#a5a5a5"
    YELLOW: str = "#FFC000"
    WHITE: str = '#FFFFFF'
    ORANGE: str = "#DD6909"
    PURPLE: str = "#7030A0"
    BLACK: str = "#000000"

    @property
    def no_hash(self):
        return self.value.lstrip("#")
    