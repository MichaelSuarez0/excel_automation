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
    
    @property
    def rgb(self):
        hex_color = self.value.lstrip("#")
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

    @property
    def win32(self):
        # Remove the '#' if present
        hex_color = self.value.lstrip('#')
        # Convert the hex values to integers
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        # Rearrange to BGR and return the combined integer
        return (b << 16) + (g << 8) + r
    