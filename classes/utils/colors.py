from enum import Enum

# TODO: DIFFERENT GRAY FOR BORDERS AND GRIDLINES
# TODO: Para grids más light
# TODO: UN poco más oscuro blue_light
class Color(str, Enum):
    RED_DARK: str = "#be0000"
    RED: str = '#d41225'
    RED_LIGHT: str = "#FFABAB"
    BLUE_LIGHT: str = "#6499DA"
    BLUE: str = "#0060a4"
    BLUE_DARK: str = "#00194b"
    GREEN_DARK: str = "#007630"
    GREEN_LIGHT: str = "#00B050"
    GRAY_LIGHT: str = "#E8E8E8"
    GRAY_GRIDS: str = "#F0F0F0"
    GRAY: str = "#a5a5a5"
    YELLOW: str = "#FEC200"
    WHITE: str = '#FFFFFF'
    ORANGE: str = "#DD6909"
    PURPLE: str = "#7030A0"
    BLACK: str = "#000000"
    BROWN: str = "#7A3D00"

    def __str__(self):
        return self.value 

    @property
    def no_hash(self):
        return self.value.lstrip("#")
    
    @property
    def rgb(self):
        hex_color = self.value.lstrip("#")
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

    @property
    def bgr(self):
        # Remove the '#' if present
        hex_color = self.value.lstrip('#')
        # Convert the hex values to integers
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        # Rearrange to BGR and return the combined integer
        return (b << 16) + (g << 8) + r
    