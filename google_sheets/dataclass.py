from __future__ import annotations
from typing import Optional, Union
from pydantic import BaseModel, validator

from .enums import (
    BorderStyle,
    CellHorizontalAlign,
    CellHyperlinkDisplayType, 
    CellTextDirection, 
    CellVerticalAlign, 
    CellWrapStrategy,
    NumberFormatType, 
    ThemeColorType,
    TextFontFamily
)

from .support import Color



def _to_camel_case(snake_case_text: str) -> str:
    components = snake_case_text.split('_')
    return components[0] + ''.join(x.title() for x in components[1:])


class DataClassSupport(BaseModel):

    class Config:
        alias_generator = _to_camel_case
        populate_by_name = True
        use_enum_values = True
        arbitrary_types_allowed = True


    def fields(self, base_name: str, _dict: Optional[dict] = None) -> str:
        if _dict is None:
            _dict = self.dict(by_alias=True)

        base_name = (base_name.rstrip(".") + ".")

        if not len(_dict):
            return base_name.rstrip(".")
        
        return ",".join([f"{base_name}{k}" for k in _dict.keys()])

    
    @staticmethod
    def color_to_style(color: Color) -> ColorStyleRgbColor:
        return ColorStyleRgbColor(rgb_color=ColorStyleColor(**color.dict()))



class Padding(DataClassSupport):
    top: int
    right: int
    bottom: int
    left: int


class BorderDirection(DataClassSupport):
    top: bool = True
    bottom: bool = True
    left: bool = True
    right: bool = True


class ColorStyleColor(DataClassSupport):
    red: float
    green: float
    blue: float
    alpha: float = 1.0


class ColorStyleRgbColor(DataClassSupport):
    rgb_color: ColorStyleColor
    

class ColorStyleThemeColor(DataClassSupport):
    theme_color: ThemeColorType


class Link(DataClassSupport):
    uri: str


class NumberFormat(DataClassSupport):
    """
    pattern examples

    Date & time format:
    - h:mm:ss.00 a/p 	     4:08:53.53 p
    - hh:mm A/P".M." 	     04:08 P.M.
    - yyyy-mm-dd 	         2016-04-05
    - mmmm d \[dddd\] 	     April 5 [Tuesday]
    - h PM, ddd mmm dd 	     4 PM, Tue Apr 05
    - dddd, m/d/yy at h:mm   Tuesday, 4/5/16 at 16:08
    - [hh]:[mm]:[ss].000 	 03:13:41.255
    - [mmmm]:[ss].000 	     0193:41.255

    Number format:
    Number 	       Pattern      Formatted Value
    - 12345.125    ####.#       12345.1
    - 12.34 	   000.0000     012.3400
    - 12 	       #.0# 	    12.0
    - 5.125 	   # ???/???    5 1/8
    - 12000        #,###        12,000
    - 1230000 	   0.0,,"M"     1.2M
    - 1234500000   0.00e+00     1.23e+09
    """
    type: NumberFormatType
    pattern: Optional[str]


class TextRotation(DataClassSupport):
    angle: float
    vertical: Optional[bool] = False


class TextFormat(DataClassSupport):
    foreground_color_style: Optional[Union[Color, ThemeColorType, ColorStyleRgbColor, ColorStyleThemeColor]] = None
    font_family: Optional[TextFontFamily] = None
    font_size: Optional[int] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    strikethrough: Optional[bool] = None
    underline: Optional[bool] = None
    link: Optional[Link] = None

    @validator("*", pre=True)
    def convert_theme_color_type(cls, value):
        if type(value) is ThemeColorType:
            return ColorStyleThemeColor(theme_color=value)
                    
        if isinstance(value, Color):
            return cls.color_to_style(value)
        
        return value


class CellFormat(DataClassSupport):
    number_format: Optional[NumberFormat] = None
    background_color_style: Optional[Union[Color, ThemeColorType, ColorStyleRgbColor, ColorStyleThemeColor]] = None
    padding: Optional[Padding] = None
    horizontal_alignment: Optional[CellHorizontalAlign] = None
    vertical_alignment: Optional[CellVerticalAlign] = None
    wrap_strategy: Optional[CellWrapStrategy] = None
    text_direction: Optional[CellTextDirection] = None
    hyperlink_display_type: Optional[CellHyperlinkDisplayType]
    text_rotation: Optional[TextRotation]
    
    @validator("*", pre=True)
    def convert_theme_color_type(cls, value):
        if type(value) is ThemeColorType:
            return ColorStyleThemeColor(theme_color=value)
        
        if isinstance(value, Color):
            return cls.color_to_style(value)
        
        return value


class Border(DataClassSupport):
    style: BorderStyle = None
    color_style: Union[Color, ThemeColorType, ColorStyleRgbColor, ColorStyleThemeColor] = None

    @validator("*", pre=True)
    def convert_theme_color_type(cls, value):
        if type(value) is ThemeColorType:
            return ColorStyleThemeColor(theme_color=value)
        
        if isinstance(value, Color):
            return cls.color_to_style(value)
        
        return value


class BorderClassicFormat(DataClassSupport):
    style: BorderStyle = BorderStyle.SOLID
    color: Union[Color, ThemeColorType, ColorStyleRgbColor, ColorStyleThemeColor] = Color(0, 0, 0)
    direction: BorderDirection = BorderDirection()

    @validator("*", pre=True)
    def convert_theme_color_type(cls, value):
        if type(value) is ThemeColorType:
            return ColorStyleThemeColor(theme_color=value)
        
        if isinstance(value, Color):
            return cls.color_to_style(value)
        
        return value
    
    def to_format(self) -> BorderFormat:
        return BorderFormat(
            style=self.style,
            color=self.color,
            direction=self.direction
        )


class BorderFormat(DataClassSupport):
    top: Optional[Border]
    bottom: Optional[Border]
    left: Optional[Border]
    right: Optional[Border]
    innerHorizontal: Optional[Border]
    innerVertical: Optional[Border]
