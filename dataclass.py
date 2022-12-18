from typing import Optional, Union
from enums import ThemeColorType, TextFontFamily
from pydantic import BaseModel



def _to_camel_case(snake_case_text: str) -> str:
    components = snake_case_text.split('_')
    return components[0] + ''.join(x.title() for x in components[1:])


class DataClassSupport(BaseModel):
    class Config:
        alias_generator = _to_camel_case
        allow_population_by_field_name = True
        use_enum_values = True


class CellPadding(DataClassSupport):
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


class TextFormat(DataClassSupport):
    foreground_color_style: Optional[
        Union[ColorStyleRgbColor, ColorStyleThemeColor]] = None
    font_family: Optional[TextFontFamily] = None
    font_size: Optional[int] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    strikethrough: Optional[bool] = None
    underline: Optional[bool] = None
    link: Optional[Link] = None

print(
    TextFormat(
        font_family=TextFontFamily.ARIAL,
        link=Link(uri="http://www.google.com"),
        foreground_color_style=ColorStyleThemeColor(theme_color=ThemeColorType.ACCENT1)
    ).dict(by_alias=True, exclude_none=True)
)
