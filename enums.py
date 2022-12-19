from enum import Enum


class BorderStyle(Enum):
    STYLE_UNSPECIFIED = "STYLE_UNSPECIFIED"
    DOTTED = "DOTTED"
    DASHED = "DASHED"
    SOLID = "SOLID"
    SOLID_MEDIUM = "SOLID_MEDIUM"
    SOLID_THICK = "SOLID_THICK"
    NONE = "NONE"
    DOUBLE = "DOUBLE"


class ThemeColorType(Enum):
    THEME_COLOR_TYPE_UNSPECIFIED = "THEME_COLOR_TYPE_UNSPECIFIED"
    TEXT = "TEXT"
    BACKGROUND = "BACKGROUND"
    ACCENT1 = "ACCENT1"
    ACCENT2 = "ACCENT2"
    ACCENT3 = "ACCENT3"
    ACCENT4 = "ACCENT4"
    ACCENT5 = "ACCENT5"
    ACCENT6 = "ACCENT6"
    LINK = "LINK"


class CellVerticalAlign(Enum):
    VERTICAL_ALIGN_UNSPECIFIED = "VERTICAL_ALIGN_UNSPECIFIED"
    TOP = "TOP"
    MIDDLE = "MIDDLE"
    BOTTOM = "BOTTOM"


class CellHorizontalAlign(Enum):
    HORIZONTAL_ALIGN_UNSPECIFIED = "HORIZONTAL_ALIGN_UNSPECIFIED"
    LEFT = "LEFT"
    CENTER = "CENTER"
    RIGHT = "RIGHT"


class CellWrapStrategy(Enum):
    WRAP_STRATEGY_UNSPECIFIED = "WRAP_STRATEGY_UNSPECIFIED"
    OVERFLOW_CELL = "OVERFLOW_CELL"
    LEGACY_WRAP = "LEGACY_WRAP"
    CLIP = "CLIP"
    WRAP = "WRAP"


class CellTextDirection(Enum):
    TEXT_DIRECTION_UNSPECIFIED = "TEXT_DIRECTION_UNSPECIFIED"
    LEFT_TO_RIGHT = "LEFT_TO_RIGHT"
    RIGHT_TO_LEFT = "RIGHT_TO_LEFT"


class CellHyperlinkDisplayType(Enum):
    HYPERLINK_DISPLAY_TYPE_UNSPECIFIED = "HYPERLINK_DISPLAY_TYPE_UNSPECIFIED"
    LINKED = "LINKED"
    PLAIN_TEXT = "PLAIN_TEXT"


class TextFontFamily(Enum):
    DEFAULT = "Arial"
    AMATIC_SC = "Amatic SC"
    ARIAL = "Arial"
    CAVEAT = "Caveat"
    COMFORTAA = "Comfortaa"
    COMIC_SANS_MS = "Comic Sans MS"
    COURIER_NEW = "Courier New"
    EB_GARAMOND = "EB Garamond"
    GEORGIA = "Georgia"
    IMPACT = "Impact"
    LEXEND = "Lexend"
    LOBSTER = "Lobster"
    LORA = "Lora"
    MERRIWEATHER = "Merriweather"
    MONTSERRAT = "Montserrat"
    NUNITO = "Nunito"
    OSWALD = "Oswald"
    PACIFICO = "Pacifico"
    PLAYFAIR_DISPLAY = "Playfair Display"
    ROBOTO = "Roboto"
    ROBOTO_MONO = "Roboto Mono"
    ROBOTO_SERIF = "Roboto Serif"
    SPECTRAL = "Spectral"
    TIMES_NEW_ROMAN = "Times New Roman"
    TREBUCHET_MS = "Trebuchet MS"
    VERDANA = "Verdana"


class NumberFormatType(Enum):
    NUMBER_FORMAT_TYPE_UNSPECIFIED = "NUMBER_FORMAT_TYPE_UNSPECIFIED"
    TEXT = "TEXT"
    NUMBER = "NUMBER"
    PERCENT = "PERCENT"
    CURRENCY = "CURRENCY"
    DATE = "DATE"
    TIME = "TIME"
    DATE_TIME = "DATE_TIME"
    SCIENTIFIC = "SCIENTIFIC"
