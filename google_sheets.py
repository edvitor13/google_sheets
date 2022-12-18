from __future__ import annotations
from typing import Any, Optional, overload
import os.path
import re
from copy import deepcopy, copy

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build, Resource
from googleapiclient.errors import HttpError

from enums import (
    BorderStyle,
    ThemeColorType,
    CellVerticalAlign,
    CellHorizontalAlign,
    CellWrapStrategy,
    CellTextDirection,
    TextFontFamily,
    NumberFormatType
)

from exceptions import (
    EmptyRangeException,
    InvalidRangeException,
    InvalidSheetException
)

from dataclass import (
    BorderDirection,
    CellPadding,
    BorderDirection,
    ColorStyleColor,
    ColorStyleRgbColor,
    ColorStyleThemeColor,
    Link,
    TextFormat
)



class Color():

    def __init__(
        self,
        red: float, 
        green: float, 
        blue: float, 
        alpha: float = 1.0
    ) -> None:
        self.red = self._check_color_range(red)
        self.green = self._check_color_range(green)
        self.blue = self._check_color_range(blue)
        self.alpha = self._check_color_range(alpha)

    
    def _check_color_range(self, color: float) -> float:
        if color < 0 or color > 1:
            raise ValueError("Invalid color range, supported range is 0 to 1")
        return color

    
    def rgb(self) -> tuple[float, float, float]:
        return (self.red, self.green, self.blue)


    def rgba(self) -> tuple[float, float, float, float]:
        return (self.red, self.green, self.blue, self.alpha)


    def dict(self) -> dict:
        return {
            "red": self.red,
            "green": self.green,
            "blue": self.blue,
            "alpha": self.alpha
        }
    

    def to_color_style(self) -> ColorStyleRgbColor:
        return ColorStyleRgbColor(rgb_color=ColorStyleColor(**self.dict()))


    def __str__(self) -> str:
        return f"Color{self.rgba()}"


    def __repr__(self) -> str:
        return self.__str__()


class ColorRGBA(Color):

    def __init__(
        self, 
        red: int, 
        green: int, 
        blue: int, 
        alpha: float = 1.0
    ) -> None:
        super().__init__(
            self.__check_convert_color_range(red), 
            self.__check_convert_color_range(green), 
            self.__check_convert_color_range(blue), 
            alpha
        )


    def __check_convert_color_range(self, color: int) -> float:
        if color < 0 or color > 255:
            raise ValueError("Invalid color range, supported range is 0 to 255")
        return float(color / 255)


class Range():

    last_range: Range | None = None

    _sheetname: str | None
    _start_row: int | None
    _start_col: int | None
    _end_row: int | None
    _end_col: int | None

    @overload
    def __init__(self, range: str | Range):
        ...

    @overload
    def __init__(
        self, sheetname: str | None, start: str, end: str | None):
        ...

    @overload
    def __init__(
        self, 
        sheetname: str | None, 
        start_col: int | None,
        start_row: int | None, 
        end_col: int | None,
        end_row: int | None
    ) -> None:
        ...

    
    def __init__(self, *args) -> None:
        match len(args):
            case 1:
                self._load_by_range(*args)
            case 3:
                self._load_by_detached_range(args[0], args[1], args[2])
            case 5:
                self._load_by_number(*args)
            case _:
                raise EmptyRangeException(
                    "A Range is required, "
                    "use the example structure: PageName!A6:F10")
            
        self.__consistency_adjust()
        Range.last_range = self

    
    def copy(self, _deepcopy: bool = False) -> Range:
        if _deepcopy:
            return deepcopy(self)
        return copy(self)


    def _load_by_range(self, _range: str | Range) -> Range:
        
        if isinstance(_range, Range):
            self._sheetname = _range._sheetname
            self._start_row = _range._start_row
            self._start_col = _range._start_col
            self._end_row = _range._end_row
            self._end_col = _range._end_col

            return self

        extract: list[str | None] = self._extract_range(_range)

        if extract is None or len(extract) != 5:
            raise InvalidRangeException(
                f"Range '{_range}' in invalid format, "
                "use the example structure: PageName!A6:F10")
        
        self._sheetname = extract[0]
        self._start_col = self._get_column_number(extract[1])
        self._start_row = self._get_row_number(extract[2])
        self._end_col = self._get_column_number(extract[3])
        self._end_row = self._get_row_number(extract[4])

        return self

    
    def _load_by_detached_range(
        self, 
        sheetname: str | None, 
        start: str, 
        end: str | None = None
    ) -> Range:
        range = "" if sheetname is None else f"{sheetname}!"
        range += f"{start}"
        range += "" if end is None else f":{end}"

        return self._load_by_range(range)

    
    def _load_by_number(
        self, 
        sheetname: str | None, 
        start_col: int | None,
        start_row: int | None, 
        end_col: int | None,
        end_row: int | None,
    ) -> Range:
        self._sheetname = sheetname
        self._start_row = start_row
        self._start_col = start_col
        self._end_row = end_row
        self._end_col = end_col

        return self

    
    def _extract_range(self, _range: str) -> list | None:
        regexs = (
            r'^(.+!)?([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)$',
            r'|^(.+!)?([A-Z]+)([0-9]+):([A-Z]+)()$',
            r'|^(.+!)?([A-Z]+)([0-9]+)()()$'
        )

        findall = re.findall("".join(regexs), _range)

        if len(findall) == 0:
            return None

        find = tuple(map(lambda r: r if r else None, findall[0]))

        result: list[str | None] = [None, None, None, None, None]
        for i in range(len(regexs)):
            for j in range(len(result)):
                try:
                    if result[j] is None:
                        result[j] = find[j + len(result) * i]
                except IndexError:
                    result[j] = None
        
        if result[0] is not None and result[0][-1] == "!":
            result[0] = result[0][:-1]

        return result

    
    def __consistency_adjust(self) -> Range:
        if self._start_row is None or self._start_row < 0:
            self._start_row = 1
        
        if self._end_row is not None and self._end_row < self._start_row:
            self._end_row = self._start_row
        
        if self._start_col is None or self._start_col < 0:
            self._start_col = 1
        
        if self._end_col is not None and self._end_col < self._start_col:
            self._end_col = self._start_col

        return self


    def _get_column_number(self, name: str | None) -> int | None:
        if name is None:
            return None

        find = re.findall("[A-Z]", name)
        
        if len(find) < 1:
            return None

        index = 0
        for letter in find:
            index = index * 26 + 1 + ord(letter) - ord('A')

        return (index)

    
    def _get_row_number(self, number: str | None) -> int | None:
        if number is None:
            return None
        
        find = re.findall("[0-9]+", number)
        
        if len(find) < 1:
            return None 
        
        return int(find[0])


    def _get_column_name(self, number: int | None) -> str | None:
        if number is None:
            return None

        result = ''

        while number > 0:
            index = (number - 1) % 26
            result += chr(index + ord('A'))
            number = (number - 1) // 26

        if result == '':
            return None

        return result[::-1]


    def _calc_row_col(
        self, 
        attr: str,
        *,
        add: int | None = None,
        sub: int | None = None,
        change: int | None = None
    ) -> Range | int | None:
        val = getattr(self, attr)

        if (
            add is None 
            and sub is None
            and change is None
        ):
            return self
        
        if val is None:
            val = 0
        
        if add is not None:
            val += add
        
        if sub is not None:
            val -= sub

        if change is not None:
            val = change

        setattr(self, attr, val)

        self.__consistency_adjust()

        return self


    def add_start_row(self, val: int | None = None) -> Range:
        return self._calc_row_col('_start_row', add=val)
    
    def sub_start_row(self, val: int | None = None) -> Range:
        return self._calc_row_col('_start_row', sub=val)
    
    def set_start_row(self, val: int | None = None) -> Range:
        return self._calc_row_col('_start_row', change=val)
    
    def get_start_row(self) -> int:
        return self._start_row


    def add_start_col(self, val: int | None = None) -> Range:
        return self._calc_row_col('_start_col', add=val)
    
    def sub_start_col(self, val: int | None = None) -> Range:
        return self._calc_row_col('_start_col', sub=val)
    
    def set_start_col(self, val: int | None = None) -> Range:
        return self._calc_row_col('_start_col', change=val)
    
    def get_start_col(self) -> int:
        return self._start_col
    
    def get_start_letter(self) -> str:
        return self._get_column_name(self.get_start_col())

    def get_start(self) -> str:
        return f"{self.get_start_letter()}{self.get_start_row()}"


    def add_end_row(self, val: int | None = None) -> Range:
        return self._calc_row_col('_end_row', add=val)
    
    def sub_end_row(self, val: int | None = None) -> Range:
        return self._calc_row_col('_end_row', sub=val)
    
    def set_end_row(self, val: int | None = None) -> Range:
        return self._calc_row_col('_end_row', change=val)
    
    def get_end_row(self) -> int:
        return self._end_row


    def add_end_col(self, val: int | None = None) -> Range:
        return self._calc_row_col('_end_col', add=val)
    
    def sub_end_col(self, val: int | None = None) -> Range:
        return self._calc_row_col('_end_col', sub=val)
    
    def set_end_col(self, val: int | None = None) -> Range:
        return self._calc_row_col('_end_col', change=val)
    
    def get_end_col(self) -> int | None:
        return self._end_col
    
    def get_end_letter(self) -> str | None:
        return self._get_column_name(self.get_end_col())

    def get_end(self) -> str | None:
        if self.get_end_letter() is None and self.get_end_row() is None:
            return None
        
        if self.get_end_row() is None:
            return self.get_end_letter()

        return f"{self.get_end_letter()}{self.get_end_row()}"
    

    def set_sheetname(self, sheetname: str | None) -> Range:
        self._sheetname = sheetname
        return self
    
    def get_sheetname(self) -> str | None:
        return self._sheetname

    
    def get_range(self) -> str:
        _range = \
            "" if self.get_sheetname() is None else f"{self.get_sheetname()}!"
        _range += self.get_start()
        _range += "" if self.get_end() is None else f":{self.get_end()}"
            
        return _range


    def get_google_sheet_api_format(
        self, google_sheets: GoogleSheets | None = None
    ) -> dict[str, Any]:

        sheetpage_id = None
        if google_sheets is not None:
            sheetpage_id = google_sheets.sheetpage_id_by_name(
                self.get_sheetname())

        start_row = (self.get_start_row() - 1)
        start_col = (self.get_start_col() - 1)
        end_row = self.get_end_row()
        end_col = self.get_end_col()
        
        result = {
            "sheetId": sheetpage_id,
            "startRowIndex": start_row,
            "startColumnIndex": start_col
        }

        if end_row is not None:
            result["endRowIndex"] = end_row

        if end_col is not None:
            result["endColumnIndex"] = end_col
        
        return result


    def get_dict(self) -> dict[str, str | int]:
        return {
            "sheetname": self.get_sheetname(), 
            "start_row": self.get_start_row(), 
            "start_col": self.get_start_col(),
            "end_row": self.get_end_row(),
            "end_col": self.get_end_col()
        }


    def __str__(self) -> str:
        return self.get_range()

    
    def __repr__(self) -> str:
        return f"Range<{self.__str__()}>"



class GoogleSheets:

    def __init__(self, sheet_id: str) -> None:
        self.sheet_id: str = sheet_id
        self.service: Resource = self._load_service()

        self.__multi_execute_mode: bool = False
        self.__multi_batch_updates: list[dict] = []
        self.__multi_batch_values_updates: list[dict] = []
        
        if not self.is_valid_sheet():
            raise InvalidSheetException(
                f"ID table '{self.sheet_id}' is invalid")


    def _load_service(self) -> Optional[Resource]:
        creds = None

        SCOPES = [
            'https://www.googleapis.com/auth/spreadsheets.readonly',
            'https://www.googleapis.com/auth/spreadsheets'
        ]

        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'client_secret.json', SCOPES)
                creds = flow.run_local_server(port=0)
            
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        try:
            return build('sheets', 'v4', credentials=creds)
        except HttpError as error:
            return None


    def update(
        self, 
        values: list[list], 
        _range: Range, 
        secure_mode: bool = False
    ) -> int:
        if secure_mode:
            if len(self.select(_range)) > 0:
                return 0

        body = {
            "values": values
        }

        result = self.service.spreadsheets().values().update( # type: ignore
            spreadsheetId=self.sheet_id, range=_range.get_range(),
            valueInputOption="USER_ENTERED", body=body).execute()

        return int(result.get('updatedCells'))


    def sheetpage_name_by_range(self, range: str):
        return range.split("!", 1)[0]


    def sheetpage_id_by_name(self, name_range: str):
        try:
            obj = self.service.spreadsheets().get( # type: ignore
                spreadsheetId=self.sheet_id, 
                ranges=self.sheetpage_name_by_range(name_range), 
                fields='sheets(data(rowData(values(userEnteredFormat))),properties(sheetId))'
            ).execute()

            return obj['sheets'][0]['properties']['sheetId']
        except:
            return None


    def _execute_batch_update(
        self, body: dict, *, cancel_multi_execute: bool=False
    ):
        if self.__multi_execute_mode and not cancel_multi_execute:
            self.__multi_batch_updates.append(body)
            return self

        try:
            self.service.spreadsheets().batchUpdate(  # type: ignore
                spreadsheetId=self.sheet_id, body=body
            ).execute() 
            return True
        except Exception as e:
            print(e)
            return False

    
    def multi_execute(self, *args):
        try:
            requests = {"requests": []}
            for batch in self.__multi_batch_updates:
                requests["requests"] = [
                    *requests["requests"], *batch["requests"]]
            
            
        finally:
            self.__reset_multi_execute()


    def __reset_multi_execute(self):
        self.__multi_execute_mode: bool = False
        self.__multi_batch_updates: list[dict] = []
        self.__multi_batch_values_updates: list[dict] = []


    def format_text(
        self,
        _range: Range,
        format: TextFormat
    ) -> bool | GoogleSheets:

        _range_api_format = _range.get_google_sheet_api_format(self)

        body = {
            "requests": [
                {
                    "repeatCell": {
                        "range": _range_api_format,
                        "cell": {
                            "userEnteredFormat": {
                                "textFormat": format.dict(
                                    by_alias=True, exclude_unset=True)
                            }
                        },
                        "fields": "userEnteredFormat.textFormat"
                    }
                }
            ],
            # "data": [
            #     {
            #         "range": _range.get_range(),
            #         "values": [["oLÁ, TUDO BEM?"], ["Novo Teste"]]
            #     }
            # ],
            # 'valueInputOption': "USER_ENTERED"
        }

        return self._execute_batch_update(body)


    def add_border(
        self, 
        _range: Range,   # type: ignore
        each_cell: bool = True,
        style: BorderStyle | None = BorderStyle.SOLID,
        color: Color | ThemeColorType | None = Color(0, 0, 0),
        direction: BorderDirection = BorderDirection()
    ) -> bool | GoogleSheets:

        _borders = dict()
        if style is not None:
            _border_style = {
                "style": style.value
            }

            if type(color) is ThemeColorType:
                _border_style['colorStyle'] = { "themeColor": color.value }
            elif isinstance(color, Color):
                _border_style['colorStyle'] = { "rgbColor": color.dict() }

            if direction.top: _borders['top'] = _border_style
            if direction.bottom: _borders['bottom'] = _border_style
            if direction.left: _borders['left'] = _border_style
            if direction.right: _borders['right'] = _border_style

        _range_api_format = _range.get_google_sheet_api_format(self)

        body = {
            "requests": [
                {
                    "updateBorders": {
                        "range": _range_api_format,
                        **_borders
                    }
                }
            ]
        }

        if each_cell:
            body = {
                "requests": [
                    {
                        "repeatCell": {
                            "range": _range_api_format,
                            "cell": {
                                "userEnteredFormat": {
                                    "borders": {**_borders}
                                }
                            },
                            "fields": "userEnteredFormat.borders"
                        }
                    }
                ]
            }

        return self._execute_batch_update(body)


    def clear_border(
        self, 
        _range: Range,   # type: ignore
        each_cell: bool = True,
        direction: BorderDirection = BorderDirection()
    ) -> bool | GoogleSheets:
        return self.add_border(_range, each_cell, None, None, direction)


    def clear_values(self, _range: Range):
        request = self.service.spreadsheets().values().clear( # type: ignore
            spreadsheetId=self.sheet_id, range=_range.get_range(), body={})
        
        return request.execute()
    

    def select(self, _range: Range):
        result = self.service.spreadsheets().values().get( # type: ignore
            spreadsheetId=self.sheet_id, range=_range.get_range()).execute()
        rows = result.get('values', [])
        
        return rows

    
    def is_valid_sheet(self) -> bool:
        try:
            self.service.spreadsheets().values().get( # type: ignore
                spreadsheetId=self.sheet_id, range="A1").execute()
            return True
        except:
            return False


    @staticmethod
    def login():
        GoogleSheets("")


if __name__ == '__main__':
    rg = Range("Página1!B2:C")
    gs = GoogleSheets('')
    
    data = [
        ["NOME", "IDADE"],
        ["Teste", "134"],
        ["Vitor", "27"],
        ["Novo", "Teste"],
        ["Mais", "Um Teste"]
    ]
    
    # Inserindo dados
    gs.update(data, rg)

    # Adicionando Borda
    gs.add_border(
        rg.add_end_row(len(data) + 1)
    )
    
    # Deixando Cabeçalho em Negrito
    gs.format_text(
        rg.copy().set_end_row(rg.get_start_row()), 
        TextFormat(bold=True)
    )
    
    # Mudando fonte do conteúdo para Comic Sans
    gs.format_text(
        rg.copy().add_start_row(1), 
        TextFormat(
            font_family=TextFontFamily.COMIC_SANS_MS,
            foreground_color_style=ColorRGBA(255, 0, 0).to_color_style()
        )
    )

    gs.update([["3"], ["4"], ["5"]], Range(rg.get_sheetname(), "C4", None))

    gs.format_text(
        Range.last_range
        .sub_start_row(1)
        .set_end_col(Range.last_range.get_start_col())
        .set_end_row(Range.last_range.get_start_row() + 4),
        TextFormat(underline=True)
    )
