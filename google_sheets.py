from __future__ import annotations
from typing import Optional
import os.path
import re
from enum import Enum
from dataclasses import dataclass
from typing import overload

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build, Resource
from googleapiclient.errors import HttpError


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


@dataclass
class BorderDirection:
    top: bool = True
    bottom: bool = True
    left: bool = True
    right: bool = True



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


class Range:

    _sheetname: str | None
    _start_row: int | None
    _start_col: int | None
    _end_row: int | None
    _end_col: int | None

    @overload
    def __init__(self, range: str) -> None:
        self.__consistency_adjust()

    @overload
    def __init__(self, sheetname: str, start: str, end: str) -> None:
        self.__consistency_adjust()

    @overload
    def __init__(
        self, 
        sheetname: str | None, 
        start_row: int | None, 
        start_col: int | None,
        end_row: int | None,
        end_col: int | None
    ) -> None:
        self.__consistency_adjust()

    
    def __consistency_adjust(self) -> Range:
        if self._start_row is None or self._start_row < 0:
            self._start_row = 0
        
        if self._end_row is None or self._end_row < self._start_row:
            self._end_row = self._start_row
        
        if self._start_col is None or self._start_col < 0:
            self._start_col = 1
        
        if self._end_col is None or self._end_col < self._start_col:
            self._end_col = self._start_col

        return self


    def _calc_row_col(
        self, *, 
        attr: str,
        add: int | None = None,
        sub: int | None = None,
        change: int | None = None
    ) -> Range | int | None:
        val = getattr(self, attr)

        if val is None or (
            add is None 
            and sub is None
            and change is None
        ):
            return self
        
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
        return self.get_start_col()


class GoogleSheets:

    def __init__(self, sheet_id: str) -> None:
        self.sheet_id: str = sheet_id
        self.service: Resource = self._load_service()


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


    def _range_to_grid_range(
        self, 
        range: str, 
        sheetpage_id: Optional[int] = None
    ) -> dict:

        def _get_column_index(range: Optional[str]) -> Optional[int]:
            if range is None:
                return None

            find = re.findall("[A-Z]", range)
            
            if len(find) < 1:
                return None

            index = 0
            for letter in find:
                index = index * 26 + 1 + ord(letter) - ord('A')

            return (index)


        def _get_row_number(range: Optional[str]) -> Optional[int]:
            if range is None:
                return None
            
            find = re.findall("[0-9]+", range)
            
            if len(find) < 1:
                return None 
            
            return (int(find[0]))


        if "!" in range:
            range = range.split("!", 1)[1]
        
        _range_split = range.split(":", 1)
        start = _range_split[0]
        end = _range_split[1] if len(_range_split) > 1 else None
        
        start_row: Optional[int] = _get_row_number(start)
        start_column: Optional[int] = _get_column_index(start)
        end_row: Optional[int] = _get_row_number(end)
        end_column: Optional[int] = _get_column_index(end)

        if start_row is not None:
            start_row -= 1
        
        if start_column is not None:
            start_column -= 1
        
        result = {
            "sheetId": sheetpage_id
        }
        
        if start_row is not None:
            result["startRowIndex"] = start_row
        
        if start_column is not None:
            result["startColumnIndex"] = start_column

        if end_row is not None:
            result["endRowIndex"] = end_row
        
        if end_column is not None:
            result["endColumnIndex"] = end_column
        
        return result


    def update(self, values: list[list], range: str) -> int:
        body = {
            "values": values
        }

        result = self.service.spreadsheets().values().update( # type: ignore
            spreadsheetId=self.sheet_id, range=range,
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


    def add_border(
        self, 
        _range: str,   # type: ignore
        each_cell: bool = True,
        style: BorderStyle | None = BorderStyle.SOLID,
        color: Color | ThemeColorType | None = Color(0, 0, 0),
        direction: BorderDirection = BorderDirection()
    ) -> bool:
        def _execute_spreadsheets(body: dict):
            try:
                self.service.spreadsheets().batchUpdate(  # type: ignore
                    spreadsheetId=self.sheet_id, body=body
                ).execute() 
                return True
            except Exception as e:
                print(e)
                return False

        _range: dict = self._range_to_grid_range(
            _range, self.sheetpage_id_by_name(_range)
        )
        
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

        body = {
            "requests": [
                {
                    "updateBorders": {
                        "range": _range,
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
                            "range": _range,
                            "cell": {
                                "userEnteredFormat": {
                                    "borders": {**_borders}
                                }
                            },
                            "fields": "userEnteredFormat"
                        }
                    }
                ]
            }

        return _execute_spreadsheets(body)


    def clear_border(
        self, 
        _range: str,   # type: ignore
        each_cell: bool = True,
        direction: BorderDirection = BorderDirection()
    ) -> bool:
        return self.add_border(_range, each_cell, None, None, direction)


    def clear_values(self, range: str):
        request = self.service.spreadsheets().values().clear( # type: ignore
            spreadsheetId=self.sheet_id, range=range, body={})
        
        return request.execute()
    

    def select(self, range: str):
        result = self.service.spreadsheets().values().get( # type: ignore
            spreadsheetId=self.sheet_id, range=range).execute()
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
    RANGE = "Página1!B2:C"
    gs = GoogleSheets('ADICIONE O ID DE SUA PLANILHA')
    
    data = [
        ["NOME", "IDADE"],
        ["Teste", "134"],
        ["Vitor", "27"],
        ["Novo", "Teste"],
        ["Mais", "Um Teste"]
    ]

    print(gs.update(data, RANGE))
    print(gs.sheetpage_id_by_name(RANGE))
    print(gs.add_border(f"{RANGE}{len(data)+1}", each_cell=True))
    print(gs.select(RANGE))
