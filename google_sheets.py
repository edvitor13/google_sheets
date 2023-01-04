from __future__ import annotations
from typing import Any, Optional
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build, Resource
from googleapiclient.errors import HttpError

from .enums import (
    BorderStyle,
    MergeType,
    ThemeColorType
)

from .exceptions import InvalidSheetException, InvalidAuth

from .dataclass import (
    BorderDirection,
    CellFormat,
    BorderDirection,
    TextFormat
)

from .support import Color, Range



class GoogleSheets:

    client_secret_filename: str = "auth/client_secret.json"
    token_filename: str = "auth/token.json"

    def __init__(
        self, 
        sheet_id: str, 
        *, 
        client_secret_filename: Optional[str] = None,
        token_filename: Optional[str] = None
    ) -> None:
        self.__client_secret_filename: str = GoogleSheets.client_secret_filename
        self.__token_filename: str = GoogleSheets.token_filename

        self.sheet_id: str = sheet_id
        self.service: Resource = self.__load_service(
            client_secret_filename, token_filename)
        
        self.__multi_execute_mode: bool = False
        self.__multi_batch_updates: list[dict] = []
        self.__multi_batch_updates_values: list[dict] = []
        
        if not self.is_valid_sheet():
            raise InvalidSheetException(
                f"ID table '{self.sheet_id}' is invalid")


    def __enter__(self) -> GoogleSheets:
        self.__active_multi_execute()
        return self


    def __exit__(self, *args) -> None:
        self.__multi_execute()
        self.__reset_multi_execute()

    
    def __multi_execute(self):
        # Update Batch Values
        body = {"data": []}
        for batch in self.__multi_batch_updates_values:
            _batch_copy = batch.copy()
            _batch_copy.pop("data")
            body = {**body, **_batch_copy}

            body["data"] = [
                *body["data"], *batch["data"]]
        
        if len(body["data"]):
            self._execute_batch_update_values(
                body, cancel_multi_execute=True)

        # Update Batch
        body = {"requests": []}
        for batch in self.__multi_batch_updates:
            body["requests"] = [
                *body["requests"], *batch["requests"]]
        
        if len(body["requests"]):
            self._execute_batch_update(
                body, cancel_multi_execute=True)


    def __active_multi_execute(self) -> GoogleSheets:
        self.__reset_multi_execute()
        self.__multi_execute_mode: bool = True
        return self


    def __reset_multi_execute(self) -> GoogleSheets:
        self.__multi_execute_mode: bool = False
        self.__multi_batch_updates: list[dict] = []
        self.__multi_batch_updates_values: list[dict] = []
        return self
    

    def __load_service(
        self, 
        client_secret_filename: Optional[str], 
        token_filename: Optional[str]
    ) -> Resource:
        if client_secret_filename is not None:
            self.__client_secret_filename = client_secret_filename
        
        if token_filename is not None:
            self.__token_filename = token_filename

        creds = None

        SCOPES = [
            'https://www.googleapis.com/auth/spreadsheets.readonly',
            'https://www.googleapis.com/auth/spreadsheets'
        ]

        if not os.path.exists(self.__client_secret_filename):
            raise InvalidAuth(
                "Could not find client configuration file: "
                f"'{self.__client_secret_filename}'"
            )

        if os.path.exists(self.__token_filename):
            creds = Credentials.from_authorized_user_file(
                self.__token_filename, SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.__client_secret_filename, SCOPES)
                creds = flow.run_local_server(port=0)
            
            with open(self.__token_filename, 'w') as token:
                token.write(creds.to_json())

        try:
            return build('sheets', 'v4', credentials=creds)
        except HttpError as e:
            raise InvalidAuth(e)


    def _execute_batch_update(
        self, body: dict, *, cancel_multi_execute: bool=False
    ) -> bool | None:
        if self.__multi_execute_mode and not cancel_multi_execute:
            self.__multi_batch_updates.append(body)
            return None

        try:
            self.service.spreadsheets().batchUpdate(  # type: ignore
                spreadsheetId=self.sheet_id, body=body
            ).execute() 
            return True
        except Exception as e:
            print(e)
            return False


    def _execute_batch_update_values(
        self, body: dict, *, cancel_multi_execute: bool=False
    ) -> int | None:
        if self.__multi_execute_mode and not cancel_multi_execute:
            self.__multi_batch_updates_values.append(body)
            return None

        try:
            result = self.service.spreadsheets().values().batchUpdate(  # type: ignore
                spreadsheetId=self.sheet_id, body=body
            ).execute()
            return int(result['responses'][0]['updatedCells'])
        except Exception as e:
            print(e)
            return 0

    
    def _range_to_api_format(
        self, _range: Range
    ) -> dict[str, Any]:
        sheetpage_id = self.sheetpage_id_by_name(_range.get_sheetname())

        start_row = (_range.get_start_row() - 1)
        start_col = (_range.get_start_col() - 1)
        end_row = _range.get_end_row()
        end_col = _range.get_end_col()
        
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


    def update(
        self, 
        values: list[list], 
        _range: Range, 
        secure_mode: bool = False
    ) -> int | None:
        if secure_mode:
            if len(self.select(_range)) > 0:
                return 0

        body = {
            "data": [
                {
                    "range": _range.get_range(),
                    "values": values
                }
            ],
            "valueInputOption": "USER_ENTERED"
        }

        return self._execute_batch_update_values(body)


    def sheetpage_id_by_name(self, name: str):
        try:
            obj = self.service.spreadsheets().get( # type: ignore
                spreadsheetId=self.sheet_id, 
                ranges=name, 
                fields='sheets(data(rowData(values(userEnteredFormat))),properties(sheetId))'
            ).execute()

            return obj['sheets'][0]['properties']['sheetId']
        except:
            return None


    def format_text(
        self,
        _range: Range,
        format: TextFormat
    ) -> bool | None:

        _range_api_format = self._range_to_api_format(_range)

        format_dict = format.dict(by_alias=True, exclude_unset=True)
        fields = format.fields("userEnteredFormat.textFormat", format_dict)

        body = {
            "requests": [
                {
                    "repeatCell": {
                        "range": _range_api_format,
                        "cell": {
                            "userEnteredFormat": {
                                "textFormat": format_dict
                            }
                        },
                        "fields": fields
                    }
                }
            ]
        }

        return self._execute_batch_update(body)

    
    def clear_format_text(self, _range: Range) -> bool | None:
        return self.format_text(_range, TextFormat())

    
    def format_cell(
        self,
        _range: Range,
        format: CellFormat
    ) -> bool | None:

        _range_api_format = self._range_to_api_format(_range)

        format_dict = format.dict(by_alias=True, exclude_unset=True)

        if len(format_dict):
            fields = format.fields("userEnteredFormat", format_dict)
        else:
            fields = format.fields("userEnteredFormat")

        body = {
            "requests": [
                {
                    "repeatCell": {
                        "range": _range_api_format,
                        "cell": {
                            "userEnteredFormat": format_dict
                        },
                        "fields": fields
                    }
                }
            ]
        }

        return self._execute_batch_update(body)


    def clear_format_cell(self, _range: Range):
        return self.format_cell(_range, CellFormat())


    def merge_cell(
        self,
        _range: Range,
        _type: MergeType | None = MergeType.MERGE_ALL
    ) -> bool | None:
        _range_api_format = self._range_to_api_format(_range)

        _merge_request = "unmergeCells" if _type is None else "mergeCells"

        body = {
            "requests": [
                {
                    _merge_request: {
                        "range": _range_api_format
                    }
                }
            ]
        }

        if _merge_request == "mergeCells":
            body['requests'][0][_merge_request]['mergeType'] = _type.value

        return self._execute_batch_update(body)


    def clear_merge_cell(self, _range: Range) -> bool | None:
        return self.merge_cell(_range, None)


    def add_border(
        self, 
        _range: Range,   # type: ignore
        each_cell: bool = True,
        style: BorderStyle | None = BorderStyle.SOLID,
        color: Color | ThemeColorType | None = Color(0, 0, 0),
        direction: BorderDirection = BorderDirection()
    ) -> bool | None:

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

        _range_api_format = self._range_to_api_format(_range)

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
