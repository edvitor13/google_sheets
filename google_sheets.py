from __future__ import print_function
from copy import deepcopy
from typing import Optional
import os.path
import re
from enum import Enum

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


    def update(self, values: list[list], range: str) -> int:
        body = {
            "values": values
        }

        result = self.service.spreadsheets().values().update( # type: ignore
            spreadsheetId=self.sheet_id, range=range,
            valueInputOption="USER_ENTERED", body=body).execute()

        return int(result.get('updatedCells'))


    def _range_to_grid_range(self, range: str, sheetpage_id: int=None) -> dict:

        def _get_column_index(range: str) -> Optional[int]:
            find = re.findall("[A-Z]", range)
            
            if len(find) < 1:
                return None

            index = 0
            for letter in find:
                index = index * 26 + 1 + ord(letter) - ord('A')

            return (index)


        def _get_row_number(range: str) -> Optional[int]:
            find = re.findall("[1-9]+", range)
            
            if len(find) < 1:
                return None 
            
            return (int(find[0]))


        if "!" in range:
            range = range.split("!", 1)[1]
        
        _range_split = range.split(":", 1)
        start = _range_split[0]
        end = _range_split[1] if len(_range_split) > 1 else None
        
        start_row: Optional[int] = (_get_row_number(start) - 1)
        start_column: Optional[int] = (_get_column_index(start) - 1)
        end_row: Optional[int] = _get_row_number(end)
        end_column: Optional[int] = _get_column_index(end)
        
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


    def sheetpage_name_by_range(self, range: str):
        return range.split("!", 1)[0]


    def sheetpage_id_by_name(self, name_range: str):
        try:
            obj = self.service.spreadsheets().get(
                spreadsheetId=self.sheet_id, 
                ranges=self.sheetpage_name_by_range(name_range), 
                fields='sheets(data(rowData(values(userEnteredFormat))),properties(sheetId))'
            ).execute()

            return obj['sheets'][0]['properties']['sheetId']
        except:
            return None


    def add_border(
        self, 
        _range: str, 
        each_cell: bool = True,
        style: BorderStyle = BorderStyle.SOLID
    ) -> bool:
        IS_FAIL = False

        def _execute_spreadsheets(body: dict):
            nonlocal IS_FAIL
            try:
                self.service.spreadsheets().batchUpdate(
                    spreadsheetId=self.sheet_id, body=body
                ).execute() 
            except Exception as e:
                print(e)
                IS_FAIL = True

        _range: dict = self._range_to_grid_range(
            _range, self.sheetpage_id_by_name(_range)
        )

        body = {
            "requests": [
                {
                    "updateBorders": {
                        "range": deepcopy(_range),
                        "top": {
                            "style": style.value
                        },
                        "bottom": {
                            "style": style.value
                        },
                        "left": {
                            "style": style.value
                        },
                        "right": {
                            "style": style.value
                        }
                    }
                }
            ]
        }

        if (
            each_cell 
            and _range['startRowIndex'] is not None
            and _range['endRowIndex'] is not None
            and _range['startColumnIndex'] is not None
            and _range['endColumnIndex'] is not None
        ):
            for row in reversed(range(_range['startRowIndex'], _range['endRowIndex'])):
                br = body['requests'][0]['updateBorders']['range']
                br['startRowIndex'] = row
                br['endRowIndex'] = row + 1

                for col in range(_range['startColumnIndex'], _range['endColumnIndex']):
                    br['startColumnIndex'] = col
                    br['endColumnIndex'] = col + 1

                    _execute_spreadsheets(body)

            return not IS_FAIL

        _execute_spreadsheets(body)

        return not IS_FAIL


    def select(self, range: str):
        result = self.service.spreadsheets().values().get( # type: ignore
            spreadsheetId=self.sheet_id, range=range).execute()
        rows = result.get('values', [])
        
        return rows

    
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
    print(gs.add_border(f"{RANGE}{len(data)+1}"))
    print(gs.select(RANGE))
