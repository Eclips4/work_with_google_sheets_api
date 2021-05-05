import httplib2
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SERVICE_URL = 'https://sheets.googleapis.com/$discovery/rest?version=v4'

CONFIG = {"KEY_FILE_LOCATION": 'Здесь указываешь путь до файла он разрешения либо .json либо .p12',
          "SERVICE_ACCOUNT_EMAIL": 'Здесь указываешь сервис-эмэйл, он необычного вида, достаточно длинный, не ошибешься'}



class SheetInit:
    """
    Инициализация build'a для работы с p12 файлом, в будущем сделаю и для Json
    """
    def __init__(self):
        self.sheet = build('sheets',
                           'v4',
                           http=ServiceAccountCredentials.from_p12_keyfile(CONFIG["SERVICE_ACCOUNT_EMAIL"],
                                                                           CONFIG["KEY_FILE_LOCATION"],
                                                                           scopes=SCOPES
                                                                           ).authorize(
                               httplib2.Http()),
                           discoveryServiceUrl=SERVICE_URL)

    def clear_range(self, sheet_id, list_name, range):
        """
        sheet_id - айди таблицы, когда переходишь по ссылке можно скопировать из юрла
        list_name - имя листа
        range - диапозон значений
        данная функция очищает значения
        example range: "Лист1!A2:B" - в данном случае очистит все значения начиная от A2 до B в Лист1
        """
        self.sheet.spreadsheets().values().clear(spreadsheetId=sheet_id,
                                                 range=f'{list_name}!{range}').execute()

    def insert_values(self, sheet_id, list_name, values, range):
        """
        sheet_id - айди таблицы, когда переходишь по ссылке можно скопировать из юрла
        list_name - имя листа
        range - диапозон значений
        values - значения
        данная функция вставляет значения
        example values, range: range = "Лист1!A2:B", values = [[1,2], [3,4]] - в данном случае
        ПЕРЕПИШЕТ значения в Лист1 по диапозону А2:B
        """
        self.sheet.spreadsheets().values().update(spreadsheetId=sheet_id,
                                                  range=f'{list_name}!{range}',
                                                  valueInputOption='USER_ENTERED',
                                                  body={'values': values}).execute()

    def append_values(self, sheet_id, list_name, values, range):
        """
        sheet_id - айди таблицы, когда переходишь по ссылке можно скопировать из юрла
        list_name - имя листа
        range - диапозон значений
        values - значения
        данная функция вставляет значения
        example values, range: range = "Лист1!A2:B", values = [[1,2], [3,4]] - в данном случае
        ДОПОЛНИТ, НО НЕ ПЕРЕПИШЕТ значения в Лист1 по диапозону А2:B
        """
        self.sheet.spreadsheets().values().append(spreadsheetId=sheet_id,
                                                  range=f'{list_name}!{range}',
                                                  valueInputOption='USER_ENTERED',
                                                  body={'values': values}).execute()


