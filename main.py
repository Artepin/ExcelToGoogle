from pprint import pprint
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
from exellib import *
from spreadsheetgoogle import *


# одключил класс чела из той статьи, перенесу в отдельную библиотеку чуть позже
def htmlColorToJSON(htmlColor):
    if htmlColor.startswith("#"):
        htmlColor = htmlColor[1:]
    if htmlColor == "000000":
        return {"red": 1, "green": 1, "blue": 1}
    return {"red": int(htmlColor[0:2], 16) / 255.0, "green": int(htmlColor[2:4], 16) / 255.0, "blue": int(htmlColor[4:6], 16) / 255.0}


path = ('test.xlsx')
sheetid = 0 # id листа

# даем библиотеке знасть с каким файлом работать
el.redFile(path)
# указываем рабочий лист
el.sheetID(sheetid)
# получаем общее количиство строк и столбцев
rows = el.getRows()
columns = el.getColumns()
""" тестовая часть
print(rows,"*",columns)
print(el.bgColorRed('A1'))
print(el.bgColorGreen('A1'))
print(el.bgColorBlue('A1'))
"""
CREDENTIALS_FILE = 'fifth-sunup-319308-14f4f2f32c5a.json'  # Имя файла с закрытым ключом, вы должны подставить свое

"""
# Читаем ключи из файла
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])

httpAuth = credentials.authorize(httplib2.Http()) # Авторизуемся в системе
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) # Выбираем работу с таблицами и 4 версию API

spreadsheet = service.spreadsheets().create(body = {
    'properties': {'title': 'Первый тестовый документ', 'locale': 'ru_RU'},
    'sheets': [{'properties': {'sheetType': 'GRID',
                               'sheetId': 0,
                               'title': 'Лист номер один',
                               'gridProperties': {'rowCount': rows, 'columnCount': columns}}}]
}).execute()

spreadsheetId = spreadsheet['spreadsheetId'] # сохраняем идентификатор файла

driveService = apiclient.discovery.build('drive', 'v3', http = httpAuth) # Выбираем работу с Google Drive и 3 версию API
access = driveService.permissions().create(
    fileId = spreadsheetId,
    body = {'type': 'user', 'role': 'writer', 'emailAddress': 'dennerblack02@gmail.com'},  # Открываем доступ на редактирование
    fields = 'id'
).execute()"""

# первичная настройка
ss = Spreadsheet(CREDENTIALS_FILE, debugMode=True)
ss.create("Первый тестовый документ", "Лист номер один")
ss.shareWithEmailForWriting("dennerblack02@gmail.com")
# подготовка значений для отправки(формирование таблицы)

for column in range(1,columns+1):
    column_letter = el.columnLetter(column)
    for row in range(1,rows+1):
        cord = column_letter + str(row)  # return 'A1' (A1 к примеру)
        cords = (column_letter + str(row)+":"+column_letter + str(row)) # return 'A1:A1'
        print(cords)
        print(htmlColorToJSON(el.bgColor(cord)))
        print(el.getNumber(cord))
        if el.getNumber(cord) != 'None':
            ss.prepare_setValues(cords, [[el.getNumber(cord)]])
        ss.prepare_setCellsFormat(cords, {"backgroundColor": htmlColorToJSON(el.bgColor(cord))}, fields="userEnteredFormat.backgroundColor")



# тут запись подготовленных значений в google
pprint(ss.requests)
ss.runPrepared()
