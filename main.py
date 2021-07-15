from pprint import pprint
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
from exellib import *
from spreadsheetgoogle import *

borders = ["top", "right", "bottom", "left"]

path = ('test.xlsx')

sheetid = 0 # id листа

# даем библиотеке знасть с каким файлом работать
el.redFile(path)
# указываем рабочий лист
el.sheetID(sheetid)
# получаем общее количиство строк и столбцев
rows = el.getRows()
columns = el.getColumns()


# первичная настройка
CREDENTIALS_FILE = 'fifth-sunup-319308-14f4f2f32c5a.json'  # Имя файла с закрытым ключом, вы должны подставить свое
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http()) # Авторизуемся в системе
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) # Выбираем работу с таблицами и 4 версию API
ss = Spreadsheet(CREDENTIALS_FILE, debugMode=True)
#ss.create("Первый тестовый документ", "Лист номер один", rows+1, columns+1)
#ss.shareWithEmailForWriting("dennerblack02@gmail.com")
# лучше по id чтобы не создавать каждый раз новый документ
print("введите ссылку на документ")
# https://docs.google.com/spreadsheets/d/1pRohAKGYrcuRjKoZqByzzRB-eTlx6wvOrGM-nnUR0No/edit#gid=0
link = "https://docs.google.com/spreadsheets/d/1pRohAKGYrcuRjKoZqByzzRB-eTlx6wvOrGM-nnUR0No/edit#gid=0" #input()
link_id = link[39:83]
ss.setSpreadsheetById(link_id)

print('введите название столбца, который нужно перенести')
# Текст (цэ название столбца!)
col_name = "Текст"#input()
print(col_name)
data_excel = []
data_google = []

# подготовка значений для отправки(формирование таблицы)
for column in range(1,columns+1): # как поправишь границы, не забудь добавить +1
    column_letter = el.columnLetter(column)
    for row in range(1,rows+1):
        cord = column_letter + str(row)  # return 'A1' (A1 к примеру)
        cords = (column_letter + str(row)+":"+column_letter + str(row)) # return 'A1:A1'
        if str(el.getNumber(cord)) == col_name:
            data_excel.append(cord)

if len(data_excel) > 1:
    print("столбцев с указанным названием несколько, пожалуйста укажите номер подходящего варианта")
    print(data_excel)
    num = 1 #int(input())
    if (num > 0) and (num < len(data_excel)+1):
        data_excel = data_excel[num-1]
    else:
        print('указан неверный вариант')
        raise SystemExit(10)
elif len(data_excel) == 0:
    print('столбцев с указанным названием необноружено')
    raise SystemExit(11)
all_sheet = ss.getData('A1:SSR1000',link_id)
print(len(all_sheet))
excel_data_raw = []
for row in range(1,len(all_sheet)): # как поправишь границы, не забудь добавить +1
    for column in range(1,len(all_sheet[row])):
        column_letter = el.columnLetter(column)
        cord = column_letter + str(row)  # return 'A1' (A1 к примеру)
        cords = (column_letter + str(row)+":"+column_letter + str(row)) # return 'A1:A1'
        if all_sheet[row][column] == col_name:
            excel_data_raw.append(row+1)
        #ss.prepare_setValues(cords, [[el.getNumber(cord)]])
if len(excel_data_raw) > 1:
    for row in range(len(excel_data_raw)):
        bin_iter = 512
        iter = 2
        index = 1
        check = []
        temp = 0
        place = 0
        iter2 = 1
        while(len(data_google) != row+1):
            half = bin_iter/iter
            temp = place
            while True:
                rng = (el.columnLetter(int(temp+1)) + str(excel_data_raw[row])+":"+ el.columnLetter(int(half+temp+1)) + str(excel_data_raw[row]))
                temp += half
                try:
                    check = ss.getData(rng,link_id)
                except:
                    break
                if (str(el.columnLetter(int(temp+1)) + str(excel_data_raw[row]))) == (str(el.columnLetter(int(half+temp+1)) + str(excel_data_raw[row]))):
                    data_google.append(el.columnLetter(int(temp+1)) + str(excel_data_raw[row]))
                    print('ppppp')
                    break
                if check.count(col_name) > 0:
                    break
                elif temp > half*2:
                    break
                elif check.count(col_name) == 0:
                    iter2 += 1
                    break
            place = half * iter2
            iter *= 2
print(data_google)


print(data_excel)
# тут запись подготовленных значений в google
#pprint(ss.requests)
#ss.runPrepared()
