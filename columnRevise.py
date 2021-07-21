from pprint import pprint
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
from exellib import *
from spreadsheetgoogle import *

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
CREDENTIALS_FILE = 'auth.json'  # Имя файла с закрытым ключом, вы должны подставить свое
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                               ['https://www.googleapis.com/auth/spreadsheets',
                                                                'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())  # Авторизуемся в системе
service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)  # Выбираем работу с таблицами и 4 версию API
ss = Spreadsheet(CREDENTIALS_FILE, debugMode=False)
def columnRevise():
    # !!ВНИМАНИЕ!! перед релизом удалить шаблоны!
    print("введите ссылку на документ")
    # https://docs.google.com/spreadsheets/d/1pRohAKGYrcuRjKoZqByzzRB-eTlx6wvOrGM-nnUR0No/edit#gid=0
    link = "https://docs.google.com/spreadsheets/d/1pRohAKGYrcuRjKoZqByzzRB-eTlx6wvOrGM-nnUR0No/edit#gid=0"  # input()
    link_id = link[39:83]
    ss.setSpreadsheetById(link_id)

    print('введите название столбца, который нужно перенести')
    # Текст (цэ название столбца!)
    col_name = "Текст"  # input()
    print(col_name)
    print('введите название столбца, в который необходимо произвести перенос')
    # Текст (цэ название столбца!)
    col_name_google = "Текст"  # input()
    print(col_name_google)
    data_excel = []
    data_google = []

    # тут ищем координаты нужных ячеек(столбцев) в excel
    for column in range(1, columns + 1):  # как поправишь границы, не забудь добавить +1
        column_letter = el.columnLetter(column)
        for row in range(1, rows + 1):
            cord = column_letter + str(row)  # return 'A1' (A1 к примеру)
            cords = (column_letter + str(row) + ":" + column_letter + str(row))  # return 'A1:A1'
            if str(el.getNumber(cord)) == col_name:
                data_excel.append(cord)

    # тут просто проверка, ну думаю это и так понятно
    if len(data_excel) > 1:
        print("столбцев с указанным названием несколько, пожалуйста укажите номер подходящего варианта")
        print(data_excel)
        num = 1  # int(input())
        if (num > 0) and (num < len(data_excel) + 1):
            data_excel = data_excel[num - 1]
        else:
            print('указан неверный вариант')
            raise SystemExit(10)
    elif len(data_excel) == 0:
        print('столбцев с указанным названием необноружено')
        raise SystemExit(11)

    all_sheet = ss.getData('A1:SSR1000', link_id)  # получаем значения с листа в гугле

    excel_data_raw = []

    # этот цикл дает нам знать в каких строках искать нужные ячейки
    for row in range(1, len(all_sheet)):
        for column in range(1, len(all_sheet[row])):
            column_letter = el.columnLetter(column)
            cord = column_letter + str(row)  # return 'A1' (A1 к примеру)
            cords = (column_letter + str(row) + ":" + column_letter + str(row))  # return 'A1:A1'
            if all_sheet[row][column] == col_name_google:
                excel_data_raw.append(row + 1)

    # эта часть ищет координаты нужных ячеек(столбцев) в гугле
    if len(excel_data_raw) > 1:
        for row in range(len(excel_data_raw)):
            bin_iter = 512
            iter = 2
            index = 1
            check = []
            temp = 0
            place = 0
            iter2 = 0
            while (len(data_google) != row + 1):
                half = bin_iter / iter
                if int(half) == 0:
                    half = 1
                temp = place
                while True:
                    if (int(temp + 1) == 0) or (int(half + temp) == 0):
                        temp = 1

                    rng = (el.columnLetter(int(temp + 1)) + str(excel_data_raw[row]) + ":" + el.columnLetter(
                        int(half + temp)) + str(excel_data_raw[row]))
                    try:
                        check = ss.getData(rng, link_id)
                    except:
                        break
                    if (str(el.columnLetter(int(temp + 1)) + str(excel_data_raw[row]))) == (
                    str(el.columnLetter(int(half + temp)) + str(excel_data_raw[row]))):
                        data_google.append(el.columnLetter(int(temp + 1)) + str(excel_data_raw[row]))
                        break

                    if check.count(col_name_google) > 0:
                        break
                    elif temp > half * 4:
                        break
                    elif check.count(col_name_google) == 0:
                        iter2 += 1
                        break
                    temp += half

                place = half * iter2
                iter *= 2

    # таже проверка, но уже для гугла
    if len(data_google) > 1:
        print("столбцев с указанным названием несколько, пожалуйста укажите номер подходящего варианта")
        print(data_google)
        num = 2  # int(input())
        if (num > 0) and (num < len(data_google) + 1):
            data_google = data_google[num - 1]
        else:
            print('указан неверный вариант')
            raise SystemExit(10)
    elif len(data_google) == 0:
        print('столбцев с указанным названием необноружено')
        raise SystemExit(11)

    data_google = str(data_google)
    data_excel = str(data_excel)

    excel_cell_vals = []
    google_cell_vals = []

    # берем данные из нужного диапазона в excel
    for row in range(int(data_excel[1]), rows):
        excel_cell_vals.append(el.getNumber(str(data_excel[0]) + str(row + 1)))

    # берем данные из нужного диапазона в гугле
    rng = (str(data_google[0]) + str(int(data_google[1]) + 1) + ":" + str(data_google[0]) + str(
        int(data_google[1]) + rows - int(data_excel[1])))
    google_cell_vals.append(ss.getData(rng, link_id))

    # проводим сравнение значений и, если значение из excel больше, то записываем его в гугл таблицу вместо прежнего
    # если на этом месте в гугле нет значения, то просто записываем в него значения из excel, если оно есть
    for cell_val in range(len(excel_cell_vals)):
        if excel_cell_vals[cell_val] != 'None':
            try:
                if len(google_cell_vals[0][cell_val]) == 0:
                    ss.prepare_setValues(str(data_google[0]) + str(int(data_google[1]) + cell_val + 1),
                                         [[excel_cell_vals[cell_val]]])
                    continue
                if int(google_cell_vals[0][cell_val][0]) < excel_cell_vals[cell_val]:
                    ss.prepare_setValues(str(data_google[0]) + str(int(data_google[1]) + cell_val + 1),
                                         [[excel_cell_vals[cell_val]]])
            except IndexError:
                ss.prepare_setValues(str(data_google[0]) + str(int(data_google[1]) + cell_val + 1),
                                     [[excel_cell_vals[cell_val]]])
        else:
            break

    # тут запись подготовленных значений в google
    # pprint(ss.requests)
    ss.runPrepared()