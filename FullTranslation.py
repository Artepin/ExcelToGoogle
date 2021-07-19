from pprint import pprint
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
from exellib import *
from spreadsheetgoogle import *

import numpy as np

# одключил класс чела из той статьи, перенесу в отдельную библиотеку чуть позже
def htmlColorToJSON(htmlColor):
    if htmlColor == "000000":
        return {"red": 1, "green": 1, "blue": 1}
    return {"red": int(htmlColor[0:2], 16) / 255.0, "green": int(htmlColor[2:4], 16) / 255.0, "blue": int(htmlColor[4:6], 16) / 255.0}


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

""" тестовая часть
print(el.getBorder("A1",borders[2]))

print(el.getFontColor("D3"))
print(el.bgColor('AD188'))

print(el.getHeight(1))
print(el.getWidth(el.columnLetter(1)))
font = el.getFont('A1')
print(font)
fontsize = el.getFontSize('A1')
print(fontsize)
bold = el.getBold('A1')
print(bold)
ital = el.getItalic('A1')
print(ital)
st = el.getStrikethrough('A1')
print(st)
undr = el.getUnderline('A1')
print(undr)

print(rows,"*",columns)
print(el.bgColorRed('A1'))
print(el.bgColorGreen('A1'))
print(el.bgColorBlue('A1'))
"""

CREDENTIALS_FILE = 'auth.json'  # Имя файла с закрытым ключом, вы должны подставить свое

border_controller = np.zeros((columns+1, rows+1, 4))

# первичная настройка
ss = Spreadsheet(CREDENTIALS_FILE, debugMode=True)
#ss.create("Первый тестовый документ", "Лист номер один", rows+1, columns+1)
#ss.shareWithEmailForWriting("dennerblack02@gmail.com")
# лучше по id чтобы не создавать каждый раз новый документ
ss.setSpreadsheetById('1pRohAKGYrcuRjKoZqByzzRB-eTlx6wvOrGM-nnUR0No')

# список с объедененными зонами
mergedlist = el.getMerged()

# подготовка значений для отправки(формирование таблицы)
for column in range(1,columns+1): # как поправишь границы, не забудь добавить +1
    column_letter = el.columnLetter(column)
    for row in range(1,rows+1):
        cord = column_letter + str(row)  # return 'A1' (A1 к примеру)
        cords = (column_letter + str(row)+":"+column_letter + str(row)) # return 'A1:A1'

        color = {"red": 0, "green": 0, "blue": 0}
        bgcolor = {"red": 1, "green": 1, "blue": 1}

        if el.getFontColor(cord) != False: color = htmlColorToJSON(el.getFontColor(cord))
        else: color = {"red": 0, "green": 0, "blue": 0}

        if el.bgColor(cord) != False: bgcolor = htmlColorToJSON(el.bgColor(cord))
        else: bgcolor = {"red": 1, "green": 1, "blue": 1}

        # форма стиля ячейки
        bodyJSON = {"backgroundColor": bgcolor,
                    'textFormat': {'foregroundColor': color,
                                   'fontFamily': el.getFont(cord),
                                   'fontSize': el.getFontSize(cord),
                                   'bold': el.getBold(cord),
                                   'italic': el.getItalic(cord),
                                   'strikethrough': el.getStrikethrough(cord),
                                   'underline': el.getUnderline(cord)}

                    }

        # заполняем формат ячейки
        ss.prepare_setCellsFormat(cords,bodyJSON)

        # заполняем значения ячеек
        if el.getNumber(cord) != 'None':
            ss.prepare_setValues(cords, [[el.getNumber(cord)]])

        # заполняем границы
        for orient in range(len(borders)):
            check_unit = {
                0: border_controller[column - 1][row-2][2],
                1: border_controller[column][row - 1][3],
                2: border_controller[column - 1][row][0],
                3: border_controller[column - 2][row - 1][1],
            }
            if check_unit[orient] != 1:
                #print(check_unit[orient])
                # форма стиля граница
                border = {'updateBorders': {'range':
                                                {'sheetId': ss.sheetId,
                                                 'startRowIndex': row - 1,
                                                 'endRowIndex': row,
                                                 'startColumnIndex': column - 1,
                                                 'endColumnIndex': column},
                                            str(borders[orient]): el.getBorder(cord, borders[orient])}}
                ss.requests.append(border)
                if el.getBorder(cord, borders[orient])['style'] != 'NONE':
                    border_controller[column - 1][row - 1][orient] = 1

# пример формы заполнения стиля ячейки
format = [{'values':
      [{'userEnteredValue': {'stringValue': 'Ячейка C2'},
        'effectiveValue': {'stringValue': 'Ячейка C2'},
        'formattedValue': 'Ячейка C2',
        'userEnteredFormat': {'backgroundColor': {'red': 1, 'green': 0.6},
                              'horizontalAlignment': 'CENTER',
                              'textFormat': {'fontSize': 14,
                                             'bold': True,
                                             'italic': True}},
        'effectiveFormat': {'backgroundColor': {'red': 1, 'green': 0.6},

                            'padding': {'top': 2, 'right': 3, 'bottom': 2, 'left': 3},
                            'horizontalAlignment': 'CENTER',
                            'verticalAlignment': 'BOTTOM',
                            'wrapStrategy': 'OVERFLOW_CELL',

                            'textFormat': {'foregroundColor': {},
                                           'fontFamily': 'Arial',
                                           'fontSize': 14,
                                           'bold': True,
                                           'italic': True,
                                           'strikethrough': False,
                                           'underline': False},
                            'hyperlinkDisplayType': 'PLAIN_TEXT'}}]}]

for i in range(len(mergedlist)):
    ss.prepare_mergeCells(str(mergedlist[i]))

for col in range(1,columns+1):
    ss.prepare_setColumnWidth(col-1, int(el.getWidth(el.columnLetter(col))))

for rw in range(1,rows+1):
    ss.prepare_setRowHeight(rw-1, int(el.getHeight(rw)))

# тут запись подготовленных значений в google
#pprint(ss.requests)
ss.runPrepared()
