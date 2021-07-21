# writed by Artepin
import gspread
import datetime
import re

from pprint import pprint
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
from spreadsheetgoogle import *


def deptControl():
    CREDENTIALS_FILE = 'C:\\PycharmProjects\\ExcelToGoogle\\auth.json'
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
    httpAuth = credentials.authorize(httplib2.Http()) # Авторизуемся в системе
    service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) # Выбираем работу с таблицами и 4 версию API
    ss = Spreadsheet(CREDENTIALS_FILE, debugMode=False)

    gp = gspread.service_account(filename='./C:\\PycharmProjects\\ExcelToGoogle\\auth.json')
    print("введите ссылку на документ")
    # https://docs.google.com/spreadsheets/d/1EPldFQGirZHS6XplnIk1RlYSFoiNsQi-lB6xqzGnCII/edit#gid=0
    link = "https://docs.google.com/spreadsheets/d/1EPldFQGirZHS6XplnIk1RlYSFoiNsQi-lB6xqzGnCII/edit#gid=0" #input()
    link_id = link[39:83]
    ss.setSpreadsheetById(link_id)

    spreadsheet = gp.open_by_url(link)

    redData = []
    yellowData = []
    complData = []
    complDataRows = []
    workerName = []


    #worksheetCompl = ss.addSheet("Выполненные", 1000, 20)

    try:
        worksheetRed = spreadsheet.get_worksheet(1).id
    except AttributeError:
        worksheetRed = ss.addSheet("Просрочено", 1000, 20)

    try:
        worksheetYellow = spreadsheet.get_worksheet(2).id
    except AttributeError:
        worksheetYellow = ss.addSheet("Подходящие", 1000, 20)

    try:
        worksheetCompl = spreadsheet.get_worksheet(3).id
    except AttributeError:
        worksheetCompl = ss.addSheet("Выполненные", 1000, 20)

    worksheet = spreadsheet.get_worksheet(0)
    worksheetRed = spreadsheet.get_worksheet(1)
    worksheetYellow = spreadsheet.get_worksheet(2)
    worksheetCompl = spreadsheet.get_worksheet(3)

    worker_column = worksheet.col_values(3)

    column = worksheet.col_values(4)

    column_fact = worksheet.col_values(5) # для оптимизации вогнал столбец с датой окончания работы в локальную память
    # из-за отсутствующих значений в столбце с датой окончания работы, его длина меньше, нужно приравнять
    delta = len(column) - len(column_fact)
    for i in range(delta):
        column_fact.append('None')

    stringSheet = worksheet.row_values(11)

    def dateTransform(data):
        if data !='None':
            day,month,year = data.split('.')
            date = datetime.date(int(year),int(month),int(day))
            return date
        else:
            print('Please,input correct date')
    def dateRazn(data1,data2):
        days = data2 - data1
        return days

    def redOrYellow(data):
        razn = data.days
        if int(razn) > 14:
            print("Red color")
        else:
            print("Yellow color")

    def validDate(data):
        if data == None:
            data = '0'
        matchOtmen = re.search(r'Отменен|отменено|-', data)
        if matchOtmen:
            return True
        match =re.search(r'\d\d.\d\d.\d{4}',data)
        if match:
            print("date is valid")
            return True
        else:
            print("date is not valid")
            return False


    def changeOfColor(coord,color):
        if color == "red":
            ss.prepare_setCellsFormat(coord+":"+coord, {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 0.0,
                            "blue": 0.0
                        }
                    })

        elif color == "yellow":
            ss.prepare_setCellsFormat(coord+":"+coord, {
                "backgroundColor": {
                    "red": 1.0,
                    "green": 1.0,
                    "blue": 0.0
                }
            })
        else:
            print("color is invalid")

    def isItLate(date):
        dateNow = datetime.date.today()
        datePlan = dateTransform(date)
        razn = dateNow - datePlan
        day = razn.days
        if int(day) > 14:
            print("Red color")
            return True
        else:
            print("Yellow color")
            return False

    def prohod(dataColumn):
        j=0
        for i in dataColumn:
            j = j + 1
            match = validDate(i)
            if match:
                print("Match true")
                cellCoord = 'E'+str(j)
                cell = column_fact[j-1] # теперь данные берутся из локального списка
                print(cell)
                if validDate(cell):
                    print("Work done")
                    if workerName.count(worker_column[j]) == 0:
                        workerName.append(worker_column[j])
                    complData.append(worksheet.row_values(j))
                    complDataRows.append(j-1)
                else:
                    print("No date")
                    if isItLate(i):
                        changeOfColor(cellCoord,"red")
                        print("changed red color on "+ cellCoord)
                    else:
                        changeOfColor(cellCoord, "yellow")
                        print("changed yellow color on "+ cellCoord)
            else:
                print("Match False")

    def copyColumn(fromColumn, startCell):
        #val = worksheet.acell(startCell).value
        valX= worksheet.acell(startCell).row
        valY=worksheet.acell(startCell).col
        if worksheet.acell(startCell).value == None:
            for i in fromColumn:
                worksheet.update_cell(valX,valY,i)
                valX=valX+1
        else:
            print("Value of your cell is not empty")

    def complSheet():
        rowid = 1
        for i in range(len(workerName)):
            ss.copyHeader(rowid,link_id,worksheetCompl.id)
            rowid += 4
            for j in range(len(complData)):
                if complData[j].count(workerName[i]) != 0:
                    ss.copyRange(complDataRows[j], rowid, link_id, worksheetCompl.id)
                    rowid += 1
            rowid += 1

    prohod(column)
    complSheet()
    ss.runPrepared()