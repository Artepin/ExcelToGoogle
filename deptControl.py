# writed by Artepin
import gspread
import datetime
import re
import os

from pprint import pprint
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
from spreadsheetgoogle import *


def deptControl():
    CREDENTIALS_FILE = os.getenv('USERPROFILE')+'\\Documents\\auth.json'
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
    httpAuth = credentials.authorize(httplib2.Http()) # Авторизуемся в системе
    service = apiclient.discovery.build('sheets', 'v4', http = httpAuth) # Выбираем работу с таблицами и 4 версию API
    ss = Spreadsheet(CREDENTIALS_FILE, debugMode=False)

    gp = gspread.service_account(filename=CREDENTIALS_FILE)
    print("введите ссылку на документ")
    # https://docs.google.com/spreadsheets/d/1EPldFQGirZHS6XplnIk1RlYSFoiNsQi-lB6xqzGnCII/edit#gid=0
    link = "https://docs.google.com/spreadsheets/d/1EPldFQGirZHS6XplnIk1RlYSFoiNsQi-lB6xqzGnCII/edit#gid=0" #input()
    link_id = link[39:83]
    ss.setSpreadsheetById(link_id)

    spreadsheet = gp.open_by_url(link)

    rowidcompl = 1
    rowidred = 1
    rowidyellow = 1

    redData = []
    redDataRows = []
    yellowData = []
    yellowDataRows = []
    complData = []
    complDataRows = []
    workerName = []
    workerNameRed = []
    workerNameYellow = []

    tasks = 3
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

    column_1 = worksheet.col_values(2)
    column_2 = worksheet.col_values(3)
    worker_column = worksheet.col_values(4)
    column = worksheet.col_values(5)
    column_fact = worksheet.col_values(6) # для оптимизации вогнал столбец с датой окончания работы в локальную память
    column_6 = worksheet.col_values(7)
    try:
        gen_header = worksheet.find("Генеральный план-график").row
    except:
        try:
            gen_header = worksheet.find("Генеральный план-график ").row
        except:
            raise SystemExit(13)

    try:
        calendar_header = worksheet.find("Календарный ПЛАН-ГРАФИК").row
    except:
        try:
            calendar_header = worksheet.find("Календарный ПЛАН-ГРАФИК ").row
        except:
            raise SystemExit(13)

    try:
        oper_header = worksheet.find("Оперативные задачи").row
    except:
        try:
            oper_header = worksheet.find("Оперативные задачи ").row
        except:
            raise SystemExit(13)

    gen_div = 4
    calendar_div = 5
    oper_div = 4

    # из-за отсутствующих значений в столбце с датой окончания работы, его длина меньше, нужно приравнять
    delta = len(column) - len(column_fact)
    for i in range(delta):
        column_fact.append('None')

    delta = len(column) - len(column_1)
    for i in range(delta):
        column_1.append('None')

    delta = len(column) - len(column_2)
    for i in range(delta):
        column_2.append('None')

    delta = len(column) - len(worker_column)
    for i in range(delta):
        worker_column.append('None')

    delta = len(column) - len(column_6)
    for i in range(delta):
        column_6.append('None')

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
        match =re.search(r'\d\d.\d\d.\d{4}',data) or re.search(r'\d\d.\d\d.\d{2}',data)
        if match:
            #print("date is valid")
            return True
        else:
            #print("date is not valid")
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
        try:
            datePlan = dateTransform(date)
        except:
            return True
        razn = dateNow - datePlan
        day = razn.days
        if int(day) > 14:
            #print("Red color")
            return True
        else:
            #print("Yellow color")
            return False

    def prohod(dataColumn, header_skip, end):
        j=header_skip-1
        for row_number in range(header_skip,end+1):
            i = dataColumn[row_number]
            j = j + 1
            match = validDate(i)
            if match:
                #print("Match true")
                cellCoord = 'E'+str(j)
                cell = column_fact[j] # теперь данные берутся из локального списка
                #print(cell)
                if validDate(cell):
                    #print("Work done")
                    if workerName.count(worker_column[j]) == 0:
                        workerName.append(worker_column[j])
                    compil = []
                    compil.append(column_1[j])
                    compil.append(column_2[j])
                    compil.append(worker_column[j])
                    compil.append(column[j])
                    compil.append(column_fact[j])
                    compil.append(column_6[j])
                    complData.append(compil)
                    complDataRows.append(j+1)
                else:
                    #print(i)
                    if isItLate(i)or(i != '-'):
                        #changeOfColor(cellCoord,"red")
                        if workerNameRed.count(worker_column[j]) == 0:
                            workerNameRed.append(worker_column[j])
                        compil = []
                        compil.append(column_1[j])
                        compil.append(column_2[j])
                        compil.append(worker_column[j])
                        compil.append(column[j])
                        compil.append(column_fact[j])
                        compil.append(column_6[j])
                        redData.append(compil)
                        redDataRows.append(j+1)
                        #print("changed red color on "+ cellCoord)
                    else:
                        #changeOfColor(cellCoord, "yellow")
                        if workerNameYellow.count(worker_column[j]) == 0:
                            workerNameYellow.append(worker_column[j])
                        compil = []
                        compil.append(column_1[j])
                        compil.append(column_2[j])
                        compil.append(worker_column[j])
                        compil.append(column[j])
                        compil.append(column_fact[j])
                        compil.append(column_6[j])
                        yellowData.append(compil)
                        yellowDataRows.append(j+1)
                        #print("changed yellow color on "+ cellCoord)
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



    def complSheet(row_start, div, rowidcompl):
        ss.copyHeader(rowidcompl,link_id,worksheetCompl.id,row_start, div, worksheet.id)
        rowidcompl += div
        for i in range(len(workerName)):
            for j in range(len(complData)):
                if complData[j].count(workerName[i]) != 0:
                    ss.copyRange(complDataRows[j], rowidcompl, link_id, worksheetCompl.id, worksheet.id)
                    rowidcompl += 1
            rowidcompl += 1
        return rowidcompl + 1

    def redSheet(row_start, div, rowidred):
        ss.copyHeader(rowidred,link_id,worksheetRed.id,row_start, div, worksheet.id)
        rowidred += div
        for i in range(len(workerNameRed)):
            for j in range(len(redData)):
                print(rowidred)

                if redData[j].count(workerNameRed[i]) != 0:
                    ss.copyRange(redDataRows[j], rowidred, link_id, worksheetRed.id, worksheet.id)
                    rowidred += 1
        return rowidred + 1

    def yellowSheet(row_start, div,rowidyellow):
        ss.copyHeader(rowidyellow,link_id,worksheetYellow.id,row_start, div, worksheet.id)
        rowidyellow += div
        for i in range(len(workerNameYellow)):
            for j in range(len(yellowData)):
                if yellowData[j].count(workerNameYellow[i]) != 0:
                    ss.copyRange(yellowDataRows[j], rowidyellow, link_id, worksheetYellow.id, worksheet.id)
                    rowidyellow += 1
        return rowidyellow + 1


    print(gen_header)
    print(calendar_header)
    print(oper_header)
    prohod(column,gen_header+gen_div-1, calendar_header)

    print(redData)
    print(workerNameRed)
    print(redDataRows)
    rowidred = redSheet(gen_header, gen_div, rowidred)
    rowidyellow = yellowSheet(gen_header, gen_div,rowidyellow)
    rowidcompl = complSheet(gen_header, gen_div, rowidcompl)
    print(rowidred)

    redData = []
    redDataRows = []
    yellowData = []
    yellowDataRows = []
    complData = []
    complDataRows = []
    workerName = []
    workerNameRed = []
    workerNameYellow = []

    print(calendar_header+calendar_div)
    prohod(column, calendar_header + calendar_div, oper_header)
    print(redData)
    print(workerNameRed)
    rowidred = redSheet(calendar_header, calendar_div, rowidred)
    rowidyellow = yellowSheet(calendar_header, calendar_div, rowidyellow)
    rowidcompl = complSheet(calendar_header, calendar_div, rowidcompl)
    redData = []
    redDataRows = []
    yellowData = []
    yellowDataRows = []
    complData = []
    complDataRows = []
    workerName = []
    workerNameRed = []
    workerNameYellow = []

    prohod(column, oper_header + oper_div, len(column))
    print(redData)
    print(workerNameRed)
    redSheet(oper_header, oper_div, rowidred)
    yellowSheet(oper_header, oper_div, rowidyellow)
    complSheet(oper_header, oper_div, rowidcompl)
    ss.runPrepared()