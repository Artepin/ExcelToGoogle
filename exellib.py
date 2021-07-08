import openpyxl
#import pyexcel
from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

class Exlib:

    fileread = openpyxl.load_workbook
    sheetid = 0

    def redFile(self, filePath):
        Exlib.fileread = openpyxl.load_workbook(filePath)

    def sheetID(self,id):
        Exlib.sheetid = id

    def getRows(self):
        file = Exlib.fileread
        sheet_id = Exlib.sheetid
        sheet = file.worksheets[sheet_id]
        return(sheet.max_row)

    def getColumns(self):
        file = Exlib.fileread
        sheet_id = Exlib.sheetid
        sheet = file.worksheets[sheet_id]
        return(sheet.max_column)

    def bgColorRed(self, cell1):
        file = Exlib.fileread
        sheet_id = Exlib.sheetid
        sheet = file.worksheets[sheet_id]
        cell = sheet[cell1]
        val = cell.fill.start_color.index
        redvalue = (int(val[2:2+2], 16))
        return(redvalue)

    def bgColorGreen(self, cell1):
        file = Exlib.fileread
        sheet_id = Exlib.sheetid
        sheet = file.worksheets[sheet_id]
        cell = sheet[cell1]
        val = cell.fill.start_color.index
        greenvalue = (int(val[4:4+2], 16))
        return(greenvalue)

    def bgColorBlue(self, cell1):
        file = Exlib.fileread
        sheet_id = Exlib.sheetid
        sheet = file.worksheets[sheet_id]
        cell = sheet[cell1]
        val = cell.fill.start_color.index
        bluevalue = (int(val[6:6+2], 16))
        return(bluevalue)

    def bgColor(self, cell1):
        file = Exlib.fileread
        sheet_id = Exlib.sheetid
        sheet = file.worksheets[sheet_id]
        cell = sheet[cell1]
        val = cell.fill.start_color.index
        ret = ((val[2:8]).format(3))
        return(ret)


el = Exlib()