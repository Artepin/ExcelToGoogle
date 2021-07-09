import openpyxl
#import pyexcel
from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

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

    def getNumber(self, cell1):
        file = Exlib.fileread
        sheet_id = Exlib.sheetid
        sheet = file.worksheets[sheet_id]
        cell = sheet[cell1]
        number = cell.value
        return(number)

    def columnLetter(self, columnnum):
        file = Exlib.fileread
        sheet_id = Exlib.sheetid
        sheet = file.worksheets[sheet_id]
        column = get_column_letter(columnnum)
        return (column)

    ################<<STYLES>>##################

    def getFont(self, cell1):
        file = Exlib.fileread
        sheet_id = Exlib.sheetid
        sheet = file.worksheets[sheet_id]
        cell = sheet[cell1]
        font = cell.font.name
        return font

    def getFontSize(self, cell1):
        file = Exlib.fileread
        sheet_id = Exlib.sheetid
        sheet = file.worksheets[sheet_id]
        cell = sheet[cell1]
        fontsize = cell.font.size
        return int(fontsize)

    def getBold(self, cell1):
        file = Exlib.fileread
        sheet_id = Exlib.sheetid
        sheet = file.worksheets[sheet_id]
        cell = sheet[cell1]
        boldstatus = cell.font.bold
        return bool(boldstatus)

    def getItalic(self, cell1):
        file = Exlib.fileread
        sheet_id = Exlib.sheetid
        sheet = file.worksheets[sheet_id]
        cell = sheet[cell1]
        italicstatus = cell.font.italic
        return bool(italicstatus)

    def getStrikethrough(self, cell1):
        file = Exlib.fileread
        sheet_id = Exlib.sheetid
        sheet = file.worksheets[sheet_id]
        cell = sheet[cell1]
        ststatus = cell.font.strikethrough
        return bool(ststatus)

    def getUnderline(self, cell1):
        file = Exlib.fileread
        sheet_id = Exlib.sheetid
        sheet = file.worksheets[sheet_id]
        cell = sheet[cell1]
        undrlstatus = bool(cell.font.underline)
        if undrlstatus != False:
            return True

el = Exlib()