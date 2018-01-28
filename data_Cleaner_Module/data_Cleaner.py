import openpyxl;
from datetime import datetime

class Cleaner:

    def __init__(self, path, sheetN, colIndex, formatIn, newPath):
        self.path = path
        self.sheetN = sheetN
        self.colIndex = colIndex
        self.formatIn  = formatIn
        self.newPath = newPath

    def openWB(self):
        self.wb = openpyxl.load_workbook(self.path)

    def anonymize(self):
        self.wb.create_sheet('ID')
        sheet_Names = self.wb.get_sheet_names()
        sheet = self.wb.get_sheet_by_name(sheet_Names[self.sheetN])
        sheet_ID = self.wb.get_sheet_by_name('ID')
        for i in range(2, sheet.max_row):
            sheet_ID.cell(row=i, column=self.colIndex).value = sheet.cell(row=i, column=self.colIndex).value
            sheet.cell(row=i, column=self.colIndex).value = i-1
        print('Succesfully anonymized column', list(sheet.rows)[0][self.colIndex-1],'.')

    def purify(self):
        try :
            sheet_Names = self.wb.get_sheet_names()
            sheet = self.wb.get_sheet_by_name(sheet_Names[self.sheetN])
            track = 0;
            banned = ['.',' ','#N/A','#DIV/0','inconnu'];
            for i in range(1, sheet.max_column) :
                for j in range(1, sheet.max_row) :
                    for ban in banned :
                        if sheet.cell(row=j, column=i).value == ban:
                            sheet.cell(row=j, column=i).value = ''
                            track=track+1
            print('Sheet purified, edited '+ str(track) +' cells.')
        except :
            print('Cleaning banned data failed.')

    def changeDate(self):
        try :
            sheet_Names = self.wb.get_sheet_names()
            sheet = self.wb.get_sheet_by_name(sheet_Names[self.sheetN])
            for n, head in enumerate(list(sheet.rows)[0]):
                dateKey = ["date", "Date", "DATE", "DT"]
                if any(key in head.value for key in dateKey):
                    for k in range(2, sheet.max_row):
                        if str(sheet.cell(row=k, column=n+1).value) != '':
                            dateObject = datetime.strptime(str(sheet.cell(row=k, column=n+1).value),self.formatIn)
                            sheet.cell(row=k, column=n+1).value = dateObject.strftime('%d/%m/%Y')
            print('Dates reformatted.')
        except ValueError:
            print('Error at cell[',k,' ',n,'], invalid cell content: ',str(sheet.cell(row=k, column=n+1).value),"""please make sure the formatIn
            ... represents the actual input date format.""")

    def saveWB(self):
        try:
            self.wb.save(self.newPath)
            print('Succesfully saved at: ', self.newPath)
        except:
            print("Incorrect path :", self.newPath)
