import openpyxl;
from datetime import datetime

class Cleaner:

    def __init__(self, path):
        self.sheetN = 0
        self.banned = ['.',' ','#N/A','#DIV/0','inconnu','?']
        self.path = path

    def openWB(self, key, path):
        mods = {'basic': 1, 'aggreg': 2, 'joint': 3}
        if key == 1:
            self.wb = openpyxl.load_workbook(self.path)
        if key == 2:
            self.wb1 = openpyxl.load_workbook(path)
        if key == 3:
            self.wb2 = openpyxl.load_workbook(path)

    def anonymize(self, colIndexAN):
        try:
            track = 1
            self.wba = openpyxl.Workbook()
            sheet = self.wb.worksheets[self.sheetN]
            seen = {}
            for i in range(2, sheet.max_row):
                self.wba.active.cell(row=i-1, column=2).value = sheet.cell(row=i, column=colIndexAN).value
                if sheet.cell(row=i, column=colIndexAN).value not in seen.values():
                    seen.update({i-track:sheet.cell(row=i, column=colIndexAN).value})
                    self.wba.active.cell(row=i-1, column=1).value = i-track
                    sheet.cell(row=i, column=colIndexAN).value = i-track
                else:
                    track = track+1
                    self.wba.active.cell(row=i-1, column=1).value = list(seen.keys())[list(seen.values()).index(sheet.cell(row=i, column=colIndexAN).value)]
                    sheet.cell(row=i, column=colIndexAN).value = list(seen.keys())[list(seen.values()).index(sheet.cell(row=i, column=colIndexAN).value)]
            print('Succesfully anonymized column', list(sheet.rows)[0][colIndexAN-1],'.')
        except:
            print('Anonymization failed at row', i)

    def purify(self):
        try :
            sheet = self.wb.worksheets[self.sheetN]
            track = 0;
            for i in range(1, sheet.max_column) :
                for j in range(1, sheet.max_row) :
                    for ban in self.banned :
                        if sheet.cell(row=j, column=i).value == ban:
                            sheet.cell(row=j, column=i).value = ''
                            track=track+1
                    if ("." in str(sheet.cell(row=j, column=i).value)) or ("," in str(sheet.cell(row=j, column=i).value)):
                        track=track+1
                        sheet.cell(row=j, column=i).value = str(sheet.cell(row=j, column=i).value).replace(',','.')
            print('Sheet purified, edited '+ str(track) +' cells.')
        except :
            print('Cleaning banned data failed.')

    def changeDate(self, formatIn):
        try :
            sheet = self.wb.worksheets[self.sheetN]
            for n, head in enumerate(list(sheet.rows)[0]):
                dateKey = ["date", "Date", "DATE", "DT"]
                if any(key in head.value for key in dateKey):
                    for k in range(2, sheet.max_row):
                        if str(sheet.cell(row=k, column=n+1).value) != '':
                            if ' ' in str(sheet.cell(row=k, column=n+1).value):
                                sheet.cell(row=k, column=n+1).value = str(sheet.cell(row=k, column=n+1).value).replace(' ','')
                            if '.' in str(sheet.cell(row=k, column=n+1).value):
                                sheet.cell(row=k, column=n+1).value = str(sheet.cell(row=k, column=n+1).value).replace('.','')
                            dateObject = datetime.strptime(str(sheet.cell(row=k, column=n+1).value),formatIn)
                            sheet.cell(row=k, column=n+1).value = dateObject.strftime('%d/%m/%Y')
            print('Dates reformatted.')
        except ValueError:
            print('Error at cell[',k,' ',n,'], invalid cell content: ',str(sheet.cell(row=k, column=n+1).value),"""please make sure the formatIn
            ... represents the actual input date format.""")

    #def categorize(self):
        #TODO: Catégoriser avec intervalle précisé par l'utilisateur.

    def aggreg(self, paths):
        for i, p in enumerate(paths):
            wbb = openpyxl.load_workbook(p)
            sheetB = wbb.active
            headersB = list(sheetB.rows)[0]
            sheet = self.wb.worksheets[self.sheetN]
            headers = list(sheet.rows)[0]
            patch_row = sheet.max_row
            for n, head in enumerate(headers):
                for k, headB in enumerate(headersB):
                    if headB.value == head.value:
                        for j in range(1, sheetB.max_row):
                            sheet.cell(row=patch_row+j, column=n+1).value = sheetB.cell(row=j+1, column=k+1).value


    #def doublons(self, columns):
        #TODO: Check si la combinaison de colonnes spécifiées en liste est la même d'une ligne à l'autre


    #def joint(self):
        #TODO: Faire correspondre Open et Client

    def param(self, listBans):
        self.banned = self.banned + listBans

    def lock(self):
        try:
            if self.wb.protection.sheet == True:
                self.wb.protection.sheet = False
            else:
                self.wb.protection.sheet = True
            print('Worbook succesfully locked.')
        except:
            print('Lock failed.')


    def saveWB(self, key, newPath):
        try:
            if key == 1:
                self.wb.save(newPath)
                print('Succesfully saved at: ', newPath)
            else:
                self.wb.save(newPath)
                index = newPath.find('.xlsx')
                pathWBA = newPath[:index] + '_ID' + newPath[index:]
                self.wb.save(newPath)
                self.wba.save(pathWBA)
                print('Succesfully saved both files at:', newPath, 'and', pathWBA)
        except:
            print("Incorrect path :", newPath)
