# -*- coding: UTF-8 -*-
import openpyxl;
from datetime import datetime

class Cleaner:

    def __init__(self, path):
        """ Cleaner object initialization """
        self.sheetN = 0
        self.banned = ['.',' ','#N/A','#DIV/0','inconnu','?']
        self.path = path

    def openWB(self, key, path):
        """ Load and open cleaner workbook:
        1: Open main workbook @self.path
        2: Open workbook to be aggregated to main workbook
        3: Open open data workbook to be jointed to main workbook
        """
        mods = {'basic': 1, 'aggreg': 2, 'joint': 3}
        if key == 1:
            self.wb = openpyxl.load_workbook(self.path)
        if key == 2:
            self.wb1 = openpyxl.load_workbook(path)
        if key == 3:
            self.wb2 = openpyxl.load_workbook(path)

    def anonymize(self, colIndexAN):
        """ Anonymization function: Sets cell values of given columns to their respective row index value
        if a duplicate is found, the set value is the same as the one encountered before. A second file is
        created for tracability purpose. This xlsx file contains the anonimyzation table with given values
        in column 1 and original values in column 2.

        1(colIndexAN is 1 integer): Anonymize column @colIndexAN
        2(colIndex is list of integers): Anonymize columns @colIndexAN, the values @column2 in the generated
        file become a concatenation of the values for each cell.
        """
        try:
            track = 1
            self.wba = openpyxl.Workbook()
            sheet = self.wb.worksheets[self.sheetN]
            seen = {}
            if len(colIndexAN)==1:
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
                colHead = {cell.value for n, cell in enumerate(list(sheet.rows)[0]) if n == colIndexminus}
                print('Succesfully anonymized column', colHead,'.')
            else:
                for i in range(2, sheet.max_row):
                    sequence = ''
                    for k in colIndexAN:
                        sequence += str(sheet.cell(row=i, column=k).value)
                        if sequence not in seen.values():
                            seen.update({i-track:sequence})
                            self.wba.active.cell(row=i, column=k).value = i-track
                            sheet.cell(row=i, column=k).value = i-track
                        else:
                            track = track+1
                            self.wba.active.cell(row=i-1, column=1).value = list(seen.keys())[list(seen.values()).index(sequence)]
                            sheet.cell(row=i, column=colIndexAN).value = list(seen.keys())[list(seen.values()).index(sequence)]
                    self.wba.active.cell(row=i-1, column=2).value = sequence
                colIndexminus = []
                for h in colIndexAN:
                        colIndexminus.append(h-1)
                colHeads = {cell.value for n, cell in enumerate(list(sheet.rows)[0]) if n in colIndexminus}
                print('Succesfully anonymized columns:', colHeads,'.')
        except:
            print('Anonymization failed at row', i,'.')

    def purify(self):
        """Purification function, removes banned characters given by the self.banned list,
        also removes ',' in integers"""
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
        """Reformats dates to the 'd/m/Y' format. The input format has to be specified by formatIn"""
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

    def aggreg(self, paths):
        """Aggregation function, appends different xlsx files @paths, no need to worry about empty or missing columns.
        WARNING: Column names have to be the same as the ones in the main workbook"""
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

    def param(self, listBans):
        """Add new banned characters"""
        self.banned = self.banned + listBans

    def lock(self):
        """Lock sheet"""
        try:
            if self.wb.protection.sheet == True:
                self.wb.protection.sheet = False
            else:
                self.wb.protection.sheet = True
            print('Worbook succesfully locked.')
        except:
            print('Locking failed.')

    def doublons(self, columns):
        """Identifies duplicates and highlights them"""
        try:
            uniq = {}
            dupes = []
            sheet = self.wb.worksheets[self.sheetN]
            for row in sheet.iter_rows():
                sequence = ''
                for cell in row:
                    if cell.col_idx in columns:
                        sequence += str(cell.value)
                if sequence not in uniq.values():
                    uniq.update({cell.row:sequence})
                else:
                    dupes.append(cell.row)
            for i in dupes:
                for j in range(1, sheet.max_column):
                    cell(row=i, column=j).fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        except:
            print("Can't find duplicates.")

    #def categorize(self):
        #TODO: Catégoriser avec intervalle précisé par l'utilisateur.

    #def joint(self):
        #TODO: Faire correspondre Open et Client

    def saveWB(self, key, newPath):
        """Save workbook:
        1: Save main workbook @newPath
        2: Save both main workbook and anonymization @newPath and newPath+'_ID'
        """
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
                print('Succesfully saved both files at:', newPath, 'and', pathWBA,'.')
        except:
            print("Incorrect path :", newPath,'.')
