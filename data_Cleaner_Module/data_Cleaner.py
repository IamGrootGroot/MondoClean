# -*- coding: UTF-8 -*-
import openpyxl
import os
import csv
import secrets
import string
import progressbar
from datetime import datetime

class Cleaner:

    def __init__(self):
        """ Cleaner object initialization """
        self.sheetN = 0
        self.banned = ['.',' ','#N/A','#DIV/0','inconnu','?','NA','None']
        self.wbList = []
        self.pathList = []
        self.taskBytes = 0
        self.maxBytes = 0

    def openWB(self, key, path):
        """ Load and open cleaner workbook:
        1: Open main workbook @self.path
        2: Open workbook to be aggregated to main workbook
        3: Open open data workbook to be jointed to main workbook
        """
        mods = {'basic': 1, 'aggreg': 2, 'joint': 3}
        if key == 1:
            self.wb = openpyxl.load_workbook(path)
            wbC = self.wb
            self.wbList.append(wbC)
            if path not in self.pathList:
                self.pathList.append(path)
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
            self.maxBytes = sheet.max_row
            self.taskBytes = 0
            seen = {}
            if len(colIndexAN) == 1:
                for i in range(2, sheet.max_row+1):
                    self.wba.active.cell(row=i-1, column=2).value = sheet.cell(row=i, column=colIndexAN[0]).value
                    if sheet.cell(row=i, column=colIndexAN[0]).value not in seen.values():
                        seen.update({i-track:sheet.cell(row=i, column=colIndexAN[0]).value})
                        self.wba.active.cell(row=i-1, column=1).value = i-track
                        sheet.cell(row=i, column=colIndexAN[0]).value = i-track
                    else:
                        track = track+1
                        self.wba.active.cell(row=i-1, column=1).value = list(seen.keys())[list(seen.values()).index(sheet.cell(row=i, column=colIndexAN).value)]
                        sheet.cell(row=i, column=colIndexAN).value = list(seen.keys())[list(seen.values()).index(sheet.cell(row=i, column=colIndexAN).value)]
                    self.taskBytes = i-1
                colHead = {cell.value for n, cell in enumerate(list(sheet.rows)[0]) if n+1 == colIndexAN}
                print('Succesfully anonymized column', colHead,'.')
            else:
                for i in range(2, sheet.max_row+1):
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
                self.taskBytes = i-1
                print('Succesfully anonymized columns:', colHeads,'.')
        except:
            print('Anonymization failed at row', i,'.')

    def purify(self):
        """Purification function, removes banned characters given by the self.banned list,
        also removes ',' in integers"""
        try :
            sheet = self.wb.worksheets[self.sheetN]
            self.maxBytes = sheet.max_column*sheet.max_row
            self.taskBytes = 0
            track = 0;
            for i in range(1, sheet.max_column+1) :
                for j in range(2, sheet.max_row+1) :
                    if sheet.cell(row=j, column=i).value in self.banned:
                            sheet.cell(row=j, column=i).value = None
                            track=track+1
                    if ("." in str(sheet.cell(row=j, column=i).value)) or ("," in str(sheet.cell(row=j, column=i).value)):
                        track=track+1
                        sheet.cell(row=j, column=i).value = str(sheet.cell(row=j, column=i).value).replace(',','.')
                    self.taskBytes=self.taskBytes+1
            print('Sheet purified, edited '+ str(track) +' cells.')
        except :
            print('Cleaning banned data failed.')

    def changeDate(self, formatIn):
        """Reformats dates to the 'd/m/Y' format. The input format has to be specified by formatIn"""
        try :
            sheet = self.wb.worksheets[self.sheetN]
            self.maxBytes = len(list(sheet.rows)[0])*sheet.max_row
            self.taskBytes = 0
            dateStyleTag = self.dateHexGen()
            dateStyle = openpyxl.styles.NamedStyle(name=dateStyleTag, number_format='DD/MM/YYYY')
            self.wb.add_named_style(dateStyle)
            for n, head in enumerate(list(sheet.rows)[0]):
                dateKey = ["date", "Date", "DATE", "DT "]
                if any(key in head.value for key in dateKey):
                    for k in range(2, sheet.max_row+1):
                        if type(sheet.cell(row=k, column=n+1).value) is not datetime:
                            if str(sheet.cell(row=k, column=n+1).value) != '' and 'None' not in str(sheet.cell(row=k, column=n+1).value):
                                if ' ' in str(sheet.cell(row=k, column=n+1).value):
                                    sheet.cell(row=k, column=n+1).value = str(sheet.cell(row=k, column=n+1).value).replace(' ','')
                                if '.' in str(sheet.cell(row=k, column=n+1).value):
                                    sheet.cell(row=k, column=n+1).value = str(sheet.cell(row=k, column=n+1).value).replace('.','')
                                dateObject = datetime.strptime(str(sheet.cell(row=k, column=n+1).value),formatIn)
                                sheet.cell(row=k, column=n+1).value = dateObject.strptime(dateObject.strftime('%d/%m/%Y'),'%d/%m/%Y')
                                sheet.cell(row=k, column=n+1).style = dateStyleTag
                        else:
                            sheet.cell(row=k, column=n+1).value = sheet.cell(row=k, column=n+1).value.strptime(sheet.cell(row=k, column=n+1).value.strftime('%d/%m/%Y'),'%d/%m/%Y')
                            sheet.cell(row=k, column=n+1).style = dateStyleTag
                self.taskBytes = self.taskBytes+1
            print('Dates reformatted.')
        except ValueError:
            print('Error at cell[',k,' ',n+1,'], invalid cell content: ', str(sheet.cell(row=k, column=n+1).value),"""please make sure the formatIn
            ... represents the actual input date format.""")

    def aggreg(self, paths):
        """Aggregation function, appends different xlsx files @paths, no need to worry about empty or missing columns.
        WARNING: Column names have to be the same as the ones in the main workbook"""
        sheet = self.wb.worksheets[self.sheetN]
        self.maxBytes = len(paths)
        self.taskBytes = 0
        headers = list(sheet.rows)[0]
        for i, p in enumerate(paths):
            wbb = openpyxl.load_workbook(p)
            sheetB = wbb.active
            headersB = list(sheetB.rows)[0]
            patch_row = sheet.max_row
            for n, head in enumerate(headers):
                for k, headB in enumerate(headersB):
                    if headB.value == head.value:
                        for j in range(2, sheetB.max_row+1):
                            sheet.cell(row=patch_row+j-1, column=n+1).value = sheetB.cell(row=j, column=k+1).value
            self.taskBytes = i+1

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
            self.maxBytes = len(columns)*sheet.max_row
            self.taskBytes = 0
            for row in sheet.iter_rows():
                sequence = ''
                for cell in row:
                    if cell.col_idx in columns:
                        sequence += str(cell.value)
                if sequence not in uniq.values():
                    uniq.update({cell.row:sequence})
                else:
                    dupes.append(cell.row)
                self.taskBytes = self.taskBytes+1
            for i in dupes:
                for j in range(1, sheet.max_column+1):
                    cell(row=i, column=j).fill = PatternFill(bgColor="FFC7CE", fill_type = "solid")
        except:
            print("Can't find duplicates.")

    def joint(self, path, colComp1, colComp2, colJoints):
        """Joint opendata @path. Finds matching values between colComp1 and colComp2
        and add the data in colJoint at matching index"""
        cleaner2.openWB(1,path)
        cleaner2.purify()
        cleaner2.changeDate(None)
        sheetB = cleaner2.wb.worksheets[cleaner2.sheetN]
        sheet = self.wb.worksheets[self.sheetN]
        self.maxBytes = sheet.max_row
        self.taskBytes = 0
        mr = sheet.max_row
        mb = sheetB.max_row
        mc = sheet.max_column
        colc1 = []
        colc2 = []
        colj = []
        for col1 in sheet.iter_cols(min_row=2, min_col=colComp1, max_col=colComp1, max_row=mr):
            for cell1 in col1:
                colc1.append(cell1.value)
        for col2 in sheetB.iter_cols(min_row=2, min_col=colComp2, max_col=colComp2, max_row=mb):
            for i, cell2 in enumerate(col2):
                joints = []
                colc2.append(cell2.value)
                for k in range(len(colJoints)):
                    for rowJ in sheetB.iter_rows(min_row=i+2, min_col=colJoints[k], max_col=colJoints[k], max_row=i+2):
                        for cell3 in rowJ:
                            joints.append(cell3.value)
                colj.append(joints)
        idx = {}
        for j in range(len(colc2)):
            self.maxBytes = sheet.max_row
            self.taskBytes = 0
            i = 0
            while colc1[i]!=colc2[j] and i<len(colc1)-1:
                i = i+1
            if i>=len(colc1):
                pass
            else:
                idx.update({i:colj[j]})
            self.taskBytes = j+1

        for i, j in enumerate(idx.keys()):
            for k in range(len(colJoints)):
                sheet.cell(row=j+2, column=k+1+mc).value = idx.get(j)[k]
        self.purify()

    def categorize(self, mod, colIndexC, changes):
        """Catégorisation, changes cell values @column
        colIndexC to the corresponding key in the changes dict"""
        if mod == "numerical":
            sheet = self.wb.worksheets[self.sheetN]
            self.maxBytes = sheet.max_row
            self.taskBytes = 0
            for col in sheet.iter_cols(min_row=2, min_col=colIndexC, max_col=colIndexC, max_row=sheet.max_row):
                for cell in col:
                    mask = None
                    mask = [mask for n, mask in enumerate(list(changes.keys())) if ((cell.value is not None) and (int(list(changes.values())[n][0]) <= int(cell.value) <= int(list(changes.values())[n][-1])))]
                    if mask:
                        cell.value = mask[0]
                    else:
                        pass
                self.taskBytes = self.taskBytes+1
        if mod == "substitute":
            sheet = self.wb.worksheets[self.sheetN]
            self.maxBytes = sheet.max_row
            self.taskBytes = 0
            for col in sheet.iter_cols(min_row=2, min_col=colIndexC, max_col=colIndexC, max_row=sheet.max_row):
                for cell in col:
                    mask = None
                    mask = [mask for n, mask in enumerate(list(changes.keys())) if list(changes.values())[n]==cell.value]
                    if mask:
                        cell.value = mask[0]
                    else:
                        pass
                self.taskBytes = self.taskBytes+1

    def timeMachine(self, request):
        """A time machine to allow undo and resets"""
        if request == 'pullBack':
            del self.wbList[-1]
            os.remove(self.pathList[-1])
            del self.pathList[-1]
            self.wb = self.wbList[-1]
            return self.pathList[-1]
        if request == 'fullReset':
            del self.wbList[1:]
            for p in self.pathList[1:]:
                os.remove(p)
            del self.pathList[1:]
            self.wb = self.wbList[0]
            return self.pathList[0]

    def dateHexGen(self):
        """Hex key generator for dateStyle"""
        try:
            return secrets.token_hex(32)
        except:
            return secrets.token_hex(32)

    def getProgress(self):
        return (self.taskBytes*100)/self.maxBytes

    def saveWB(self, key, newPath):
        """Save workbook:
        1: Save main workbook @newPath
        2: Save both main workbook and anonymization @newPath and newPath+'_ID'
        """
        try:
            if key == 1:
                self.wb.save(newPath)
                self.pathList.append(newPath)
                print('Succesfully saved at: ', newPath)
            else:
                self.wb.save(newPath)
                self.pathList.append(newPath)
                index = newPath.find('.xlsx')
                pathWBA = newPath[:index] + '_ID' + newPath[index:]
                self.wb.save(newPath)
                self.wba.save(pathWBA)
                print('Succesfully saved both files at:', newPath, 'and', pathWBA,'.')
        except:
            print("Incorrect path :", newPath,'.')
