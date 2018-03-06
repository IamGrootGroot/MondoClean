from collections import Counter
import openpyxl


path = '/Users/maxencepelloux/Documents/PFE/PFE_Data/Dirty_Data/TestTypes.xlsx'
wb = openpyxl.load_workbook(path)
sheet = wb.worksheets[0]
sheet.cell(row=10, column=1).value = float('1,1'.replace(',','.'))
sheet.cell(row=10, column=1).number_format = '0.00'
print(type(str(sheet.cell(row=10, column=1).value)))

float(sheet.cell(row=5, column=2).value) is float

l = [1, 0, 2, 3]
del l[l.index(0):]
l
