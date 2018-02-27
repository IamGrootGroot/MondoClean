from collections import Counter
import openpyxl


path = '/Users/maxencepelloux/Documents/PFE/PFE_Data/Dirty_Data/Book1.xlsx'
colIndex = [7, 8]
cOlAdd = 9
wb = openpyxl.load_workbook(path)
sheet = wb.worksheets[0]
sequences = {}
colCount = sheet.max_column+1
sheet.cell(row=1, column=colCount).value = 'SUM'
for k, row in enumerate(sheet.iter_rows()):
    sequence = ''
    if k>0:
        for n, cell in enumerate(row):
            if n+1 in colIndex:
                sequence += str(cell.value)
            if n+1==colAdd:
                val = int(cell.value)   #Has to be numerical
        if sequence in sequences:
            sequences[sequence]+=val
        else:
            sequences.update({sequence:val})
for k, row in enumerate(sheet.iter_rows()):
    sequence = ''
    for n, cell in enumerate(row):
        if n+1 in colIndex:
            sequence += str(cell.value)
    if sequence in sequences:
        sheet.cell(row=k+1, column=colCount).value =  sequences.get(sequence)
wb.save(path.replace('Book1','BookTestSUM'))
