from collections import Counter
import openpyxl


path = '/Users/maxencepelloux/Documents/PFE/PFE_Data/Dirty_Data/Fichier source étude durée AT VF TIS.xlsx'
colIndex = [7, 8]
wb = openpyxl.load_workbook(path)
sheet = wb.worksheets[0]
sequences = []
colCount = sheet.max_column+1
sheet.cell(row=1, column=colCount).value = 'COUNT'
for k, row in enumerate(sheet.iter_rows()):
    sequence = ''
    if k>0:
        for n, cell in enumerate(row):
            if n+1 in colIndex:
                sequence += str(cell.value)
        sequences.append(sequence)
occurences = Counter(sequences)
for n, s in enumerate(sequences):
    if s in occurences.keys():
        sheet.cell(row=n+2, column=colCount).value = occurences.get(s)
