from collections import Counter
import openpyxl


path = '/Users/maxencepelloux/Documents/PFE/PFE_Data/Dirty_Data/TestTypes.xlsx'
wb = openpyxl.load_workbook(path)
sheet = wb.worksheets[0]
