import sys;
from data_Cleaner_Module import data_Cleaner as DC

path = '/Users/maxencepelloux/Documents/PFE/PFE_Data/Dirty_Data/Sample1.xlsx'
newPath = '/Users/maxencepelloux/Documents/PFE/PFE_Data/Clean_Data/SampleCleanCommas.XLSX'
sheetN = 0
formatIn = '%Y%m%d'
colIndex = 1
cleaner = DC.Cleaner(path, sheetN, colIndex, formatIn, newPath)
cleaner.openWB()
cleaner.purify()
cleaner.changeDate()
cleaner.anonymize()
cleaner.saveWB()
