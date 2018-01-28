import sys;
from PFE_Data_Cleaner.Data_Cleaner.data_Cleaner_Module import data_Cleaner as DC

path = '/Users/maxencepelloux/Documents/PFE/PFE_Data/Dirty_Data/MOEBIUS_Mondobrain_EVA.xlsx'
newPath = '/Users/maxencepelloux/Documents/PFE/PFE_Data/Clean_Data/SampleCleanID.XLSX'
sheetN = 0
formatIn = '%Y%m%d'
colIndex = 1
cleaner = DC.Cleaner(path, sheetN, colIndex, formatIn, newPath)
cleaner.openWB()
cleaner.purify()
cleaner.changeDate()
cleaner.anonymize()
cleaner.saveWB()
