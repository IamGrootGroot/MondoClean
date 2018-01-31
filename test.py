import sys;
from PFE_MondoClean.MondoClean.data_Cleaner_Module import data_Cleaner as DC

path = '/Users/maxencepelloux/Documents/PFE/PFE_Data/Dirty_Data/MOEBIUS_Mondobrain_EVA.xlsx'
newPath = '/Users/maxencepelloux/Documents/PFE/PFE_Data/Clean_Data/SampleIDV2.xlsx'
sheetN = 0
formatIn = '%Y%m%d'
colIndexAN = 1
colIndexCat = None
cleaner = DC.Cleaner(path)
cleaner.openWB(1, None)
cleaner.purify()
cleaner.changeDate(formatIn)
cleaner.anonymize(colIndexAN)
cleaner.saveWB(2, newPath)
