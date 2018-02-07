import sys;
from data_Cleaner_Module import data_Cleaner as DC

path = '/Users/maxencepelloux/Documents/PFE/PFE_Data/Dirty_Data/Book1.xlsx'
paths = ['/Users/maxencepelloux/Documents/PFE/PFE_Data/Dirty_Data/Book2.xlsx']
pathJ = '/Users/maxencepelloux/Documents/PFE/PFE_Data/Dirty_Data/Gastro.xlsx'
newPath = '/Users/maxencepelloux/Documents/PFE/PFE_Data/Clean_Data/GROSTEST.xlsx'
colComp1 = 6
colComp2 = 1
colJoints = [2]
formatIn = '%Y%m%d'
colIndexAN = 1
colIndexCat = None
cleaner = DC.Cleaner(path)
cleaner.openWB(1, None)
cleaner.aggreg(paths)
cleaner.purify()
cleaner.changeDate(formatIn)
cleaner.joint(pathJ, colComp1, colComp2, colJoints)
cleaner.anonymize(colIndexAN)
cleaner.purify()
cleaner.saveWB(2, newPath)
