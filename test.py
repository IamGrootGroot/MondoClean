import sys;
from data_Cleaner_Module import data_Cleaner as DC

path = '/Users/maxencepelloux/Documents/PFE/PFE_Data/Clean_Data/mondo1.xlsx'
newPath = '/Users/maxencepelloux/Documents/PFE/PFE_Data/Clean_Data/SampleIDV2.xlsx'
paths = ['/Users/maxencepelloux/Documents/PFE/PFE_Data/Clean_Data/mondo2.xlsx',
        '/Users/maxencepelloux/Documents/PFE/PFE_Data/Clean_Data/mondo3.xlsx']

sheetN = 0
formatIn = '%Y%m%d'
colIndexAN = 1
colIndexCat = None
cleaner = DC.Cleaner(path)
cleaner.openWB(1, None)
cleaner.aggreg(paths)
cleaner.purify()
cleaner.changeDate(formatIn)
cleaner.anonymize(colIndexAN)
cleaner.saveWB(2, newPath)
