import sys;
from data_Cleaner_Module import data_Cleaner as DC

path = '/Users/dagepel/Documents/PFE/PFE_Data/Dirty_Data/MOEBIUS_Mondobrain_EVA.xlsx'
newPath = '/Users/dagepel/Documents/PFE/PFE_Data/Clean_Data/anonym12Moeb.xlsx'
#paths = ['/Users/maxencepelloux/Documents/PFE/PFE_Data/Clean_Data/mondo2.xlsx',
#        '/Users/maxencepelloux/Documents/PFE/PFE_Data/Clean_Data/mondo3.xlsx']

formatIn = '%Y%m%d'
colIndexCat = None
cleaner = DC.Cleaner(path)
cleaner.openWB(1, None)
#cleaner.aggreg(paths)
cleaner.purify()
cleaner.changeDate(formatIn)
cleaner.anonymize([1,2])
cleaner.saveWB(2, newPath)
