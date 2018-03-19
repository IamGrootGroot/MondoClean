from collections import Counter
import openpyxl
import PFE_MondoClean.MondoClean.data_Cleaner_Module.data_cleaner as dc
c = dc.Cleaner()

path = '/Users/maxencepelloux/Documents/PFE/PFE_Data/Dirty_Data/Book1.x'

a = c.openWB(1, path)
print(a)
