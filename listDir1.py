import os
lstDir = "C:/Users/e1ze4m0/Documents/Indivior Program/Indivior Hub/MS Excel/SP Status/Input"

#fl_list.clear()
fl_list = os.listdir(lstDir)

for x in fl_list:
     print(x, sep="\n")
