from xlrd import open_workbook, cellname
import datetime, time

fileIn = "c:/Users/e1ze4m0/Documents/Indivior Program/Indivior Hub/SP Status/SP PA Check In 09.28.xlsx"

book = open_workbook(fileIn)
sheet = book.sheet_by_index(0)

print(book.nsheets)
print(book.sheet_names())

sheetsIn = book.sheet_names()

for x in sheetsIn:
    print(x)

print(sheet.name)
print(sheet.nrows)
print(sheet.ncols)

rowCnt = sheet.nrows
print("rowCnt: ", rowCnt)

clm00 = sheet.col_values(0, 1)
clm01 = sheet.col_values(1, 1)
clm10 = sheet.col_values(10, 1)
clm12 = sheet.col_values(12, 1)

dateoffset = 693594

toList1 = list(clm00)
toList2 = list(clm01)
toList3 = list(clm10)
toList4 = list(clm12)

toListA = (list(clm01), ",", list(clm12))


z = []
for j, p in enumerate(toList1):
    print(j, datetime.date.fromordinal(dateoffset + int(p)))
    ciDt = datetime.date.fromordinal(dateoffset + int(p))



print("*********************printing tL1****************")
print(tL1)

#for y in toList:
#    a = datetime.date.fromordinal(dateoffset + int(y))
#    print(y)
