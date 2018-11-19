from xlrd import open_workbook
import datetime

fileIn = "c:/Users/e1ze4m0/Documents/Indivior Program/Indivior Hub/MS Excel/SP Status/Input/SP PA Check In 09.28.xlsx"

wb = open_workbook(fileIn, on_demand=True)

sheetCnt = wb.nsheets

print("Number of Worksheets: ", sheetCnt)

def cnvt_excel_dt(dt):

    dateoffset = 693594
    return datetime.date.fromordinal(dateoffset + int(dt))

for s in wb.sheets():
    print('Sheet Name: ', s.name)

sheet1 = wb.sheet_by_index(0)

rowCnt1 = sheet1.nrows

print("Value of Row Count 1 : ", rowCnt1)


for row in range(1, sheet1.nrows):
    rowVal = sheet1.row_values(row)
    dt = rowVal[0]
    ciDate = cnvt_excel_dt(dt)
    dt = rowVal[10]
    rtDate = cnvt_excel_dt(dt)
    dtDiff = datetime.timedelta(rowVal[0] - rowVal[10])
    print(int(rowVal[1]), ",", rtDate, ",", ciDate, ",", rowVal[11], ",", dtDiff)

wb.release_resources()





