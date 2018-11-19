import os
from xlrd import open_workbook
import datetime
import math
from xlwt import Workbook
import collections
from collections import defaultdict
import pandas as pd


def bld_output(sp_pa_audict, cdatein):
    rw = 1
    cl = 0

    for ia, va in sp_pa_audict.items():
#        print("audict.items key: ", ia, " audict.items value: ", va)
        rowkey = ia.split('.')
#        print("rowkey: ", rowkey)
        prtcase = rowkey[0]
        prtrtdt = rowkey[1]
#        print("prcase: ", prtcase, " prcidt: ", prtrtdt)
        itmkeys = list(va.keys())
        itmvals = list(va.values())
#        print("rw:", rw)
        write_row_key(prtcase, prtrtdt, rw)
        lpx = 0
        while lpx < len(va):
            prtkey = itmkeys[lpx]
            prtvals = itmvals[lpx]
#            print("itmkey: ", prtkey, " itmvals: ", prtvals, " lpx:", lpx)
            lpx = lpx + 1
        write_row_value(itmvals, itmkeys, rw)
        ofst = lpx
        rw = rw + 1


def write_row_key(prtcase, prtrtdt, rw):

#    print("Print row key row value: ", rw)
    sheet1.write(rw, 0, prtcase)
    sheet1.write(rw, 1, prtrtdt)


def write_row_value(rowval, itmkeys, rw):
    idxe = 0
    cdatein = pd.datetime.now()
#    print("cdatein: ", cdatein)
    for val in rowval:
        edate = pd.to_datetime(itmkeys[idxe])
        edate2 = pd.datetime.now()
        print("edate: ", edate, " edate2: ", edate2)
        wcol = edate2 - edate
        wcol2 = wcol.days
        print("wcol2: ", wcol2)
        if wcol2 < 45:
            sheet1.col(46 - wcol2).width = 256 * 25
            sheet1.write(rw, 46 - wcol2, val)
            idxe = idxe + 1


def dt_tbl():
    cdate = datetime.date.today()
    pycurdt = datetime.date.toordinal(cdate)
    pycurdt2 = pycurdt - 45
    cdate2 = pycurdt2 - 693594

    print("Today's Date is: ", cdate)
    print("Today's Date as ordinal value: ", cdate2)
    print("************** Start List *****************")

    sheet1.col(0).width = 256 * 20
    sheet1.write(0, 0, "Case ID")
    sheet1.col(1).width = 256 * 20
    sheet1.write(0, 1, "Routing Date")

    for x in range(1, 45, 1):
        mdate = cdate2 + x
        dt = mdate
        mdate2 = cnvt_excel_dt(dt)
#        print("2nd Date table build - value of mdate: ", mdate, " display value: ", mdate2, " ndx: ", x)
        smdate2 = str(mdate2)
        smdateout = "D" + smdate2.replace("-", "_")
        sheet1.col(x + 1).width = 256 * 13
        if x == 0:
            print("in header x = 0")
            sheet1.write(0, x + 2, smdateout)
        sheet1.write(0, x + 1, smdateout)
    return cdate2


def makehash():

    return collections.defaultdict(makehash)


def cnvt_excel_dt(dt):

    dateoffset = 693594
    return datetime.date.fromordinal(dateoffset + int(dt))


def write_dict1(dictKey, dictValue, dictKey2, dictValue2):

    sp_pa_audict[dictKey][dictKey2] = dictValue2


def proc_inpt(fileIn, cdatein):

    wb = open_workbook(fileIn, on_demand=True)

    sheetCnt = wb.nsheets

    print("Number of Worksheets: ", sheetCnt)

    for s in wb.sheets():
        print('Sheet Name: ', s.name)

    sheet = wb.sheet_by_index(0)

    rowCnt = sheet.nrows

    print("Value of Row Count: ", rowCnt)

    for row in range(1, sheet.nrows):
        tstrslt = ""
        rowVal = sheet.row_values(row)
        dt = rowVal[0]
        ciDate = cnvt_excel_dt(dt)
        tstrslt = math.isnan(float(rowVal[18]))
        if tstrslt is False:
            dt = rowVal[18]
            rtDate = cnvt_excel_dt(dt)
        else:
            print("Test Result is True:", rowVal[18])
            rtDate = 0
        dtDiff = datetime.timedelta(cdatein - int(rowVal[18]))
        dictKey = str(int(rowVal[4])) + "." + str(rtDate)
        if row == 1:
            print("first time")
            ciUpdate = "CI_Updt_" + str(ciDate)
            print("Check In date: ", ciUpdate)
        dictKey2 = str(ciDate)
        dictValue = str(rowVal[4]) + "." + str(rtDate)
        dictValue2 = str(rowVal[19])
        write_dict1(dictKey, dictValue, dictKey2, dictValue2)

    wb.release_resources()


lstDir = "C:/Users/e1ze4m0/Documents/Indivior Program/Indivior Hub/MS Excel/SP Status/Input/"

fl_list = os.listdir(lstDir)

sp_pa_audict = defaultdict(dict)

wbout = Workbook()
sheet1 = wbout.add_sheet('SP Prior Auths 1', cell_overwrite_ok=True)

for x in fl_list:
    filein = lstDir + x
    print(filein, sep="\n")
    print("value of x: ", x)
    cdatein = dt_tbl()
    proc_inpt(filein, cdatein)
    bld_output(sp_pa_audict, cdatein)

wbout.save('xlwt SP_PA3.xls')

print("delete dictionary")
del sp_pa_audict
