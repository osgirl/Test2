import os
import pandas as pd
import datetime

def cnvt_excel_dt(dt):

    dateoffset = 693594
    return datetime.date.fromordinal(dateoffset + int(dt))


def add_checkins():


    cdate = datetime.date.today()
    pycurdt = datetime.date.toordinal(cdate)
    cdate2 = pycurdt - 693594
    #cdate2 = float(str(cdate))

    print("Today's Date is: ", cdate)
    print("Today's Date as ordinal value: ", cdate2)
    print("************** Start List *****************")

    for x in range(60, 0, -1):
        mdate = cdate2 - x
        dt = mdate
        mdate2 = cnvt_excel_dt(dt)
        newcol = "ci_dt_" + str(mdate2)
        print("value of mdate: ", mdate, " display value: ", mdate2)

def do_df():

    newDF = pd.DataFrame()
#    newDF = newDF.append(other= , ignore_index = True)
    # try printing some data from newDF
    print(newDF.head())
    return newDF

def do_excel(fileIn):

#    df = pd.read_excel(fileIn, sheet_name="sheet1")

    excel_file = fileIn
    pa_sp = pd.read_excel(excel_file, index_col=None, na_values=['NA'], usecols = "A,E,S,T")

    print("column Headings")
    print(pa_sp.index)
    print(pa_sp.columns)

    for index, row in pa_sp.iterrows():
        case = row[1]
        routedt1 = row[2]
        routedt2 = str(routedt1).split(" ")
        print(case, " / ", str(routedt2[0]))
        print("row value - ", row[1], " index value - ", index)
    fnd = 0
    lp = 0
    if len(newDF) == 0:
        print("at len 0")
        newDF.at[lp, 'A'] = case
        newDF.at[lp, '2'] = routedt2[0]
        print(newDF)
    else:
        while lp < len(newDF):
            print("new df len: ", len(newDF), lp)
            print(newDF)
            if newDF.at[lp, 'A'] == case & newDF.at[lp, 'B'] == routedt2[0]:
                print("search key found")
            else:
                print("no search key found")
                newDF.at[lp, '1'] = case
                newDF.at[lp, '2'] = routedt2[0]
            lp = lp + 1
    print(newDF)

#        += lp



def process_input_files():
    for x in fl_list:
        filein = lstDir + x
        print(filein, sep="\n")
        print("value of x: ", x)
        do_excel(filein)

lstDir = "C:/Users/e1ze4m0/Documents/Indivior Program/Indivior Hub/MS Excel/SP Status/Input/"

fl_list = os.listdir(lstDir)

add_checkins()
newDF = do_df()
process_input_files()

