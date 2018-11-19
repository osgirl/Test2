import datetime, time

testdict = {}
testdict['testdict2'] = {}
print("test print 1")
print(testdict)

for x in range(1, 30, 1):
    newkey = "case:" + str(x)
    tm = time.perf_counter_ns()
    newkey2 = "ciDate:" + str(x)
    str_tm = str(tm)
    print(str_tm)
    newvalue = "ba" + str_tm
    newvalue2 = "xx2" + str_tm
    testdict[newkey] = newvalue
    testdict['testdict2'][newkey2] = newvalue2


print("test print 2")
print(testdict)

for x in testdict:
    print(x)

for x, y in testdict.items():
    print(x, y)
