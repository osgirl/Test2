import pypyodbc

conn =
pypyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};r'DBQ=C:\Users\e1ze4m0\Documents\PlanPayerScrub.accdb;')
cursor = conn.cursor()
cursor.execute('select * from tracking_sales')

for row in cursor.fetchall():
    print(row)
