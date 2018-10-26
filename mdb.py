
import csv, pyodbc
from pyodbc import *
db = connect('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\\RC_base.mdb') #order.mdb - собственно мое файло БД
dbc = db.cursor()
rows = dbc.execute('select * from [Leagues]').fetchall()
db.close()


# you could change the mode from 'w' to 'a' (append) for any subsequent queries
with open('report.csv', 'w') as fou:
    csv_writer = csv.writer(fou) # default field-delimiter is ","
    csv_writer.writerows(rows)