
import csv, pyodbc
from pyodbc import *

class League:
    def __init__(self, league_name, url_whoscored, url_championat):
        self.league_name = league_name
        self.url_whoscored = url_whoscored
        self.url_championat = url_championat

db = connect('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\\RC_base.mdb')
dbc = db.cursor()
rows = dbc.execute('select * from [Leagues]').fetchall()
db.close()


league_length = len(rows)
league = []
for i in range(0,league_length):
    league.append(League(rows[i][1],rows[i][2],rows[i][3]))

