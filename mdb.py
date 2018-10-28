from pyodbc import *
import gspread
from oauth2client.service_account import ServiceAccountCredentials

class League:
    def __init__(self, league_name, url_whoscored, url_championat):
        self.league_name = league_name
        self.url_whoscored = url_whoscored
        self.url_championat = url_championat

def write_to_spreadsheets():
    # Создаём Service-объект, для работы с Google-таблицами:
    CREDENTIALS_FILE = 'RedCardsProject-90325d995892.json'  # имя выгруженного файла с закрытым ключом
    # В Scope укажем к каким API мы хотим получить доступ:
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    # Заполним массив с учётными данными:
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    # Авторизуемся с этими учётными данными:
    gc = gspread.authorize(credentials)
    # Откроем Spreadsheet с указанным именем:
    sh = gc.open('Python')
    # Получим первый лист этого Spreadsheet:
    worksheet = sh.sheet1
    # удалить лист из файла (удалится только если лист не единственный):
    #sh.del_worksheet(worksheet)
    # предоставить себе роль владельца файла:
    # sh.share('egorshustov.93@gmail.com', perm_type='user', role='owner')
    # Запишем каждую строку списка league[] в лист worksheet:
    for i in range(0,league_length):
        worksheet.append_row([league[i].league_name,league[i].url_whoscored,league[i].url_championat])

# Подключимся к БД Microsoft Access через экземпляр ODBC
db = connect('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\\RC_base.mdb')
dbc = db.cursor()
# Получим информацию из БД, занесём её в rows
rows = dbc.execute('select * from [Leagues]').fetchall()
db.close()

# Определим длину списка rows:
league_length = len(rows)
# Инициализируем список league[] и заполним его экземплярами класса League:
league = []
for i in range(0,league_length):
    league.append(League(rows[i][1],rows[i][2],rows[i][3]))

write_to_spreadsheets()
