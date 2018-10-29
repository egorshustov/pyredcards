import pyodbc
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import selenium.webdriver.support.ui as ui
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time
import datetime
import locale

class League:
    def __init__(self, league_name, url_whoscored, url_championat,):
        self.league_name = league_name
        self.url_whoscored = url_whoscored
        self.url_championat = url_championat
        self.url_past_season = ""
        self.url_referee_statistics = ""
        self.matches_found = True

class Match:
    def __init__(self):
        self.team_home_url = ""
        self.team_away_url = ""
        self.match_url = "" #ссылка на личную встречу команд
        self.league_name = "" #Лига
        self.team_home_name = "" #Дома
        self.team_away_name = "" #Гости
        self.match_datetime = "" #Дата
        self.referee_name_championat = "" #Судья, Имя на championat
        self.referee_name_championat_translated_to_en = ""
        self.referee_name_whoscored = "" #Судья, Имя на whoscored
        self.referee_url = ""  #Судья, URL на whoscored
        self.referee_this_season_average = "" #Судья, Этот сезон
        self.referee_this_season_matches_count = -1
        self.referee_all_seasons_average = "" #Судья, Все сезоны
        self.referee_all_seasons_matches_count = -1
        self.referee_to_team_home_average = "" #Судья, Командам
        self.referee_to_team_away_average = ""
        self.referee_last_twenty_home_count = 0 #Судья, Посл. 20 игр
        self.referee_last_twenty_away_count = 0
        self.referee_last_twenty_last_kk_date = ""
        self.team_home_kk_this_season_count = 0 #Команды, КК этот сезон
        self.team_away_kk_this_season_count = 0
        self.team_home_found_in_last_season = False #Команды, КК предыдущий сезон
        self.team_home_kk_last_season_count = 0
        self.team_away_found_in_last_season = False
        self.team_away_kk_last_season_count = 0
        self.team_home_last_kk_date = "" #Команды, Дата посл 1
        self.team_away_last_kk_date = "" #Команды, Дата посл 2
        self.team_home_personal_meetings_kk_count_home = 0 #Команды, Личн. вст
        self.team_home_personal_meetings_kk_count_away = 0
        self.team_away_personal_meetings_kk_count_home = 0
        self.team_away_personal_meetings_kk_count_away = 0
        self.teams_personal_meetings_last_kk_date = ""
        self.teamsstring = ""

def datestring_to_unix(datestring):
    # Преобразует строку типа 'суббота, окт 27 2018' в unix-формат.
    # Удалим день недели:
    datestring = datestring[datestring.find(', ')+2:]
    # Преобразуем строку в datetime:
    datestring_dt = datetime.datetime.strptime(datestring, u'%b %d %Y')
    # Преобразуем datetime в unix:
    datestring_unix = time.mktime(datestring_dt.timetuple())
    return datestring_unix

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

def main():
    # Подключимся к БД Microsoft Access через экземпляр ODBC
    db = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\\RC_base2_lessdata.mdb')
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

    #write_to_spreadsheets()

    includes_path = "C:/Users/BRAXXMAN/source/repos/includes/"
    os.environ["webdriver.chrome.driver"] = includes_path+"chromedriver.exe"
    driver = webdriver.Chrome(executable_path=includes_path+'chromedriver.exe')
    #driver=webdriver.Firefox()
    
    required_date = 'четверг, ноя 1 2018'
    #required_date = 'суббота, ноя 3 2018'
    required_date_unix = datestring_to_unix(required_date)
    match = []

    # Для каждой лиги переходим на страницу Календаря Игр сайта whoscored:
    for i in range(0,league_length):
        next_clicked = True
        driver.get(league[i].url_whoscored)
        while next_clicked != False:
            # Получим всю таблицу календаря игр:
            time.sleep(sleep_time) # Пауза для прогрузки таблицы
            tournament_fixture = ui.WebDriverWait(driver, 15).until(lambda driver: driver.find_element_by_id('tournament-fixture'))
            tournament_fixture_innerhtml = tournament_fixture.get_property('innerHTML')
            #print(tournament_fixture_innerhtml)
            if required_date in tournament_fixture_innerhtml:
                # Если искомая дата присутствует в таблице, спарсим все матчи на искомую дату для данной лиги.
                # Спарсим все дни с матчами в таблице:
                days = []
                days = tournament_fixture_innerhtml.split('<tr class=\"rowgroupheader\"><th colspan=\"7\">')
                # Пройдёмся по всем дням в таблице и найдём тот день, в котором присуствует искомая дата:
                for day in days:
                    if required_date in day:
                        # Это искомый день. Спарсим всю информацию по матчам для искомого дня данной лиги и выйдем из цикла:
                        teams_home = []
                        teams_away = []
                        match_urls = []
                        soup = BeautifulSoup(day)
                        teams_home = soup.findAll('td', { 'class' : 'team home' })
                        teams_away = soup.findAll('td', { 'class' : 'team away' })
                        match_urls = soup.findAll('a', { 'class' : 'result-4 rc' })
                        #for j in range(0,len(teams_home))
                        #    teams_away[j]

                        break
            else:
                rowgroupheaders = []
                rowgroupheaders = tournament_fixture.find_elements_by_class_name('rowgroupheader')
                # Возьмём последний rowgroupheader в таблице и достанем из него дату:
                # (также уберём из его innerText последний символ (символ перевода строки через операцию среза [0:-1]):
                last_date_in_table = rowgroupheaders[-1].get_property('innerText')[0:-1]
                last_date_in_table_unix = datestring_to_unix(last_date_in_table)
                if last_date_in_table_unix > required_date_unix:
                    # Если последний день в текущем диапазоне дней больше указанного
				    # (и при этом матчей до этого дня (в условии if required_date in tournament_fixture_innerhtml не было обнаружено),
				    # то делаем вывод, что в текущей лиге не нашлось матчей в указанный день:
                    next_clicked = False # дальше таблицу не листаем
                    league[i].matches_found = False
                    print('В ' + league[i].league_name + ' на ru.whoscored.com нет матчей в указанный день!')
                else:
                    # Если последний день в текущем диапазоне дней ещё меньше указанного
				    # (и при этом матчей до этого дня (в условии if required_date in tournament_fixture_innerhtml не было обнаружено),
				    # то пролистаем таблицу дальше:
                    driver.find_element_by_css_selector('.next').click()

    #driver.close()
    #sleep(10)

if __name__ == '__main__':
    sleep_time = 1
    # Получаем текущую локаль:
    default_loc = locale.getlocale()
    # Изменяем локаль для корректной конвертации строки с русской датой в datetime:
    locale.setlocale(locale.LC_ALL, ('RU','UTF8'))
    # Вызываем главную функцию:
    main()
    # Меняем локаль назад:
    locale.setlocale(locale.LC_ALL, default_loc)