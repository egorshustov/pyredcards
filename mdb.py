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
import sys
import PyQt5.QtGui as QtGui
from PyQt5.QtCore import pyqtSlot, QThread
from PyQt5.QtWidgets import (QApplication,QDialog,QMainWindow,QWidget, QCalendarWidget)
from PyQt5.uic import loadUi

class Window(QMainWindow):

    def __init__(self):
        super(Window, self).__init__()
        # Загрузим UI из файла:
        loadUi('redcardsdesigner.ui', self)
        # Создадим обработчик для кнопки:
        self.startButton.clicked.connect(self.on_startbutton_clicked)
        # Инициализируем объект класса нити:
        self.workerThread = WorkerThread()
        # Запустим форму окна:
        self.show()

    def on_startbutton_clicked(self):
        # При нажатии на кнопку запустим нить workerThread:
        # self.workerThread.start()
        # cal = self.calendarWidget
        main()

class WorkerThread(QThread):

    def __init__(self):
        super(WorkerThread, self).__init__()
    
    def run(self):
        # Вызываем главную функцию:
        #main()
        time.sleep(1)
       
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
        self.personal_meetings_count = 0
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
    #os.environ["webdriver.chrome.driver"] = includes_path+"chromedriver.exe"
    #driver = webdriver.Chrome(executable_path=includes_path+'chromedriver.exe')
    driver = webdriver.Firefox(executable_path=includes_path+'geckodriver.exe')
    #required_date = 'четверг, ноя 1 2018'

    #Получим объект calendarWidget с окна программы wind:
    cal = wind.calendarWidget
    # Получим выбранную дату и выполним преобразование из QDate в datetime:
    required_date = cal.selectedDate().toPyDate()
    # Преобразуем datetime в unix:
    required_date_unix = time.mktime(required_date.timetuple())

    match = []
    i_match = 0
    ##################################################################################
    ################################# МАТЧИ ЗА ДЕНЬ
    ##################################################################################
    print('Найдём матчи для каждой из лиг в указанный день:')
    # Для каждой лиги переходим на страницу Календаря Игр сайта whoscored:
    for i in range(0,league_length):
        next_clicked = True
        driver.get(league[i].url_whoscored)
        time.sleep(sleep_page_time) # Пауза для прогрузки страницы:
        while next_clicked != False:
            # Получим всю таблицу календаря игр:
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
                        match_datetime = []
                        teams_home = []
                        teams_away = []
                        match_urls = []
                        soup = BeautifulSoup(day, 'lxml')
                        match_datetime = soup.findAll('td', {'class' : 'time'})
                        teams_home = soup.findAll('td', {'class' : 'team home'})
                        teams_away = soup.findAll('td', {'class' : 'team away'})
                        match_urls = soup.findAll('a', {'class' : 'result-4 rc'})
                        # Для каждого матча искомого дня текущей лиги получим его данные и занесём в массив match[]:
                        for j in range(0,len(teams_home)):
                            match.append(Match())
                            match[i_match].league_name = league[i].league_name
                            match[i_match].match_datetime = required_date + " " + match_datetime[j].text
                            match[i_match].team_home_url = 'https://ru.whoscored.com'+teams_home[j].find('a', {'class' : 'team-link '})['href']
                            match[i_match].team_home_name = teams_home[j].find('a', {'class' : 'team-link '}).text
                            match[i_match].team_away_url = 'https://ru.whoscored.com'+teams_away[j].find('a', {'class' : 'team-link '})['href']
                            match[i_match].team_away_name = teams_away[j].find('a', {'class' : 'team-link '}).text
                            match[i_match].match_url = 'https://ru.whoscored.com'+match_urls[j]['href']
                            next_clicked = False  # дальше таблицу не листаем
                            i_match = i_match + 1
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
                    time.sleep(sleep_table_time) # Пауза для прогрузки таблицы
    
    # Определим длину списка matches:
    matches_length = len(match)
    ##################################################################################
    ################################# ЛИЧНЫЕ ВСТРЕЧИ
    ##################################################################################
    print('Спарсим информацию личных встреч команд для каждого найденного матча:')
    for i in range(0,matches_length):
        driver.get(match[i].match_url)
        print('Команды '+match[i].team_home_name+' и '+match[i].team_away_name)
        time.sleep(sleep_page_time)
		# Проверим наличие таблицы предыдущих встреч двух команд
		# по тексту заголовка previous-meetings-count
        previous_meetings_count = ui.WebDriverWait(driver, 15).until(lambda driver: driver.find_element_by_id('previous-meetings-count')).get_property('innerText')
        # Если текст заголовка previous-meetings-count не пустой (например, "(Последние N матчей)")
        if previous_meetings_count != '':
            # Получим таблицу предыдущих встреч двух команд:
            previous_meetings_grid = ui.WebDriverWait(driver, 15).until(lambda driver: driver.find_element_by_id('previous-meetings-grid')).get_property('innerHTML')
            soup = BeautifulSoup(previous_meetings_grid, 'lxml')
            # Получим все предыдущие встречи двух команд и занесём их в список:
            previous_matches = []
            previous_matches = soup.findAll('tr', {'class' : 'item'})
            # Определим количество этих предыдущих матчей:
            match[i].personal_meetings_count = len(previous_matches)
            # Пройдемся в цикле по каждому из предыдущих матчей:
            kk_found = False
            for previous_match in previous_matches:
                # Если в матче найдена КК:
                if previous_match.find('span', {'class' : 'rcard ls-e'}) != None:
                    # Если матчей с КК ещё не было найдено, то занесём в match[i] дату последней КК
                    if kk_found == False:
                        kk_found = True
                        match[i].teams_personal_meetings_last_kk_date = previous_match.find('td', {'class' : 'date'}).text
                    # Спарсим класс домашней команды для данного матча:
                    team_home = previous_match.find('td', 'home') # Найдём тег td, содержащий класс home
                    if team_home.find('span', {'class' : 'rcard ls-e'}) != None:
                        if match[i].team_home_name in team_home.text:
                            match[i].team_home_personal_meetings_kk_count_home += int(team_home.find('span', {'class' : 'rcard ls-e'}).text)
                        if match[i].team_away_name in team_home.text:
                            match[i].team_away_personal_meetings_kk_count_home += int(team_home.find('span', {'class' : 'rcard ls-e'}).text) 
                    # Спарсим класс гостевой команды для данного матча:
                    team_away = previous_match.find('td', 'away') # Найдём тег td, содержащий класс away
                    if team_away.find('span', {'class' : 'rcard ls-e'}) != None:
                        if match[i].team_home_name in team_away.text:
                            match[i].team_home_personal_meetings_kk_count_away += int(team_away.find('span', {'class' : 'rcard ls-e'}).text)
                        if match[i].team_away_name in team_away.text:
                            match[i].team_away_personal_meetings_kk_count_away += int(team_away.find('span', {'class' : 'rcard ls-e'}).text) 
        else:
            # Если текст заголовка previous-meetings-count пустой
            # пометим все атрибуты личных встреч как -1:
            match[i].team_home_personal_meetings_kk_count_home = -1
            match[i].team_home_personal_meetings_kk_count_away = -1
            match[i].team_away_personal_meetings_kk_count_home = -1
            match[i].team_away_personal_meetings_kk_count_away = -1
            print('У команд '+match[i].team_home_name+' и '+match[i_match].team_away_name+' не было совместных встреч!')

    #driver.close()
    time.sleep(1)

if __name__ == '__main__':
    sleep_page_time = 5
    sleep_table_time = 1
    # Получаем текущую локаль:
    default_loc = locale.getlocale()
    # Изменяем локаль для корректной конвертации строки с русской датой в datetime:
    locale.setlocale(locale.LC_ALL, ('RU','UTF8'))

    app = QApplication(sys.argv)
    wind = Window()
    app.exec_()

    # Меняем локаль назад:
    locale.setlocale(locale.LC_ALL, default_loc)