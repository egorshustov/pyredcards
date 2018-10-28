import pyodbc
import gspread
from oauth2client.service_account import ServiceAccountCredentials

import os
import selenium.webdriver.support.ui as ui
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from time import sleep

from bs4 import BeautifulSoup

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

def main():
    # Подключимся к БД Microsoft Access через экземпляр ODBC
    db = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\\RC_base2.mdb')
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

    required_date = 'воскресенье, окт 28 2018'
    #required_date = 'суббота, янв 19 2019'

    # Для каждой лиги переходим на страницу Календаря Игр сайта whoscored:
    for i in range(0,league_length):
        driver.get(league[i].url_whoscored)
        # Получим всю таблицу календаря игр:
        tournament_fixture = ui.WebDriverWait(driver, 15).until(lambda driver: driver.find_element_by_id('tournament-fixture'))
        tournament_fixture_innerhtml = tournament_fixture.get_property('innerHTML')
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
                    break
        else:
            driver.find_element_by_css_selector('.next').click()

    # Получаем указатель на поле ввода текста в форме постинга
    textarea=driver.find_element_by_css_selector('#index_email')
    # Печатаем в поле ввода какой-либо текст
    textarea.send_keys('shustov_egor@mail.ru')

    # Получаем указатель на поле ввода пароля
    textarea=driver.find_element_by_css_selector('#index_pass')
    # Печатаем в поле ввода пароль
    textarea.send_keys('Password')

    #Получаем указатель на кнопку "Войти"
    submit=driver.find_element_by_css_selector('#index_login_button')
    submit.click()
    #Ждём пока загрузится кнопка "диалоги"
    messages_button = ui.WebDriverWait(driver, 15).until(lambda driver: driver.find_element_by_id('l_msg'))
    #Нажимаем эту кнопку через Ctrl (в новой вкладке)
    ActionChains(driver) \
        .key_down(Keys.CONTROL) \
        .click(messages_button) \
        .key_up(Keys.CONTROL) \
        .perform()
    #driver.close()
    #sleep(10)

if __name__ == '__main__':
    main()