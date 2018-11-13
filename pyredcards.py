import pyodbc
import httplib2
import apiclient.discovery
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
from PyQt5.QtCore import pyqtSlot, QThread, Qt
from PyQt5.QtWidgets import (QApplication, QDialog, QMainWindow, QWidget, QCalendarWidget)
from PyQt5.uic import loadUi


class Window(QMainWindow):

    def __init__(self):
        super(Window, self).__init__(flags=Qt.WindowFlags())
        # Загрузим UI из файла:
        loadUi('redcardsdesigner.ui', self)

        # Создадим обработчик для кнопки:
        self.startButton.clicked.connect(self.on_startbutton_clicked)
        self.reportButton.clicked.connect(self.on_reportbutton_clicked)
        # Инициализируем объект класса нити:
        self.workerThread = WorkerThread()
        self.reportThread = ReportThread()
        # Запустим форму окна:
        self.show()

    def on_startbutton_clicked(self):
        # При нажатии на кнопку запустим нить workerThread:
        self.workerThread.start()
        # cal = self.calendarWidget

    def on_reportbutton_clicked(self):
        # При нажатии на кнопку запустим нить workerThread:
        self.reportThread.start()


class WorkerThread(QThread):

    def __init__(self):
        super(WorkerThread, self).__init__()

    def run(self):
        # Вызываем главную функцию:
        main()


class ReportThread(QThread):

    def __init__(self):
        super(ReportThread, self).__init__()

    def run(self):
        write_to_spreadsheets()


class League:

    def __init__(self, league_name, url_whoscored, url_championat, ):
        self.league_name = league_name
        self.url_whoscored = url_whoscored
        self.url_championat = url_championat
        self.url_past_season = ''
        self.url_referee_statistics = ''
        self.url_games_calendar_past_season = ''
        self.matches_found = True


class Match:

    def __init__(self):
        self.team_home_url = ''
        self.team_away_url = ''
        self.match_url = ''  # ссылка на личную встречу команд
        self.league_name = ''  # Лига
        self.team_home_name = ''  # Дома
        self.team_away_name = ''  # Гости
        self.match_datetime = ''  # Дата
        self.referee_name_championat = ''  # Судья, Имя на championat
        self.referee_name_championat_translated_to_en = ''
        self.referee_name_whoscored = ''  # Судья, Имя на whoscored
        self.referee_url = ''  # Судья, URL на whoscored
        self.referee_this_season_average = ''  # Судья, Этот сезон
        self.referee_this_season_matches_count = -1
        self.referee_all_seasons_average = ''  # Судья, Все сезоны
        self.referee_all_seasons_matches_count = -1
        self.referee_to_team_home_average = ''  # Судья, Командам
        self.referee_to_team_away_average = ''
        self.referee_last_twenty_home_count = 0  # Судья, Посл. 20 игр
        self.referee_last_twenty_away_count = 0
        self.referee_last_twenty_last_kk_date = ''
        self.team_home_kk_this_season_count = 0  # Команды, КК этот сезон
        self.team_away_kk_this_season_count = 0
        self.team_home_found_in_last_season = False  # Команды, КК предыдущий сезон
        self.team_home_kk_last_season_count = 0
        self.team_away_found_in_last_season = False
        self.team_away_kk_last_season_count = 0
        self.team_home_last_kk_date = ''  # Команды, Дата посл 1
        self.team_away_last_kk_date = ''  # Команды, Дата посл 2
        self.team_home_personal_meetings_kk_count_home = 0  # Команды, Личн. вст
        self.team_home_personal_meetings_kk_count_away = 0
        self.team_away_personal_meetings_kk_count_home = 0
        self.team_away_personal_meetings_kk_count_away = 0
        self.personal_meetings_count = 0
        self.teams_personal_meetings_last_kk_date = ''
        self.teamsstring = ''


def datestring_to_unix(datestring):
    # Преобразует строку типа 'суббота, окт 27 2018' в unix-формат.
    # Удалим день недели:
    datestring = datestring[datestring.find(', ') + 2:]
    # Преобразуем строку в datetime:
    datestring_dt = datetime.datetime.strptime(datestring, u'%b %d %Y')
    # Преобразуем datetime в unix:
    datestring_unix = time.mktime(datestring_dt.timetuple())
    return datestring_unix


def datestring_format(datestring):
    # Преобразует строку типа 'суббота, окт 27 2018' в строку типа '27-10-2018'.
    # Удалим день недели:
    datestring = datestring[datestring.find(', ') + 2:]
    # Преобразуем строку в datetime:
    datestring_dt = datetime.datetime.strptime(datestring, u'%b %d %Y')
    # Преобразуем datetime в строку:
    datestring_formatted = datestring_dt.strftime("%d-%m-%Y")
    return datestring_formatted


def get_matches():
    ##################################################################################
    # МАТЧИ ЗА ДЕНЬ
    ##################################################################################
    print('Найдём матчи для каждой из лиг в указанный день:')
    global match
    i_match = 0
    # Для каждой лиги переходим на страницу Календаря Игр сайта whoscored:
    for i in range(0, league_length):
        next_clicked = True
        driver.get(league[i].url_whoscored)
        time.sleep(sleep_page_time)  # Пауза для прогрузки страницы
        print('Лига ' + league[i].league_name + '...')

        # Параллельно получим ссылку на прошлый сезон каждой лиги:
        league[i].url_past_season = 'https://ru.whoscored.com' +\
                                    driver.find_element_by_css_selector('#seasons'
                                                                        ' > option:nth-child(2)').get_property('value')
        if league[i].league_name == 'Англия 2':  # Исключение для сезона Англия 2 - задаём вручную
            league[i].url_past_season = 'https://ru.whoscored.com/Regions/252/Tournaments/7/' \
                                        'Seasons/6848/Stages/15177/Show/Англия-2-2017-2018'

        # Параллельно получим ссылку на статистику судей каждой лиги:
        league[i].url_referee_statistics = driver.find_element_by_css_selector('#sub-navigation > ul:nth-child(1) >'
                                                                               ' li:nth-child(5) > a:nth-child(1)') \
            .get_property('href')

        while next_clicked is True:
            # Получим всю таблицу календаря игр:
            tournament_fixture = ui.WebDriverWait(driver, 15).until(
                lambda driver1: driver.find_element_by_id('tournament-fixture'))
            tournament_fixture_innerhtml = tournament_fixture.get_property('innerHTML')
            # print(tournament_fixture_innerhtml)
            if required_date in tournament_fixture_innerhtml:
                # Если искомая дата присутствует в таблице, спарсим все матчи на искомую дату для данной лиги.
                # Спарсим все дни с матчами в таблице:
                days = tournament_fixture_innerhtml.split('<tr class="rowgroupheader"><th colspan="7">')
                # Пройдёмся по всем дням в таблице и найдём тот день, в котором присуствует искомая дата:
                for day in days:
                    if required_date in day:
                        # Это искомый день. Спарсим всю информацию по матчам для искомого дня
                        # данной лиги и выйдем из цикла:
                        soup = BeautifulSoup(day, 'html.parser')
                        match_datetime = soup.findAll('td', {'class': 'time'})
                        teams_home = soup.findAll('td', {'class': 'team home'})
                        teams_away = soup.findAll('td', {'class': 'team away'})
                        match_urls = soup.findAll('a', {'class': 'result-4 rc'})
                        # Для каждого матча искомого дня текущей лиги получим его данные и занесём в массив match[]:
                        for j in range(0, len(teams_home)):
                            match.append(Match())
                            match[i_match].league_name = league[i].league_name
                            match[i_match].match_datetime = match_datetime[j].text
                            match[i_match].team_home_url = 'https://ru.whoscored.com' + \
                                                           teams_home[j].find('a', {'class': 'team-link '})['href']
                            match[i_match].team_home_name = teams_home[j].find('a', {'class': 'team-link '}).text
                            match[i_match].team_away_url = 'https://ru.whoscored.com' + \
                                                           teams_away[j].find('a', {'class': 'team-link '})['href']
                            match[i_match].team_away_name = teams_away[j].find('a', {'class': 'team-link '}).text
                            match[i_match].match_url = 'https://ru.whoscored.com' + match_urls[j]['href']
                            next_clicked = False  # дальше таблицу не листаем
                            i_match = i_match + 1
                        break
            else:
                rowgroupheaders = tournament_fixture.find_elements_by_class_name('rowgroupheader')
                # Возьмём последний rowgroupheader в таблице и достанем из него дату:
                # (также уберём из его innerText символ перевода строки:
                last_date_in_table = rowgroupheaders[-1].get_property('innerText')
                last_date_in_table = last_date_in_table.replace('\n', '')

                last_date_in_table_unix = datestring_to_unix(last_date_in_table)
                if last_date_in_table_unix > required_date_unix:
                    # Если последний день в текущем диапазоне дней больше указанного
                    # (и при этом матчей до этого дня
                    # (в условии if required_date in tournament_fixture_innerhtml)
                    # не было обнаружено),
                    # то делаем вывод, что в текущей лиге не нашлось матчей в указанный день:
                    next_clicked = False  # дальше таблицу не листаем
                    league[i].matches_found = False
                    print('В ' + league[i].league_name + ' на ru.whoscored.com нет матчей в указанный день!')
                else:
                    # Если последний день в текущем диапазоне дней ещё меньше указанного
                    # (и при этом матчей до этого дня
                    # (в условии if required_date in tournament_fixture_innerhtml)
                    # не было обнаружено),
                    # И если кнопка next активна,
                    # то пролистаем таблицу дальше.
                    # Проверим, активна ли кнопка "предыдущий месяц". Получим перечень её классов:
                    classes_of_button = driver.find_element_by_css_selector('.next').get_property('className')
                    if 'is-disabled' in classes_of_button:
                        # Если кнопка неактивна, то продолжать листать календарь вперёд уже нельзя,
                        # И это значит, что сезон завершился.
                        next_clicked = False  # дальше таблицу не листаем
                        league[i].matches_found = False
                        print('Сезон лиги ' + league[i].league_name + ' уже завершился на указанный день!')
                    else:
                        # Если кнопка активна, перейдём на предыдущий месяц календаря:
                        driver.find_element_by_css_selector('.next').click()
                        time.sleep(sleep_table_time)  # Пауза для прогрузки таблицы
    # Определим длину списка matches:
    global matches_length
    matches_length = len(match)
    print('Поиск матчей для каждой из лиг в указанный день завершен.')


def get_personal_meetengs():
    ##################################################################################
    # ЛИЧНЫЕ ВСТРЕЧИ
    ##################################################################################
    print('Спарсим информацию личных встреч команд для каждого найденного матча:')
    for i in range(0, matches_length):
        driver.get(match[i].match_url)
        print('Команды ' + match[i].team_home_name + ' и ' + match[i].team_away_name)
        time.sleep(sleep_page_time)
        # Проверим наличие таблицы предыдущих встреч двух команд
        # по тексту заголовка previous-meetings-count
        previous_meetings_count = ui.WebDriverWait(driver, 15).until(
            lambda driver1: driver.find_element_by_id('previous-meetings-count')).get_property('innerText')
        # Если текст заголовка previous-meetings-count не пустой (например, '(Последние N матчей)')
        if previous_meetings_count != '':
            # Получим таблицу предыдущих встреч двух команд:
            previous_meetings_grid = ui.WebDriverWait(driver, 15).until(
                lambda driver1: driver.find_element_by_id('previous-meetings-grid')).get_property('innerHTML')
            soup = BeautifulSoup(previous_meetings_grid, 'html.parser')
            # Получим все предыдущие встречи двух команд и занесём их в список:
            previous_matches = soup.findAll('tr', {'class': 'item'})
            # Определим количество этих предыдущих матчей:
            match[i].personal_meetings_count = len(previous_matches)
            # Пройдемся в цикле по каждому из предыдущих матчей:
            kk_found = False
            for previous_match in previous_matches:
                # Если в матче найдена КК:
                if previous_match.find('span', {'class': 'rcard ls-e'}) is not None:
                    # Если матчей с КК ещё не было найдено, то занесём в match[i] дату последней КК
                    if kk_found is False:
                        kk_found = True
                        match[i].teams_personal_meetings_last_kk_date = previous_match.find('td',
                                                                                            {'class': 'date'}).text
                    # Спарсим класс домашней команды для данного матча:
                    team_home = previous_match.find('td', 'home')  # Найдём тег td, содержащий класс home
                    if team_home.find('span', {'class': 'rcard ls-e'}) is not None:
                        if match[i].team_home_name in team_home.text:
                            match[i].team_home_personal_meetings_kk_count_home += int(
                                team_home.find('span', {'class': 'rcard ls-e'}).text)
                        if match[i].team_away_name in team_home.text:
                            match[i].team_away_personal_meetings_kk_count_home += int(
                                team_home.find('span', {'class': 'rcard ls-e'}).text)
                            # Спарсим класс гостевой команды для данного матча:
                    team_away = previous_match.find('td', 'away')  # Найдём тег td, содержащий класс away
                    if team_away.find('span', {'class': 'rcard ls-e'}) is not None:
                        if match[i].team_home_name in team_away.text:
                            match[i].team_home_personal_meetings_kk_count_away += int(
                                team_away.find('span', {'class': 'rcard ls-e'}).text)
                        if match[i].team_away_name in team_away.text:
                            match[i].team_away_personal_meetings_kk_count_away += int(
                                team_away.find('span', {'class': 'rcard ls-e'}).text)
        else:
            # Если текст заголовка previous-meetings-count пустой
            # пометим все атрибуты личных встреч как -1:
            match[i].team_home_personal_meetings_kk_count_home = -1
            match[i].team_home_personal_meetings_kk_count_away = -1
            match[i].team_away_personal_meetings_kk_count_home = -1
            match[i].team_away_personal_meetings_kk_count_away = -1
            print('У команд ' + match[i].team_home_name + ' и '
                  + match[i].team_away_name + ' не было совместных встреч!')
    print('Информация личных встреч команд для каждого найденного матча получена.')


def get_url_games_calendar_past_season():
    print('Достанем ссылки на календарь игр прошлых сезонов:')
    for i in range(0, league_length):
        if league[i].matches_found is True:
            driver.get(league[i].url_past_season)
            time.sleep(sleep_page_time)
            league[i].url_games_calendar_past_season = \
                driver.find_element_by_css_selector('#sub-navigation > ul:nth-child(1) >'
                                                    ' li:nth-child(2) > a:nth-child(1)').get_property('href')
    print('Ссылки успешно получены.')


def get_kk_this_or_last_season(this_season):
    ##################################################################################
    # КК ЗА ЭТОТ ИЛИ ПРЕДЫДУЩИЙ СЕЗОН, ДАТА ПОСЛЕДНЕЙ КК
    ##################################################################################
    if this_season:
        print('Получим информацию о КК за текущий сезон:')
    else:
        print('Получим информацию о КК за прошлый сезон:')

    for i in range(0, league_length):
        if league[i].matches_found is True:
            previous_clicked = True
            if this_season:
                driver.get(league[i].url_whoscored)
            else:
                driver.get(league[i].url_games_calendar_past_season)
            time.sleep(sleep_page_time)

            while previous_clicked is True:
                # Определим месяц (написан на кнопке):
                current_month = driver.find_element_by_css_selector('span.text:nth-child(1)').get_property('innerHTML')
                if this_season:
                    print('Текущий сезон лиги ' + league[i].league_name + ', месяц ' + current_month + ';')
                else:
                    print('Прошлый сезон лиги ' + league[i].league_name + ', месяц ' + current_month + ';')

                # Получим тело таблицы "Календарь Игр & Результаты":
                tournament_fixture = ui.WebDriverWait(driver, 15).until(
                    lambda driver1: driver.find_element_by_css_selector('#tournament-fixture > tbody:nth-child(1)'))
                tournament_fixture_innerhtml = tournament_fixture.get_property('innerHTML')

                # Спарсим все дни с матчами в таблице:
                days = tournament_fixture_innerhtml.split('<tr class="rowgroupheader"><th colspan="7">')
                # Пройдёмся по всем дням в таблице в обратном порядке:
                for day in reversed(days):
                    soup = BeautifulSoup(day, 'html.parser')
                    # Если в дне найдена хотя бы одна красная карточка
                    if soup.find('span', {'class': 'rcard ls-e'}) is not None:
                        # Получим список гостевых и домашних команд в данном дне:
                        teams_home = soup.findAll('td', 'home')
                        teams_away = soup.findAll('td', 'away')
                        teams_all = teams_home + teams_away
                        # Пройдёмся по общему списку команд в данном дне (teams_all):
                        for j in range(0, len(teams_all)):
                            # Если у класса команды есть КК,
                            if teams_all[j].find('span', {'class': 'rcard ls-e'}) is not None:
                                # спарсим количество КК, которое эта команда получила в матче:
                                kk_count_in_match = int(teams_all[j].find('span', {'class': 'rcard ls-e'}).text)
                                # и проверим наличие этой команды с КК в массиве match[]:
                                for k in range(0, matches_length):
                                    # Пройдёмся по всем матчам массива match[] (но только для текущей лиги):
                                    if match[k].league_name == league[i].league_name:

                                        if not this_season:
                                            # Если прошлый сезон, то для каждой команды данной лиги массива match[]
                                            # проверим её присутствие в таблице tournament_fixture:
                                            if (match[k].team_home_found_in_last_season is False) and\
                                                    (match[k].team_home_name in tournament_fixture_innerhtml):
                                                # Если команда не была найдена ранее в прошлом сезоне
                                                # и при этом присутствует в таблице прошлого сезона,
                                                # отметим, что команда присутствует в прошлом сезоне данной лиги:
                                                match[k].team_home_found_in_last_season = True
                                            if (match[k].team_away_found_in_last_season is False) and\
                                                    (match[k].team_away_name in tournament_fixture_innerhtml):
                                                # Если команда не была найдена ранее в прошлом сезоне
                                                # и при этом присутствует в таблице прошлого сезона,
                                                # отметим, что команда присутствует в прошлом сезоне данной лиги:
                                                match[k].team_away_found_in_last_season = True

                                        if match[k].team_home_name in teams_all[j].text:
                                            # Если в массиве match[] присутствует название команды:
                                            if match[k].team_home_last_kk_date == '':
                                                # Если дата последней КК для этой команды не была найдена ранее,
                                                # спарсим её из rowgroupheader текущего дня:
                                                match[k].team_home_last_kk_date = \
                                                    datestring_format(day.split('</th>')[0])
                                            # Прибавим к счётчику КК этой команды kk_count_in_match:
                                            if this_season:
                                                match[k].team_home_kk_this_season_count += kk_count_in_match
                                            else:
                                                match[k].team_home_kk_last_season_count += kk_count_in_match

                                        if match[k].team_away_name in teams_all[j].text:
                                            # Если в массиве match[] присутствует название команды:
                                            if match[k].team_away_last_kk_date == '':
                                                # Если дата последней КК для этой команды не была найдена ранее,
                                                # спарсим её из rowgroupheader текущего дня:
                                                match[k].team_away_last_kk_date = \
                                                    datestring_format(day.split('</th>')[0])
                                            # Прибавим к счётчику КК этой команды kk_count_in_match:
                                            if this_season:
                                                match[k].team_away_kk_this_season_count += kk_count_in_match
                                            else:
                                                match[k].team_away_kk_last_season_count += kk_count_in_match

                # Проверим, активна ли кнопка "предыдущий месяц". Получим перечень её классов:
                classes_of_button = driver.find_element_by_css_selector('.previous').get_property('className')
                if 'is-disabled' in classes_of_button:
                    # Если кнопка неактивна, то продолжать листать календарь назад уже нельзя.
                    if this_season:
                        print('Парсинг текущего сезона завершён.')
                    else:
                        print('Парсинг прошлого сезона завершён.')
                    previous_clicked = False
                else:
                    # Если кнопка активна, перейдём на предыдущий месяц календаря:
                    driver.find_element_by_css_selector('.previous').click()
                    time.sleep(sleep_table_time)  # Пауза для прогрузки таблицы
    if this_season:
        print('Информация о КК за текущий сезон получена.')
    else:
        print('Информация о КК за прошлый сезон получена.')


def write_to_spreadsheets():
    ##################################################################################
    # ВЫВОД ДАННЫХ МАССИВА match[] В GOOGLE SHEETS
    ##################################################################################
    print('Запишем полученную информацию в Google Sheets...')
    # Создаём Service-объект, для работы с Google-таблицами:
    credentials_file = 'RedCardsProject-90325d995892.json'  # имя выгруженного файла с закрытым ключом
    # В Scope укажем к каким API мы хотим получить доступ:
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    # Заполним массив с учётными данными:
    credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
    http_auth = credentials.authorize(httplib2.Http())
    # Создаём Service-объект для работы с Google-таблицами:
    service = apiclient.discovery.build('sheets', 'v4', http=http_auth)
    # Укажем идентификатор документа, к которому хотим получить доступ.
    spreadsheet_id = '10PPb2Tk51-68fBqew-tEnDqbA_MaCEQGyrHAxcnQ4Jc'
    ranges = []
    include_grid_data = False
    # Получим spreadsheet с указанным spreadsheet_id:
    request = service.spreadsheets().get(spreadsheetId=spreadsheet_id, ranges=ranges, includeGridData=include_grid_data)
    spreadsheet = request.execute()
    title = datestring_format(required_date)
    # Попробуем найти в spreadsheet лист с названием title и если он есть, получим его sheetId:
    sheet_exists = False
    sheet_id = 0

    for sheet in spreadsheet['sheets']:
        if sheet['properties']['title'] == title:
            sheet_exists = True
            sheet_id = sheet['properties']['sheetId']
            break

    if not sheet_exists:
        # Если лист не существует, то создадим его:
        result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'], body={
            'requests': [
                {
                    "addSheet": {
                        "properties": {
                            "title": title,
                            'index': 0,  # Порядковый номер листа в списке листов. Если 0 - то самый левый
                            "gridProperties": {
                                "rowCount": 1000,
                                "columnCount": 15
                            }
                        }
                    }

                }
            ]
        }).execute()
        # Получим sheet_id только что созданного листа:
        sheet_id = result['replies'][0]['addSheet']['properties']['sheetId']
    '''
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'], body={
        'requests': [
            {
                # Обновить свойства листа:
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': sheet_id,
                        'title': title,
                        # 'index': number,
                        'gridProperties': {
                            'columnCount': 15
                        },
                    },
                    'fields': 'title, gridProperties(columnCount)'
                }
            }
        ]
    }).execute()
    '''

    # Определим объект границы (тип данных словарь), чтобы применять его в дальнейшем при рисовании границ:
    border = {
        'style': 'SOLID', 'width': 3,
        'color':
            {
                'red': 0, 'green': 0, 'blue': 0, 'alpha': 1
            }
    }

    # Прочитаем первые две строки листа:
    row_index = 1
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet['spreadsheetId'], range=title+'!A' + str(row_index) + ':O' + str(row_index + 1))\
        .execute()
    two_rows = result.get('values')

    if two_rows is None:
        # Если первые две строки листа пустые, то сделаем заголовок:
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'], body={
            'valueInputOption': 'USER_ENTERED',
            'data': [
                {
                    'range': title+'!A1:O2',
                    'majorDimension': 'ROWS',
                    # сначала заполнять ряды, затем столбцы (т.е. самые внутренние списки в values - это ряды)
                    'values':
                    [
                        ['Лига', 'Дома', 'Гости', 'Дата', 'Судья', '', '', '', 'Команды', '', '', '', '', '', 'Судья'],
                        ['', '', '', '', 'Этот сезон', 'Все сезоны', 'Командам', 'Посл. 20 игр', 'КК этот сезон',
                         'КК прош. сезон', 'Дата последней 1/2', '', 'Личные встречи', '',
                         'Имя на Championat (имя на Whoscored)']
                    ]
                }
            ]
        }).execute()
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'], body={
            'requests': [
                # Нарисуем границы для всех ячеек:
                {
                    'updateBorders': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': 0,
                            'endRowIndex': 2,
                            'startColumnIndex': 0,
                            'endColumnIndex': 15
                        },
                        'top': border,
                        'bottom': border,
                        'left': border,
                        'right': border,
                        'innerHorizontal': border,
                        'innerVertical': border
                    }
                },
                # Удалим границы для ячеек столбца N:
                {
                    'updateBorders': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': 0,
                            'endRowIndex': 2,
                            'startColumnIndex': 13,
                            'endColumnIndex': 14
                        },
                        'top': {'style': 'NONE'},
                        'bottom': {'style': 'NONE'},
                        'innerHorizontal': {'style': 'NONE'}
                    }
                },
                # Выполним слияние необходимых ячеек в заголовке таблицы:
                {
                    'mergeCells': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': 0,
                            'endRowIndex': 1,
                            'startColumnIndex': 4,
                            'endColumnIndex': 8
                        },
                        'mergeType': 'MERGE_ALL'
                    }
                },
                {
                    'mergeCells': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': 0,
                            'endRowIndex': 1,
                            'startColumnIndex': 8,
                            'endColumnIndex': 13
                        },
                        'mergeType': 'MERGE_ALL'
                    }
                },
                {
                    'mergeCells': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': 1,
                            'endRowIndex': 2,
                            'startColumnIndex': 0,
                            'endColumnIndex': 4
                        },
                        'mergeType': 'MERGE_ALL'
                    }
                },
                {
                    'mergeCells': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': 1,
                            'endRowIndex': 2,
                            'startColumnIndex': 10,
                            'endColumnIndex': 12
                        },
                        'mergeType': 'MERGE_ALL'
                    }
                },
                # Заморозим первые две строки заголовка:
                {
                    'updateSheetProperties': {
                        'properties': {
                            'sheetId': sheet_id,
                            'gridProperties': {
                                'frozenRowCount': 2
                            }
                        },
                        'fields': 'gridProperties.frozenRowCount'
                    }
                }
            ]
        }).execute()
        row_index = 3

    else:
        # Если первые две строки листа не пустые, то предполагаем, что заготовок уже существует.
        # Найдём индекс первых двух пустых строк на листе, чтобы в дальнейшем начать с них запись:
        while two_rows is not None:
            row_index += 1
            result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet['spreadsheetId'],
                range=title + '!A' + str(row_index) + ':O' + str(row_index + 1)) \
                .execute()
            two_rows = result.get('values')

    n = 0
    if row_index == 3:
        # Если есть только заголовок, а данных нет,
        # То начинаем вывод данных о матчах с третьей по счёту строчки таблицы:
        n = row_index

    if row_index > 3:
        # Если какие-то данные о матчах уже присутствуют,
        # сделаем отступ в одну строку от последнего матча:
        n = row_index + 1
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'], body={
            'requests': [
                # Нарисуем границу для всех ячеек:
                {
                    'updateBorders': {
                        'range': {
                            'sheetId': sheet_id,
                            'startRowIndex': row_index,
                            'endRowIndex': row_index+1,
                            'startColumnIndex': 0,
                            'endColumnIndex': 15
                        },
                        'top': border
                    }
                }
            ]
        }).execute()

    # Выведем все матчи в цикле:
    previous_league_name = ''
    for i in range(0, matches_length):
        # Соберём строку str_personal_meetings для её последующего вывода:
        if ((match[i].team_home_personal_meetings_kk_count_home != -1) and
                (match[i].team_home_personal_meetings_kk_count_away != -1) and
                (match[i].team_away_personal_meetings_kk_count_home != -1) and
                (match[i].team_away_personal_meetings_kk_count_away != -1)):

                str_personal_meetings = str(match[i].team_home_personal_meetings_kk_count_home) + 'д' +\
                                       str(match[i].team_home_personal_meetings_kk_count_away) + 'г/' +\
                                       str(match[i].team_away_personal_meetings_kk_count_home) + 'д' +\
                                       str(match[i].team_away_personal_meetings_kk_count_away) + 'г из ' +\
                                       str(match[i].personal_meetings_count) + ' (' +\
                                       match[i].teams_personal_meetings_last_kk_date + ')'

        else:
                str_personal_meetings = "Не встречались"

        # Подготовим строки для их последующего вывода:
        if match[i].team_home_found_in_last_season is True:
            if match[i].team_home_kk_this_season_count == 0 and match[i].team_home_kk_last_season_count == 0:
                match[i].team_home_last_kk_date = 'Давно'
            str_team_home_kk_last_season_count = str(match[i].team_home_kk_last_season_count)
        else:
            str_team_home_kk_last_season_count = '???'
            if match[i].team_home_kk_this_season_count == 0:
                match[i].team_home_last_kk_date = '???'

        if match[i].team_away_found_in_last_season is True:
            if match[i].team_away_kk_this_season_count == 0 and match[i].team_away_kk_last_season_count == 0:
                match[i].team_away_last_kk_date = 'Давно'
            str_team_away_kk_last_season_count = str(match[i].team_away_kk_last_season_count)
        else:
            str_team_away_kk_last_season_count = '???'
            if match[i].team_away_kk_this_season_count == 0:
                match[i].team_away_last_kk_date = '???'

        # Если у нового матча лига сменилась, разделим лиги при выводе пустой строчкой:
        if match[i].league_name != previous_league_name and previous_league_name != '':
            n += 1

        previous_league_name = match[i].league_name

        # Выведем строку матча:
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'], body={
            'valueInputOption': 'USER_ENTERED',
            'data': [
                {
                    'range': title+'!A'+str(n)+':O'+str(n)+'',
                    'majorDimension': 'ROWS',
                    # сначала заполнять ряды, затем столбцы (т.е. самые внутренние списки в values - это ряды)
                    'values':
                        [
                            # ['Лига', 'Дома', 'Гости', 'Дата', 'Судья', '', '', '', 'Команды', '', '', '', '', '', 'Судья']
                            [match[i].league_name, match[i].team_home_name, match[i].team_away_name,
                             match[i].match_datetime, '', '', '', '',
                             str(match[i].team_home_kk_this_season_count) + '/' +
                             str(match[i].team_away_kk_this_season_count),
                             str_team_home_kk_last_season_count + '/' +
                             str_team_away_kk_last_season_count,
                             match[i].team_home_last_kk_date, match[i].team_away_last_kk_date,
                             str_personal_meetings,
                             '', '']
                        ]
                }
            ]
        }).execute()
        n += 1

    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'], body={
        'requests': [
            # Задать размер по содержимому для всех столбцов
            {
                'autoResizeDimensions': {
                    'dimensions': {
                        'sheetId': sheet_id,
                        'dimension': 'COLUMNS',  # COLUMNS - потому что столбец
                        'startIndex': 0,  # Столбцы нумеруются с нуля
                        'endIndex': 15  # startIndex берётся включительно, endIndex - НЕ включительно
                    }
                }
            },
            # Применим форматирование (цвет фона и текста, выравнивание) к ячейкам:
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 515
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 1.0,
                                "green": 1.0,
                                "blue": 1.0
                            },
                            "horizontalAlignment": "CENTER",
                            "textFormat": {
                                "foregroundColor": {
                                    "red": 0.0,
                                    "green": 0.0,
                                    "blue": 0.0
                                },
                                # "fontSize": 12,
                                # "bold": True
                            }
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
                }
            }
        ]
    }).execute()

    '''
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'], body={
        'requests': [
            # Задать размеры для всех столбцов
            {
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': sheet_id,
                        'dimension': 'COLUMNS',
                        'startIndex': 0,
                        'endIndex': 15
                    },
                    'properties': {
                        'pixelSize': 160
                    },
                    'fields': 'pixelSize'
                }
            }
        ]
    }).execute()
    '''
    print('Полученная информация записана в Google Sheets.')


def main():
    # Подключимся к БД Microsoft Access через экземпляр ODBC
    db = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\\RC_base2_lessdata.mdb')
    dbc = db.cursor()
    # Получим информацию из БД, занесём её в rows
    rows = dbc.execute('select * from [Leagues]').fetchall()
    db.close()
    # Определим длину списка rows:
    global league_length
    league_length = len(rows)
    # Инициализируем список league[] и заполним его экземплярами класса League:
    global league
    league.clear()
    for i in range(0, league_length):
        league.append(League(rows[i][1], rows[i][2], rows[i][3]))

    includes_path = 'C:/Users/BRAXXMAN/PycharmProjects/includes/'
    global driver
    # os.environ['webdriver.chrome.driver'] = includes_path+'chromedriver.exe'
    # driver = webdriver.Chrome(executable_path=includes_path+'chromedriver.exe')
    # Отключим изображения в браузере Firefox:
    options = webdriver.FirefoxOptions()
    options.set_preference('permissions.default.image', 2)
    options.set_preference('dom.ipc.plugins.enabled.libflashplayer.so', 'false')
    driver = webdriver.Firefox(options=options, executable_path=includes_path + 'geckodriver.exe')
    # Получим объект calendarWidget с окна программы wind:
    cal = wind.calendarWidget
    # Получим выбранную дату и выполним преобразование из QDate в datetime:
    global required_date
    required_date = cal.selectedDate().toPyDate()
    # Преобразуем datetime в unix:
    global required_date_unix
    required_date_unix = time.mktime(required_date.timetuple())
    # Преобразуем datetime в string:
    required_date = required_date.strftime('%A, %b %e %Y')
    # Удалим двойной пробел, который образуется, если %e (день месяца) меньше 10:
    required_date = required_date.replace('  ', ' ')

    # Получим информацию о матчах в указанный день:
    match.clear()
    get_matches()

    if matches_length > 0:
        # Получим информацию о личных встречах команд:
        #get_personal_meetengs()
        # Достанем ссылки на календарь игр прошлых сезонов:
        #get_url_games_calendar_past_season()
        # Получим информацию о количестве КК у команд за этот сезон, о дате последней КК:
        #get_kk_this_or_last_season(True)
        # Получим информацию о количестве КК у команд за прошлый сезон, о дате последней КК
        # (если она не была найдена в последнем сезоне):
        #get_kk_this_or_last_season(False)
        # Запишем полученную информацию в Google Sheets:
        write_to_spreadsheets()
    else:
        print('Ни в одной из лиг не найдено матчей на ' + required_date + '!')
    # Завершим сессию браузера и закроем его окно:
    driver.quit()


if __name__ == '__main__':
    league = []
    league_length = 0

    match = []
    matches_length = 0

    sleep_page_time = 5
    sleep_table_time = 1
    # Получаем текущую локаль:
    default_loc = locale.getlocale()
    # Изменяем локаль для корректной конвертации строки с русской датой в datetime:
    locale.setlocale(locale.LC_ALL, ('RU', 'UTF8'))

    app = QApplication(sys.argv)
    wind = Window()
    app.exec_()

    # Меняем локаль назад:
    locale.setlocale(locale.LC_ALL, default_loc)
