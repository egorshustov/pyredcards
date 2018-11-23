import pyodbc
import httplib2
import requests
import json
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
import os
import selenium.webdriver.support.ui as ui
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from lxml import html
import time
import datetime
import locale
import sys
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi


class Window(QMainWindow):

    def __init__(self):
        super(Window, self).__init__(flags=Qt.WindowFlags())
        # Загрузим UI из файла:
        loadUi('redcardsdesigner.ui', self)

        # Получим виджет календаря:
        self.cal = self.calendarWidget
        # Установим в нём выбранную дату по умолчанию как завтрашнюю (отстоящую на 1 день):
        self.cal.setSelectedDate(self.cal.selectedDate().addDays(1))

        # Подключимся к БД Microsoft Access через экземпляр ODBC
        db = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\\RC_base2.mdb')
        dbc = db.cursor()
        # Получим информацию из БД, занесём её в rows
        rows = dbc.execute('select * from [Leagues]').fetchall()
        db.close()
        # Определим длину списка rows:
        global mdb_league_length
        mdb_league_length = len(rows)

        # Получим виджет списка для лиг:
        leagues_list = self.listViewLeagues
        # Сделаем так, чтобы выведенные в список строки нельзя было редактировать в GUI:
        leagues_list.setEditTriggers(QAbstractItemView.NoEditTriggers)
        # Создадим объект модели для элементов
        # (если указываем 'self.', то объект из локального становится объектом класса):
        self.model = QStandardItemModel()
        # Заполним список mdb_league[] экземплярами класса League:
        global mdb_league
        for i in range(0, mdb_league_length):
            mdb_league.append(League(rows[i][1], rows[i][2], rows[i][3]))
            # Создадим элемент для каждой лиги:
            item = QStandardItem(mdb_league[i].league_name)
            # Элемент может принимать два значения - True или False:
            item.setCheckState(2)
            # Добавим встроенный CheckBox для каждого элемента:
            item.setCheckable(True)
            # Применим элемент к модели:
            self.model.appendRow(item)
        # Применим модель к списку:
        leagues_list.setModel(self.model)

        # Получим виджет списка для записи в лог:
        self.log_list = self.listViewLog
        # Создадим объект модели для элементов лога:
        self.model_log = QStandardItemModel()
        # Применим модель к списку:
        self.log_list.setModel(self.model_log)

        # Создадим обработчик для кнопки:
        self.startButton.clicked.connect(self.on_startbutton_clicked)
        # Создадим обработчик для пункта меню action_checkbox:
        self.action_checkbox.triggered.connect(self.on_invert_checkbox_clicked)
        # Инициализируем объект класса нити:
        self.workerThread = WorkerThread()
        # Запустим форму окна:
        self.show()

    def on_startbutton_clicked(self):
        # Выключим кнопку startButton:
        self.startButton.setEnabled(False)
        # При нажатии на кнопку запустим нить workerThread:
        self.workerThread.start()
        # cal = self.calendarWidget

    def log(self, datestring):
        # Выводим данные в командную строку и в GUI:
        print(datestring)
        self.model_log.appendRow(QStandardItem(datestring))

    def on_invert_checkbox_clicked(self):
        # Пройдёмся по всем элементам лиг в leagues_list:
        for i in range(0, mdb_league_length):
            # Если чекбокс элемента лиги включен:
            if self.model.item(i).checkState() == 2:
                # Выключим его:
                self.model.item(i).setCheckState(0)
            # Если чекбокс элемента лиги выключен:
            else:
                # Выключим его:
                self.model.item(i).setCheckState(2)


class WorkerThread(QThread):

    def __init__(self):
        super(WorkerThread, self).__init__()

    def run(self):
        # Вызываем главную функцию:
        main()
        self.exec_()


class League:

    def __init__(self, league_name, url_whoscored, url_championat, ):
        self.league_name = league_name
        self.url_whoscored = url_whoscored
        self.url_championat = url_championat
        self.url_past_season = ''
        self.url_referee_statistics = ''
        self.url_games_calendar_past_season = ''
        self.matches_found = True
        self.referees_found = False


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
        self.referee_team_home = False  # Судья, Командам
        self.referee_to_team_home_average = ''
        self.referee_team_home_matches_count = -1
        self.referee_team_away = False
        self.referee_to_team_away_average = ''
        self.referee_team_away_matches_count = -1
        self.referee_last_twenty_home_count = 0  # Судья, Посл. 20 игр
        self.referee_last_twenty_away_count = 0
        self.referee_last_twenty_last_kk_date = '-'
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
        self.championat_teamsstring = ''


class MatchChampionat:

    def __init__(self):
        self.match_url = ''  # ссылка на матч
        self.league_name = ''  # Лига
        self.team_home_name = ''  # Дома
        self.team_away_name = ''  # Гости
        self.referee_name = ''  # Судья, Имя на championat
        self.teamsstring = ''
        self.score = 0


class Referee:

    def __init__(self):
        self.referee_name_whoscored = ''  # Судья, Имя на whoscored
        self.referee_url = ''  # Судья, URL на whoscored
        self.score = 0


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
    wind.log('Найдём матчи для каждой из лиг на ' + datestring_format(required_date) + ':')
    global match
    i_match = 0
    # Для каждой лиги переходим на страницу Календаря Игр сайта whoscored:
    for i in range(0, league_length):
        next_clicked = True
        wind.log('Лига ' + league[i].league_name + '...')
        driver.get(league[i].url_whoscored)
        time.sleep(sleep_page_time)  # Пауза для прогрузки страницы

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
            # wind.log(tournament_fixture_innerhtml)
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
                    wind.log('В ' + league[i].league_name + ' на ru.whoscored.com нет матчей в указанный день!')
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
                        wind.log('Сезон лиги ' + league[i].league_name + ' уже завершился на указанный день!')
                    else:
                        # Если кнопка активна, перейдём на предыдущий месяц календаря:
                        driver.find_element_by_css_selector('.next').click()
                        time.sleep(sleep_table_time)  # Пауза для прогрузки таблицы
    # Определим длину списка matches:
    global matches_length
    matches_length = len(match)
    wind.log('Поиск матчей для каждой из лиг в указанный день завершен.')


def get_personal_meetengs():
    ##################################################################################
    # ЛИЧНЫЕ ВСТРЕЧИ
    ##################################################################################
    wind.log('Спарсим информацию личных встреч команд для каждого найденного матча:')
    for i in range(0, matches_length):
        driver.get(match[i].match_url)
        wind.log('Команды ' + match[i].team_home_name + ' и ' + match[i].team_away_name)
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
            wind.log('У команд ' + match[i].team_home_name + ' и ' +
                     match[i].team_away_name + ' не было совместных встреч!')
    wind.log('Информация личных встреч команд для каждого найденного матча получена.')


def get_url_games_calendar_past_season():
    wind.log('Достанем ссылки на календарь игр прошлых сезонов:')
    for i in range(0, league_length):
        if league[i].matches_found is True:
            driver.get(league[i].url_past_season)
            time.sleep(sleep_page_time)
            league[i].url_games_calendar_past_season = \
                driver.find_element_by_css_selector('#sub-navigation > ul:nth-child(1) >'
                                                    ' li:nth-child(2) > a:nth-child(1)').get_property('href')
    wind.log('Ссылки успешно получены.')


def get_kk_this_or_last_season(this_season):
    ##################################################################################
    # КК ЗА ЭТОТ ИЛИ ПРЕДЫДУЩИЙ СЕЗОН, ДАТА ПОСЛЕДНЕЙ КК
    ##################################################################################
    if this_season:
        wind.log('Получим информацию о КК за текущий сезон:')
    else:
        wind.log('Получим информацию о КК за прошлый сезон:')

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
                    wind.log('Текущий сезон лиги ' + league[i].league_name + ', месяц ' + current_month + ';')
                else:
                    wind.log('Прошлый сезон лиги ' + league[i].league_name + ', месяц ' + current_month + ';')

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
                        wind.log('Парсинг текущего сезона завершён.')
                    else:
                        wind.log('Парсинг прошлого сезона завершён.')
                    previous_clicked = False
                else:
                    # Если кнопка активна, перейдём на предыдущий месяц календаря:
                    driver.find_element_by_css_selector('.previous').click()
                    time.sleep(sleep_table_time)  # Пауза для прогрузки таблицы
    if this_season:
        wind.log('Информация о КК за текущий сезон получена.')
    else:
        wind.log('Информация о КК за прошлый сезон получена.')


def get_referee_championat():
    ##################################################################################
    # ПОЛУЧИМ ИМЯ СУДЬИ НА CHAMPIONAT И СОПОСТАВИМ МАТЧИ НА ОБОИХ САЙТАХ
    ##################################################################################
    wind.log('Начинаем парсинг championat.com (получим имя судьи для матчей):')
    date = datestring_format(required_date)
    date = date.replace('-', '.')
    h = httplib2.Http()  # disable_ssl_certificate_validation=True
    # Инициализируем список элементов класса MatchChampionat:
    match_championat = []
    i_match_champ = 0
    # Для каждой лиги:
    for i in range(0, league_length):
        # Но только если в лиге найдены матчи:
        if league[i].matches_found:
            # Получим страницу календаря игр для каждой лиги:
            tournament_fixture_championat = h.request(league[i].url_championat, 'GET')
            # Разобъём с помощью указанной даты страницу календаря игр на строки:
            all_matches = tournament_fixture_championat[1].decode('utf-8').split(date)
            # Пройдёмся с первого по последний элемент в списке полученных строк
            # (каждая строка содержит информацию об отдельном матче):
            for j in range(1, len(all_matches)):
                match_championat.append(MatchChampionat())
                # Занесём в массив название текущей лиги:
                match_championat[i_match_champ].league_name = league[i].league_name
                soup = BeautifulSoup(all_matches[j], 'html.parser')
                # Получим все теги 'a' в строке:
                all_a_tags = soup.findAll('a')
                # Первый тег содержит имя домашней команды:
                match_championat[i_match_champ].team_home_name = all_a_tags[0].text
                # Второй тег содержит имя гостевой команды:
                match_championat[i_match_champ].team_away_name = all_a_tags[1].text
                # Третий тег содержит ссылку на будущий матч двух команд:
                match_championat[i_match_champ].match_url = 'https://www.championat.com' + all_a_tags[2]['href']

                # Выполним новый GET-запрос чтобы получить имя судьи (для этого перейдём по match_url):
                match_page = h.request(match_championat[i_match_champ].match_url, 'GET')
                # Страницв закодирована с utf-8, раскодируем её:
                match_page_st = match_page[1].decode('utf-8')
                if 'Главный судья:' in match_page_st:
                    # Если строка 'Главный судья: ' присутствует на странице, то судья известен. Спарсим его имя:
                    tree = html.fromstring(match_page_st)
                    # Получим элемент с именем судьи по его XPath с помощью библиотеки lxml:
                    referee_element = tree.xpath('/html/body/div[5]/div[6]/div[1]/div/div/div[4]/div[2]/div[1]/a')[0]
                    match_championat[i_match_champ].referee_name = referee_element.text
                    league[i].referees_found = True
                else:
                    # Если строка 'Главный судья: ' не присутствует на странице, то судья пока известен.
                    # Занесём '???' вместо его имени:
                    match_championat[i_match_champ].referee_name = '???'
                    wind.log('Для матча ' + match_championat[i_match_champ].team_home_name + '-' +
                             match_championat[i_match_champ].team_away_name + ' лиги ' +
                             match_championat[i_match_champ].league_name + ' судья ещё не известен!')
                i_match_champ += 1

    wind.log('Парсинг championat.com завершён.')

    wind.log('Сопоставим матчи на whoscored и championat (проверь соответствие строк!):')
    # Cольём названия команд в одну строку и удалим пробелы, чтобы выполнить побуквенное сравнение:
    for i in range(0, matches_length):
        match[i].teamsstring = (match[i].team_home_name + match[i].team_away_name).replace(' ', '')
    for i in range(0, len(match_championat)):
        match_championat[i].teamsstring = (match_championat[i].team_home_name + match_championat[i].team_away_name).\
            replace(' ', '')

    # Выполним побуквенное сравнение массивов match[].teamsstring и match_championat[].teamsstring
    # и найдём соответствие для матчей, спарсенных с whoscored, матчам, спарсенным с championat,
    # для того, чтобы заполнить массив match[].referee_name_championat
    # (найти имя судьи на championat для каждого найденного матча на whoscored):

    # Для каждой лиги:
    for i in range(0, league_length):
        # Но только если в лиге найдены матчи:
        if league[i].matches_found:
            # Пройдёмся по всем матчам, найденным на сайте whoscored:
            for j in range(0, matches_length):
                # которые принадлежат i-той лиге:
                if match[j].league_name == league[i].league_name:
                    # Возьмём строку с названиями двух команд (на сайте whoscored) и в цикле пройдемся по ней,
                    # перебирая каждые её два символа, идущие подряд, с шагом, равным 1
                    # (пройдёмся по строке "стяжками"):
                    incr = 0
                    while len(match[j].teamsstring[incr:incr+2]) == 2:
                        # Если длина стяжка по прежнему равна 2 (строка не заканчивается),
                        # то получим этот двухсимвольный стяжок и запишем его в переменную two_symbols:
                        two_symbols = match[j].teamsstring[incr:incr+2]
                        # Найдём вхождения этого стяжка в каждую из строк массива match_championat[].teamsstring:
                        for n in range(0, len(match_championat)):
                            if match_championat[n].league_name == league[i].league_name:
                                if two_symbols in match_championat[n].teamsstring:
                                    # Если вхождение найдено,
                                    # добавим "очко" этому матчу, найденному на championat:
                                    match_championat[n].score += 1
                        incr += 1

                    # Выполнили перебор и подсчёт количества совпадений
                    # (очков match_championat[].score) для каждой строки массива match_championat[].teamsstring,
                    # теперь найдём максимальное значение match_championat[].score в массиве match_championat[]
                    # и запомним индекс этого элемента массива
                    # (именно этот элемент и будет содержать информацию о судье для матча whoscored):
                    max_score = 0
                    max_score_index = 0
                    for n in range(0, len(match_championat)):
                        if match_championat[n].league_name == league[i].league_name:
                            if max_score < match_championat[n].score:
                                max_score = match_championat[n].score
                                max_score_index = n

                    # Выведем строки с названиями команд на whoscored и championat для проверки пользователем:
                    whoscored_matchstring = match[j].teamsstring
                    founded_championat_matchstring = match_championat[max_score_index].teamsstring
                    wind.log(whoscored_matchstring + ' = ' + founded_championat_matchstring
                             + ' (' + str(max_score) + ' совпадений);')

                    match[j].championat_teamsstring = match_championat[max_score_index].teamsstring
                    match[j].referee_name_championat = match_championat[max_score_index].referee_name
                    if match[j].referee_name_championat != '???':
                        translate_url = 'https://translate.yandex.net/api/v1.5/tr.json/translate?key=' \
                                        'trnsl.1.1.20180922T150311Z.5ccad8013c0e69ed.3d4026b2fe47ae4dd0' \
                                        'cc09e8e9017f678fcbe3d9&text=' + match[j].referee_name_championat +\
                                        '&lang=ru-en'
                        translated_responce = requests.get(translate_url)
                        # Сконвертируем тип bytes в словарь dict:
                        translated_dict = json.loads(translated_responce.content)
                        match[j].referee_name_championat_translated_to_en = translated_dict['text'][0]

                    # Обнулим все match_championat[].score, перед тем, как перейти к следующему матчу match[j]:
                    for n in range(0, len(match_championat)):
                        match_championat[n].score = 0

    wind.log('Сопоставление завершено.')


def get_referee_whoscored():
    ##################################################################################
    # ПОЛУЧИМ ИМЕНА СУДЕЙ И ИХ URL НА WHOSCORED
    ##################################################################################
    # Подключимся к БД Microsoft Access через экземпляр ODBC
    db = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=.\\RC_base2.mdb')
    dbc = db.cursor()
    # Получим информацию о судьях, для которых чётко определено соответствие имён на championat и whoscored:
    rows = dbc.execute('select * from [Referees]').fetchall()
    db.close()
    referee_line = []
    # Считаем информацию о заданных судьях в двумерный список:
    referee_defined = []
    for i in range(0, len(rows)):
        referee_line.clear()
        referee_line.append(rows[i][1])
        referee_line.append(rows[i][2])
        referee_defined.append(referee_line)
        referee_line = []

    # Объявим список, в который будут загружены все судьи в таблице,
    # для того, чтобы в дальнейшем выбрать наиболее подходящего:
    referee = []
    # Для каждой лиги:
    for i in range(0, league_length):
        # Но только если в лиге найдены судьи (на сайте championat):
        if league[i].referees_found:
            wind.log('Получим данные по судьям лиги ' + league[i].league_name + '...')
            i_referee = 0
            driver.get(league[i].url_referee_statistics)
            time.sleep(sleep_page_time)
            end_tag_found = False
            # Пока не найдём последний тег:
            while not end_tag_found:
                # Получим тело таблицы "Статистика Рефери":
                try:
                    referee_tournaments = ui.WebDriverWait(driver, 15).until(
                        lambda driver1: driver.find_element_by_id('referee-tournaments-table-body'))
                except Exception:
                    driver.get(league[i].url_referee_statistics)
                    time.sleep(sleep_page_time)
                    referee_tournaments = ui.WebDriverWait(driver, 15).until(
                        lambda driver1: driver.find_element_by_id('referee-tournaments-table-body'))
                    pass

                referee_tournaments_innerhtml = referee_tournaments.get_property('innerHTML')
                soup = BeautifulSoup(referee_tournaments_innerhtml, 'html.parser')
                # Получим теги всех судей на текущей странице таблицы:
                referees_tags = soup.findAll('tr')
                for referee_tag in referees_tags:
                    # Пройдёмся по каждому полученному тегу и спарсим из него ссылку на судью и имя судьи на whoscored.
                    # Попробуем получить тег 'a' с классом tournament-link, содержащий информацию о судье:
                    tournament_link_tag = referee_tag.find('a', {'class': 'tournament-link'})
                    if tournament_link_tag is not None:
                        # Если тег 'a' с классом tournament-link имеется в теге referee_tag,
                        # достанем из него информацию:
                        referee.append(Referee())
                        referee[i_referee].referee_url = 'https://ru.whoscored.com' +\
                                                         referee_tag.find('a', {'class': 'tournament-link'})['href']
                        referee[i_referee].referee_name_whoscored =\
                            referee_tag.find('a', {'class': 'tournament-link'}).text
                        i_referee += 1
                    else:
                        # Значит, referee_tag - это итоговый тег таблицы, содержащий строку 'Сумма/Среднее количество':
                        end_tag_found = True
                # Если итоговый тег таблицы не был найден, значит таблицу можно листать дальше:
                if not end_tag_found:
                    # time.sleep(sleep_table_time)
                    driver.find_element_by_css_selector('#next').click()
                    time.sleep(sleep_table_time)  # Пауза для прогрузки таблицы

            # Найдём самого созвучного судью для каждого матча, где referee_name_championat != "???"
            # и занесём его имя и url в массив match[]:
            wind.log('Сопоставим имена судей на championat и whoscored (проверь соответствие строк!):')
            # Пройдёмся по всем матчам, найденным на сайте whoscored:
            for j in range(0, matches_length):
                # которые принадлежат i-той лиге и по которым известны судьи:
                if match[j].league_name == league[i].league_name and match[j].referee_name_championat != '???':
                    # Возьмём строку с переведённым на английский именем судьи
                    # (match[j].referee_name_championat_translated_to_en) и в цикле пройдемся по ней,
                    # перебирая каждые её два символа, идущие подряд, с шагом, равным 1
                    # (пройдёмся по строке "стяжками"):
                    incr = 0
                    while len(match[j].referee_name_championat_translated_to_en[incr:incr+2]) == 2:
                        # Если длина стяжка по прежнему равна 2 (строка не заканчивается),
                        # то получим этот двухсимвольный стяжок и запишем его в переменную two_symbols:
                        two_symbols = match[j].referee_name_championat_translated_to_en[incr:incr+2]
                        # Найдём вхождения этого стяжка в каждую из строк массива referee[].referee_name_whoscored:
                        for n in range(0, len(referee)):
                            if two_symbols in referee[n].referee_name_whoscored:
                                # Если вхождение найдено,
                                # добавим 'очко' этому судье:
                                referee[n].score += 1
                        incr += 1

                    # Выполнили перебор и подсчёт количества совпадений
                    # (очков referee[].score) для каждого судьи массива referee[],
                    # теперь найдём максимальное значение referee[].score в массиве referee[]
                    # и запомним индекс этого элемента массива
                    # (именно этот элемент и будет содержать информацию об имени судьи и его Url на whoscored):
                    max_score = 0
                    max_score_index = 0
                    for n in range(0, len(referee)):
                        if max_score < referee[n].score:
                            max_score = referee[n].score
                            max_score_index = n

                    # Выведем строки с именами судей на championat и whoscored для проверки пользователем:
                    match[j].referee_name_whoscored = referee[max_score_index].referee_name_whoscored
                    match[j].referee_url = referee[max_score_index].referee_url

                    # Пройдёмся по всему списку задефайненных судей:
                    for k in range(0, len(referee_defined)):
                        # Проверим, если данный судья уже задефайнен в БД:
                        if match[j].referee_name_championat == referee_defined[k][0]:
                            # Пройдёмся по всему списку судей текущей лиги:
                            for n in range(0, len(referee)):
                                # Если судья с задефайненым именем был найден в текущей лиге на whoscored:
                                if referee[n].referee_name_whoscored == referee_defined[k][1]:
                                    # Присвоим ему задефайненное имя и найденный ранее URL:
                                    match[j].referee_name_whoscored = referee_defined[k][1]
                                    match[j].referee_url = referee[n].referee_url

                    wind.log(match[j].referee_name_championat + ' = ' + match[j].referee_name_whoscored
                             + ' (' + str(max_score) + ' совпадений);')

                    # Обнулим все referee[].score, перед тем, как перейти к следующему матчу match[j]:
                    for n in range(0, len(referee)):
                        referee[n].score = 0

            wind.log('Получили данные по судьям лиги ' + league[i].league_name + '.')
            referee.clear()


def get_referee_info():
    ##################################################################################
    # ПОЛУЧИМ ИНФОРМАЦИЮ ПО СУДЬЕ С WHOSCORED
    ##################################################################################
    for i in range(0, matches_length):
        if match[i].referee_name_championat != '???':
            wind.log('Получим информацию по судье ' + match[i].referee_name_whoscored + '...')
            driver.get(match[i].referee_url)
            time.sleep(1.5*sleep_page_time)

            try:
                # Получим таблицу 'Турниры':
                referee_tournaments = ui.WebDriverWait(driver, 15).until(
                    lambda driver1: driver.find_element_by_id('referee-tournaments-table-body'))
            except Exception:
                driver.get(match[i].referee_url)
                time.sleep(1.5 * sleep_page_time)
                referee_tournaments = ui.WebDriverWait(driver, 15).until(
                    lambda driver1: driver.find_element_by_id('referee-tournaments-table-body'))
                pass

            referee_tournaments_innerhtml = referee_tournaments.get_property('innerHTML')
            # Достанем из неё среднее количество КК за текущий сезон:
            soup = BeautifulSoup(referee_tournaments_innerhtml, 'html.parser')
            # Получим теги всех лиг, которые судил текущий судья:
            leagues_tags = soup.findAll('tr')

            for league_tag in leagues_tags:
                # Пройдёмся по каждому полученному тегу и определим, в каком из них находится интересующая нас лига:
                if match[i].league_name in league_tag.text:
                    # Нужная лига найдена! Получим информацию о поведении судьи в этой лиге:
                    td_tags = league_tag.findAll('td')
                    match[i].referee_this_season_matches_count = int(td_tags[2].text)
                    match[i].referee_this_season_average = td_tags[8].text
                    if match[i].referee_this_season_average == '0.00':
                        match[i].referee_this_season_average = '0'
                    break

            # Нажмём на кнопки 'Все' для таблиц 'Турниры' и 'Команды':
            driver.find_element_by_id('alltime-referee-stats').click()
            driver.find_element_by_css_selector('#referee-team-filter-summary > div:nth-child(2) > div:nth-child(2)'
                                                ' > dl:nth-child(1) > dd:nth-child(3) > a:nth-child(1)').click()
            time.sleep(sleep_table_time)  # Пауза для прогрузки таблицы

            # Получим тело таблицы 'Последние Матчи':
            latest_matches = ui.WebDriverWait(driver, 15).until(
                lambda driver1: driver.find_element_by_css_selector('.fixture > tbody:nth-child(2)'))
            latest_matches_innerhtml = latest_matches.get_property('innerHTML')
            # Достанем из неё информацию о количестве КК дома и в гостях, которые дал судья за последние 20 матчей,
            # дату последней КК:
            soup = BeautifulSoup(latest_matches_innerhtml, 'html.parser')
            # Получим теги всех последних 20 матчей, которые судил текущий судья:
            matches_tags = soup.findAll('tr')
            for match_tag in matches_tags:
                # Пройдёмся по каждому полученному тегу:
                red_cards = match_tag.find('span', {'class': 'incidents-icon ui-icon red'})
                if red_cards is not None:
                    # Красные карточки (одна или больше) найдены в матче.
                    # Если это первая КК, с которой мы столкнулись в таблице, получим её дату (это дата последней КК):
                    if match[i].referee_last_twenty_home_count == 0 and match[i].referee_last_twenty_away_count == 0:
                        match[i].referee_last_twenty_last_kk_date = match_tag.find('td', {'class': 'date'}).text
                    # Получим теги с классами referee-home-data и referee-away-data:
                    referee_home_data = match_tag.find('td', {'class': 'referee-home-data'})
                    # Попробуем найти КК в referee_home_data:
                    red_cards = referee_home_data.find('span', {'class': 'incidents-icon ui-icon red'})
                    if red_cards is not None:
                        # Получим все incidents-wrapper в referee_home_data:
                        incidents_wrappers = referee_home_data.findAll('div', {'class': 'incidents-wrapper'})
                        # Пройдёмся по всем incidents-wrapper в цикле:
                        for incidents_wrapper in incidents_wrappers:
                            # Попробуем найти КК в incidents_wrapper:
                            red_cards = incidents_wrapper.find('span', {'class': 'incidents-icon ui-icon red'})
                            if red_cards is not None:
                                # Если КК присутствуют в incidents_wrapper, спарсим их количество:
                                match[i].referee_last_twenty_home_count += int(incidents_wrapper.text.replace('x', ''))
                                break

                    referee_away_data = match_tag.find('td', {'class': 'referee-away-data'})
                    # Попробуем найти КК в referee_away_data:
                    red_cards = referee_away_data.find('span', {'class': 'incidents-icon ui-icon red'})
                    if red_cards is not None:
                        # Получим все incidents-wrapper в referee_away_data:
                        incidents_wrappers = referee_away_data.findAll('div', {'class': 'incidents-wrapper'})
                        # Пройдёмся по всем incidents-wrapper в цикле:
                        for incidents_wrapper in incidents_wrappers:
                            # Попробуем найти КК в incidents_wrapper:
                            red_cards = incidents_wrapper.find('span', {'class': 'incidents-icon ui-icon red'})
                            if red_cards is not None:
                                # Если КК присутствуют в incidents_wrapper, спарсим их количество:
                                match[i].referee_last_twenty_away_count += int(incidents_wrapper.text.replace('x', ''))
                                break

            # Получим таблицу 'Турниры' (теперь с нажатой кнопкой 'все'):
            referee_tournaments = ui.WebDriverWait(driver, 15).until(
                lambda driver1: driver.find_element_by_id('referee-tournaments-table-body'))
            referee_tournaments_innerhtml = referee_tournaments.get_property('innerHTML')
            # Достанем из неё среднее количество КК за все сезоны:
            soup = BeautifulSoup(referee_tournaments_innerhtml, 'html.parser')
            # Получим теги всех лиг, которые судил текущий судья:
            leagues_tags = soup.findAll('tr')

            for league_tag in leagues_tags:
                # Пройдёмся по каждому полученному тегу и определим, в каком из них находится интересующая нас лига:
                if match[i].league_name in league_tag.text:
                    # Нужная лига найдена! Получим информацию о поведении судьи в этой лиге:
                    td_tags = league_tag.findAll('td')
                    match[i].referee_all_seasons_matches_count = int(td_tags[2].text)
                    match[i].referee_all_seasons_average = td_tags[8].text
                    if match[i].referee_all_seasons_average == '0.00':
                        match[i].referee_all_seasons_average = '0'
                    break

            end_tag_found = False
            while (not end_tag_found) and (not match[i].referee_team_home or not match[i].referee_team_away):
                # Получим таблицу 'Команды' (с нажатой кнопкой 'все'):
                referee_teams = ui.WebDriverWait(driver, 15).until(
                    lambda driver1: driver.find_element_by_css_selector('#referee-team-table-summary > div:nth-child(1)'
                                                                        ' > table:nth-child(1) > tbody:nth-child(2)'))
                referee_teams_innerhtml = referee_teams.get_property('innerHTML')
                # Достанем из неё среднее количество КК для каждой из команд:
                soup = BeautifulSoup(referee_teams_innerhtml, 'html.parser')
                # Получим теги всех команд, которые судил текущий судья:
                teams_tags = soup.findAll('tr')
                for team_tag in teams_tags:
                    tournament_link_tag = team_tag.find('a', {'class': 'tournament-link'})
                    if tournament_link_tag is not None:
                        if match[i].team_home_name in tournament_link_tag.text:
                            match[i].referee_team_home = True
                            td_tags = team_tag.findAll('td')
                            match[i].referee_team_home_matches_count = int(td_tags[2].text)
                            match[i].referee_to_team_home_average = td_tags[8].text
                            if match[i].referee_to_team_home_average == '0.00':
                                match[i].referee_to_team_home_average = '0'
                        if match[i].team_away_name in tournament_link_tag.text:
                            match[i].referee_team_away = True
                            td_tags = team_tag.findAll('td')
                            match[i].referee_team_away_matches_count = int(td_tags[2].text)
                            match[i].referee_to_team_away_average = td_tags[8].text
                            if match[i].referee_to_team_away_average == '0.00':
                                match[i].referee_to_team_away_average = '0'
                        # Если уже нашли обе команды в таблице, то нет смысла листать её дальше:
                        if match[i].referee_team_home and match[i].referee_team_away:
                            break
                    else:
                        # Значит, referee_tag - это итоговый тег таблицы, содержащий строку 'Сумма/Среднее количество':
                        end_tag_found = True
                # Если итоговый тег таблицы не был найден, значит таблицу можно листать дальше:
                if (not end_tag_found) and (not (match[i].referee_team_home and match[i].referee_team_away)):
                    driver.find_element_by_css_selector('#next').click()
                    time.sleep(sleep_table_time)  # Пауза для прогрузки таблицы

            if not match[i].referee_team_home:
                wind.log('Судья '+match[i].referee_name_whoscored+' не судил команду '+match[i].team_home_name+'!')
            if not match[i].referee_team_away:
                wind.log('Судья '+match[i].referee_name_whoscored+' не судил команду '+match[i].team_away_name+'!')
            wind.log('Получили информацию по судье ' + match[i].referee_name_whoscored + '.')


def write_to_spreadsheets():
    ##################################################################################
    # ВЫВОД ДАННЫХ МАССИВА match[] В GOOGLE SHEETS
    ##################################################################################
    wind.log('Запишем полученную информацию в Google Sheets...')
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
                                "columnCount": 16
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
                            'columnCount': 16
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
        spreadsheetId=spreadsheet['spreadsheetId'], range=title+'!A' + str(row_index) + ':P' + str(row_index + 1))\
        .execute()
    two_rows = result.get('values')

    if two_rows is None:
        # Если первые две строки листа пустые, то сделаем заголовок:
        service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'], body={
            'valueInputOption': 'USER_ENTERED',
            'data': [
                {
                    'range': title+'!A1:P2',
                    'majorDimension': 'ROWS',
                    # сначала заполнять ряды, затем столбцы (т.е. самые внутренние списки в values - это ряды)
                    'values':
                    [
                        ['Лига', 'Дома', 'Гости', 'Дата', 'Судья', '', '', '', 'Команды', '', '', '', '', '',
                         'Проверь судей', 'Проверь команды'],
                        ['', '', '', '', 'Этот сезон', 'Все сезоны', 'Командам', 'Посл. 20 игр', 'КК этот сезон',
                         'КК прош. сезон', 'Дата последней 1/2', '', 'Личные встречи', '',
                         'На Championat (на Whoscored)', 'На Championat (на Whoscored)']
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
                            'endColumnIndex': 16
                        },
                        'top': border,
                        'bottom': border,
                        'left': border,
                        'right': border,
                        'innerHorizontal': border,
                        'innerVertical': border
                    }
                },
                # Применим форматирование (цвет фона) к ячейкам заголовка (установим светло-серый цвет):
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 0,
                            "endRowIndex": 2
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": 0.8,
                                    "green": 0.8,
                                    "blue": 0.8
                                }
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor)"
                    }
                },
                # Применим форматирование (цвет фона) к ячейкам заголовка (вернём белый цвет столбцу N):
                {
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 0,
                            "endRowIndex": 2,
                            'startColumnIndex': 13,
                            'endColumnIndex': 14
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": 1.0,
                                    "green": 1.0,
                                    "blue": 1.0
                                }
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor)"
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
                range=title + '!A' + str(row_index) + ':P' + str(row_index + 1)) \
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
                            'endColumnIndex': 16
                        },
                        'top': border
                    }
                }
            ]
        }).execute()

    # Выведем все матчи в цикле:
    previous_league_name = ''
    for i in range(0, matches_length):
        # Подготовим информацию по судье для её последующего вывода:
        referee_this_season = '???'
        referee_all_seasons = '???'
        referee_to_teams = '???'
        referee_last_twenty = '???'
        if match[i].referee_name_championat != '???':
            referee_this_season = \
                match[i].referee_this_season_average + ' (' + str(match[i].referee_this_season_matches_count) + ')'
            referee_all_seasons = \
                match[i].referee_all_seasons_average + ' (' + str(match[i].referee_all_seasons_matches_count) + ')'

            if not match[i].referee_team_home:
                match[i].referee_to_team_home_average = '-'
                match[i].referee_team_home_matches_count = 0

            if not match[i].referee_team_away:
                match[i].referee_to_team_away_average = '-'
                match[i].referee_team_away_matches_count = 0

            referee_to_teams = match[i].referee_to_team_home_average + '/' + match[i].referee_to_team_away_average + \
                ' (' + str(match[i].referee_team_home_matches_count) + '/' + \
                str(match[i].referee_team_away_matches_count) + ')'

            referee_last_twenty = str(match[i].referee_last_twenty_home_count) + 'д' + \
                str(match[i].referee_last_twenty_away_count) + 'г ' + match[i].referee_last_twenty_last_kk_date

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
                    'range': title+'!A'+str(n)+':P'+str(n)+'',
                    'majorDimension': 'ROWS',
                    # сначала заполнять ряды, затем столбцы (т.е. самые внутренние списки в values - это ряды)
                    'values':
                        [
                            [match[i].league_name, match[i].team_home_name, match[i].team_away_name,
                             match[i].match_datetime,
                             referee_this_season, referee_all_seasons,
                             referee_to_teams, referee_last_twenty,
                             str(match[i].team_home_kk_this_season_count) + '/' +
                             str(match[i].team_away_kk_this_season_count),
                             str_team_home_kk_last_season_count + '/' +
                             str_team_away_kk_last_season_count,
                             match[i].team_home_last_kk_date, match[i].team_away_last_kk_date,
                             str_personal_meetings,
                             '',
                             match[i].referee_name_championat + ' (' + match[i].referee_name_whoscored + ')',
                             match[i].championat_teamsstring + ' (' + match[i].teamsstring + ')']
                        ]
                }
            ]
        }).execute()
        n += 1

    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet['spreadsheetId'], body={
        'requests': [
            # Задать размер по содержимому для всех столбцов:
            {
                'autoResizeDimensions': {
                    'dimensions': {
                        'sheetId': sheet_id,
                        'dimension': 'COLUMNS',  # COLUMNS - потому что столбец
                        'startIndex': 0,  # Столбцы нумеруются с нуля
                        'endIndex': 16  # startIndex берётся включительно, endIndex - НЕ включительно
                    }
                }
            },
            # Применим форматирование (выравнивание) к ячейкам:
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 515
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER"
                        }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment)"
                }
            },
            # Изменим вручную ширину столбов 'K' и 'L':
            {
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': sheet_id,
                        'dimension': 'COLUMNS',
                        'startIndex': 10,
                        'endIndex': 12
                    },
                    'properties': {
                        'pixelSize': 82
                    },
                    'fields': 'pixelSize'
                }
            }
        ]
    }).execute()

    wind.log('Полученная информация записана в Google Sheets.')


def main():
    wind.log('Время начала: ' + datetime.datetime.now().strftime('%H:%M'))
    # Получим виджет списка:
    leagues_list = wind.listViewLeagues
    # Получим объект модели, применённый к списку:
    model = leagues_list.model()
    # Объявим список глобальным league[] и очистим его:
    global league
    league.clear()
    # Пройдёмся по каждому элементу списка mdb_league_length[], в который были занесены все лиги из БД:
    for i in range(0, mdb_league_length):
        # Если элемент лиги был отмечен пользователем,
        if model.item(i).checkState() == 2:
            # Добавим его в список лиг, с которыми программа будет работаь далее:
            league.append(mdb_league[i])

    # Определим длину списка league[]:
    global league_length
    league_length = len(league)

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
        # Если матчи найдены:
        # Получим информацию о личных встречах команд:
        get_personal_meetengs()
        # Достанем ссылки на календарь игр прошлых сезонов:
        get_url_games_calendar_past_season()
        # Получим информацию о количестве КК у команд за этот сезон, о дате последней КК:
        get_kk_this_or_last_season(True)
        # Получим информацию о количестве КК у команд за прошлый сезон, о дате последней КК
        # (если она не была найдена в последнем сезоне):
        get_kk_this_or_last_season(False)
        # Получим имя судьи на championat и сопоставим матчи на championat и whoscored:
        get_referee_championat()
        # Получим имя судьи и его Url на whoscored:
        get_referee_whoscored()
        # Получим информацию по судье с whoscored::
        get_referee_info()
        # Запишем полученную информацию в Google Sheets:
        write_to_spreadsheets()
    else:
        wind.log('Ни в одной из лиг не найдено матчей на ' + datestring_format(required_date) + '!')
    # Завершим сессию браузера и закроем его окно:
    driver.quit()
    # Включим кнопку startButton:
    wind.startButton.setEnabled(True)

    wind.log('Время завершения: ' + datetime.datetime.now().strftime('%H:%M'))

    # Если checkBoxShutdown выбран:
    if wind.checkBoxShutdown.isChecked():
        # Выключить компьютер через 10 сек:
        os.system('shutdown /s /t 10')


if __name__ == '__main__':
    # Список, куда будут занесены все лиги из БД:
    mdb_league = []
    mdb_league_length = 0

    # Список, который будет сформирован на основе тех лиг, которые выбрал пользователь из списка mdb_league[]:
    league = []
    league_length = 0

    # Список, в который будут занесены все матчи и информация по ним. Основной список в программе:
    match = []
    matches_length = 0

    # Временные задержки для того, чтобы сервер не разорвал соединение по причине слишком частых обращений:
    sleep_page_time = 5  # Задержка после загрузки страницы
    sleep_table_time = 1  # Задержка после пролистывания таблицы
    # Получаем текущую локаль:
    default_loc = locale.getlocale()
    # Изменяем локаль для корректной конвертации строки с русской датой в datetime:
    locale.setlocale(locale.LC_ALL, ('RU', 'UTF8'))

    app = QApplication(sys.argv)
    wind = Window()
    app.exec_()

    # Меняем локаль назад:
    locale.setlocale(locale.LC_ALL, default_loc)
