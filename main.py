import time

from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException as Nse
from selenium.common.exceptions import StaleElementReferenceException as Sere
from selenium.common.exceptions import ElementClickInterceptedException as E_click
from selenium.common.exceptions import TimeoutException as Te


# ожидание для нажатия на элемент
def wait_click(d, path):
    try:  # Ждем пока элемент не появится, потом на него нажимаем
        WebDriverWait(d, 60).until(
            EC.presence_of_element_located((By.XPATH, path))
        ).click()
    except Sere or E_click:  # обрабатываем исключение ожиданием и рекурсией
        time.sleep(5)
        wait_click(d, path)


# ожидание для внедрения значения
def wait_send_keys(d, path, paste):
    try:  # Ждем пока элемент не появится, потом вставляем в него значение
        WebDriverWait(d, 60).until(
            EC.presence_of_element_located((By.XPATH, path))
        ).send_keys(paste)
    except Sere:  # обрабатываем исключение ожиданием и рекурсией
        time.sleep(5)
        wait_send_keys(d, path, paste)


# Функция создания объекта браузера с необходимыми параметрами
def browse(org):
    options = webdriver.ChromeOptions()

    # Добавление пути к профилю Хром
    options.add_argument(r"user-data-dir=C:\Users\РДС\AppData\Local\Google\Chrome\User Data")

    # TODO: Если организация МС выбираем профиль 1 иначе если АТМ выбираем профиль 2
    match org:
        case 'МС':
            # Указание папки профиля
            options.add_argument("profile-directory=Profile 1")
        case 'АТМ':
            # Указание папки профиля
            options.add_argument("profile-directory=Profile 2")

    # Путь и подключение плагина Госуслуг TODO: прописать универсальный путь
    path_to_extension = r'C:\Users\РДС\Pictures\crx\gu.crx'
    options.add_extension(path_to_extension)

    # Отключение метки об автоматическом контроле
    options.add_argument("--disable-blink-features=AutomationControlled")

    # Создание объекта класса Chrome
    return webdriver.Chrome(options=options)


# Функция внесения сведений о поверках
def update_info(d, r):
    # Переход в рабочую область
    d.get("http://10.250.74.17/roei/verification-measuring-instruments")

    # ожидание
    # Нажатие на кнопку "Добавить"
    wait_click(d, "/html/body/fgis-root/div/fgis-roei/fgis-roei-verification-measuring-instruments/div/div/div["
                  "1]/fgis-table-toolbar/section/div/div[1]/div/fgis-toolbar/div/div[1]/fgis-toolbar-button")

    # Внесение "Номер рез. поверки СИ"
    wait_send_keys(d, "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments-card"
                      "-edit/div/div/div/div/fgis-verification-measuring-instruments-card-edit-common"
                      "/fgis-card-block/div/div[2]/div/fgis-card-edit-row-two-columns["
                      "1]/fgis-card-edit-row[1]/div["
                      "2]/fgis-field-input/fgis-field-wrapper/div/div/input", r[0].value)

    # Внесение "Тип поверяемого средства измерения"
    wait_send_keys(d, '/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring'
                      '-instruments-card-edit/div/div/div/div/fgis-verification-measuring'
                      '-instruments-card-edit-common/fgis-card-block/div/div['
                      '2]/div/fgis-card-edit-row-two-columns[1]/fgis-card-edit-row[2]/div['
                      '2]/fgis-field-input/fgis-field-wrapper/div/div/input', r[1].value)

    # Внесение "Дата результатов поверки"
    wait_send_keys(d, "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments-card"
                      "-edit/div/div/div/div/fgis-verification-measuring-instruments-card-edit-common"
                      "/fgis-card-block/div/div[2]/div/fgis-card-edit-row-two-columns["
                      "2]/fgis-card-edit-row[1]/div["
                      "2]/fgis-field-calendar/fgis-field-wrapper/div/div/fgis-calendar/div/div/inp"
                      "ut", r[2].value)

    # Закрытие формы "Дата результатов поверки" нажатием на экран
    wait_click(d, '/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments-card-edit/div/div/div'
                  '/div/fgis-verification-measuring-instruments-card-edit-common/fgis-card-block/div/div[1]/div[1]')

    # Выбор пригодности результата поверки - совершается с помощью 2 действий
    # Открытие списка выбора
    wait_click(d, "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments-card-edit/div/div"
                  "/div/div/fgis-verification-measuring-instruments-card-edit-common/fgis-card-block/div/div["
                  "2]/div/fgis-card-edit-row[1]/div["
                  "2]/fgis-field-selectbox/fgis-field-wrapper/div/div/fgis-selectbox/div/div/div[1]")

    # Выбор действий по условию
    if r[4].value == 'Пригодно':

        # Выбор элемент "Пригодно"
        wait_click(d, "/html/body/fgis-root/fgis-select-dropdown/div/div/div["
                      "2]/fgis-virtual-list/div/div[2]/div[1]/fgis-virtual-list-item")

        # Внесение "Дата действия результатов поверки"
        wait_send_keys(d, "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments"
                          "-card-edit/div/div/div/div/fgis-verification-measuring-instruments-card-edit"
                          "-common/fgis-card-block/div/div[2]/div/fgis-card-edit-row-two-columns["
                          "2]/fgis-card-edit-row[2]/div["
                          "2]/fgis-field-calendar/fgis-field-wrapper/div/div/fgis-calendar/div/div/input", r[3].value)

        # Закрытие формы "Дата результатов поверки" нажатием на экран
        wait_click(d, '/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments-card-edit/div/div/div'
                      '/div/fgis-verification-measuring-instruments-card-edit-common/fgis-card-block/div/div[1]/div[1]')

        # # Закрытие формы "Дата действия результатов поверки"
        # wait_click(d, "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments"
        #               "-card-edit/div/div/div/div/fgis-verification-measuring-instruments-card-edit"
        #               "-common/fgis-card-block/div/div[2]/div/fgis-card-edit-row-two-columns["
        #               "2]/fgis-card-edit-row[2]/div["
        #               "2]/fgis-field-calendar/fgis-field-wrapper/div/div/fgis-calendar/div/div")

    else:
        # Выбор элемент "Непригодно"
        wait_click(d, "/html/body/fgis-root/fgis-select-dropdown/div/div/div[2]/fgis-virtual-list/div/div[2]/div["
                      "2]/fgis-virtual-list-item")

        # Выбор лица проводившего поверку - совершается с помощью 2 действий
    # Открытие списка выбора
    wait_click(d, "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments-card"
                  "-edit/div/div/div/div/fgis-verification-measuring-instruments-card-edit-common"
                  "/fgis-card-block/div/div[2]/div/fgis-card-edit-row[2]/div["
                  "2]/fgis-field-selectbox/fgis-field-wrapper/div/div/fgis-selectbox/div/div/div[1]")

    # Внесение в поле фамилии поверителя
    wait_send_keys(d, "/html/body/fgis-root/fgis-select-dropdown/div/div/div[1]/input", r[5].value)

    # Выбор фамилии из списка
    if r[5].value == 'Гипич':  # для Гипич отдельное условие так как есть совпадение фамилий
        try:
            wait_click(d, "/html/body/fgis-root/fgis-select-dropdown/div/div/div[2]/fgis-virtual-list/div/div[2]/div["
                          "2]/fgis-virtual-list-item")
        except Nse:
            print('ошибка: NSE')
            d.refresh()
            return 'refresh'
    else:
        try:
            wait_click(d, "/html/body/fgis-root/fgis-select-dropdown/div/div/div[2]/fgis-virtual-list/div/div["
                          "2]/div/fgis-virtual-list-item")
        except Nse:
            print('ошибка: NSE')
            d.refresh()
            return 'refresh'

    # нажатие на подтверждение заполненности формы
    wait_click(d, "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments-card-edit"
                  "/fgis-verification-measuring-instruments-card-edit-toolbar/div/fgis-toolbar/div/div["
                  "1]/fgis-toolbar-button")

    return 'Черновик РА'


# Функция публикации сведения в РА
def public_info(d, num_res):
    # Переход в рабочую область
    d.get("http://10.250.74.17/roei/verification-measuring-instruments")

    # Вставка в поле номера рез.пов в соответствующее поле отбора
    wait_send_keys(d, '/html/body/fgis-root/div/fgis-roei/fgis-roei-verification-measuring-instruments/div'
                      '/fgis-roei-verification-measuring-instruments-advanced-search/fgis-filters-panel/fgis'
                      '-left-panel/div[2]/div[2]/div/div['
                      '1]/fgis-field-input/fgis-field-wrapper/div/div/input', num_res)

    # Нажатие на кнопку найти
    wait_click(d, '/html/body/fgis-root/div/fgis-roei/fgis-roei-verification-measuring-instruments/div'
                  '/fgis-roei-verification-measuring-instruments-advanced-search/fgis-filters-panel/fgis'
                  '-left-panel/div[3]/div/button')

    # Ожидание появления строки с необходимым номером
    WebDriverWait(d, 60).until(
        EC.text_to_be_present_in_element((By.XPATH, '/html/body/fgis-root/div/fgis-roei/fgis-roei-verification'
                                                    '-measuring-instruments/div/div '
                                                    '/div[2]/fgis-table/div[2]/div/table/tbody/a[1]/td['
                                                    '5]/div/div[1]/span'), str(num_res)))

    # Выбор первого элемента записи
    wait_click(d, '/html/body/fgis-root/div/fgis-roei/fgis-roei-verification-measuring-instruments/div/div'
                  '/div[2]/fgis-table/div[2]/div/table/tbody/a/td[1]/div')

    # Нажатие на кнопку "Опубликовать"
    wait_click(d, '/html/body/fgis-root/div/fgis-roei/fgis-roei-verification-measuring-instruments/div/div'
                  '/div[1]/fgis-table-toolbar/section/div/div[1]/div/fgis-toolbar/div/div['
                  '4]/fgis-toolbar-button/button')

    # Нажатие на кнопку "Подтверждаю"
    wait_click(d, '/html/body/fgis-root/div/fgis-roei/fgis-roei-verification-measuring-instruments/fgis'
                  '-modal/div/div[1]/div[3]/div/button[1]')

    return 'Опубликовано в РА'


# Функция авторизации в ГУ
def authorization_gu(d, org):
    d.get("https://esia.gosuslugi.ru/login/")
    match org:
        case 'МС':
            wait_click(d, '/html/body/esia-root/div/esia-login/div/div[1]/form/div[4]/div/div[2]/div/button')
            wait_click(d, '/html/body/esia-root/div/esia-login/div/div[1]/esia-eds/button')
            wait_click(d, '/html/body/esia-root/esia-modal-placeholder/div/div[2]/div/div/div/ng-component/div/esia-eds-item[1]')
            wait_click(d, '/html/body/esia-root/esia-modal-placeholder/div/div[2]/div/div/div/ng-component/div/div[2]/a')
        case 'АТМ':
            # TODO: дописать кейс для организации АТМ
            pass


# Функция авторизации в РА
def authorization_ra(d, org):
    # Переход на ФСА после входа в EСИА
    d.get("http://10.250.74.17/auth/?redirect_uri=http://10.250.74.17/logout")

    # ожидание
    # Нажатие на кнопку "Вход в ЕСИА"
    wait_click(d, "/html/body/div/div[2]/div/div[2]/a[1]")

    # Переключение рабочей вкладки
    windows = driver.window_handles  # Список всех вкладок

    time.sleep(5)  # ожидание

    d.close()  # Закрываем первое окно
    d.switch_to.window(windows[1])  # Переключаем драйвер на рабочую область

    time.sleep(5)  # ожидание

    match org:
        case 'МС':
            # Выбор организации "МС"
            wait_click(d, "/html/body/div/div[2]/div/div/table/tbody/tr[2]/td/a")

            # Открытие списка выбора рабочей области
            wait_click(d, "/html/body/fgis-root/div/fgis-lk/div/fgis-lk-selector/div[5]/div/div["
                          "2]/fgis-selectbox/div/div/div[2]")

            # Выбор рабочей области в нашем случае АЛ
            wait_click(d, "/html/body/fgis-root/fgis-select-dropdown/div/div/div[2]/fgis-virtual-list/div/div[2]/div["
                          "1]/fgis-virtual-list-item")

        case 'ATM':
            # TODO: дописать кейс для организации АТМ
            pass

    # Нажатие на кнопку входа в личный кабинет
    wait_click(d, "/html/body/fgis-root/div/fgis-lk/div/fgis-lk-selector/div[5]/div/div[2]/button")


# Функция для определения продолжить работу или начать авторизацию
def all_authorization(d, org):
    # переход в рабочую область
    d.get('http://10.250.74.17/roei/verification-measuring-instruments')

    # Ожидание появления строки с необходимым номером
    try:
        WebDriverWait(d, 10).until(EC.text_to_be_present_in_element((By.XPATH, '/html/body/fgis-root/div/fgis'
                                                                               '-header/header/div['
                                                                               '1]/fgis-title/div/div'),
                                                                    'Сведения по обеспечению единства '
                                                                    'измерений'))
    except Te:
        # прохождение авторизации В ГУ
        authorization_gu(d, org)
        # прохождение авторизации в RA
        authorization_ra(d, org)


if __name__ == '__main__':

    # TODO: прикрутить проход по папке
    file_name = 'МС 28.02.2023 - 28.02.2023.xlsx'

    organization = file_name.split()[0]

    # получение драйвера
    driver = browse(organization)

    # Проверка на необходимость авторизации
    all_authorization(driver, organization)

    # Открытие файла Excel
    wb = load_workbook(filename=file_name)
    sheet = wb.active

    # Проходим по строкам
    for row in sheet.iter_rows(min_row=2):
        if row[1].value:  # проверка на обязательное наличие сведений

            # проверка поля статуса Если статус Выгружена в АРШИН тогда переходим к созданию черновика
            if row[6].value == 'Выгружена в АРШИН':
                call = update_info(driver, row)  # Заполняем и сохраняем черновик

                # отработка исключения
                while call == 'refresh':
                    time.sleep(10)  # ожидание
                    driver = browse(organization)  # обновляем объект браузера
                    all_authorization(driver, organization)  # повторно проходим круг с авторизацией
                    call = update_info(driver, row)  # Заполняем и сохраняем черновик

                row[6].value = call
                time.sleep(5)  # ожидание
                wb.save(file_name)  # сохраняем статус

            # Если статус = Черновик РА тогда приступаем к публикации сведения
            if row[6].value == 'Черновик РА':
                row[6].value = public_info(driver, row[0].value)  # публикуем запись
                time.sleep(5)  # ожидание
                wb.save(file_name)  # сохраняем статус
        else:
            break

    wb.close()

    driver.close()
