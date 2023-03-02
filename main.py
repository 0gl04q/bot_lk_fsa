import time

from selenium import webdriver
from selenium.webdriver.common.by import By


# Функция создания объекта браузера с необходимыми параметрами
def browse():
    options = webdriver.ChromeOptions()

    # Добавление пути к профилю Хром
    options.add_argument(r"user-data-dir=C:\Users\РДС\AppData\Local\Google\Chrome\User Data")

    # Указание папки профиля
    options.add_argument("profile-directory=Profile 1")

    # Путь и подключение плагина Госуслуг
    path_to_extension = r'C:\Users\РДС\Pictures\crx\gu.crx'
    options.add_extension(path_to_extension)

    # Отключение метки об автоматическом контроле
    options.add_argument("--disable-blink-features=AutomationControlled")

    # Создание объекта класса Chrome
    return webdriver.Chrome(options=options)


# Функция авторизации в РА
def ra_authorization(d):
    # Переход на ФСА после входа в EСИА
    d.get("http://10.250.74.17/auth/?redirect_uri=http://10.250.74.17/logout")

    time.sleep(5)  # ожидание

    # Нажатие на кнопку "Вход в ЕСИА"
    d.find_element(By.XPATH, "/html/body/div/div[2]/div/div[2]/a[1]").click()

    time.sleep(5)  # ожидание

    # Переход на страницу аутентификации
    d.get("https://auth.fsa.gov.ru/oauth2/lk")

    time.sleep(5)  # ожидание

    # Выбор организации "МС"
    d.find_element(By.XPATH,
                   "/html/body/div/div[2]/div/div/table/tbody/tr[2]/td/a").click()

    time.sleep(5)  # ожидание

    # Открытие списка выбора рабочей области
    d.find_element(By.XPATH,
                   "/html/body/fgis-root/div/fgis-lk/div/fgis-lk-selector/div[5]/div/div["
                   "2]/fgis-selectbox/div/div/div[2]").click()

    time.sleep(2)  # ожидание

    # Выбор рабочей области в нашем случае АЛ
    d.find_element(By.XPATH,
                   "/html/body/fgis-root/fgis-select-dropdown/div/div/div[2]/fgis-virtual-list/div/div[2]/div["
                   "1]/fgis-virtual-list-item").click()

    time.sleep(2)  # ожидание

    # Нажатие на кнопку входа в личный кабинет
    d.find_element(By.XPATH,
                   "/html/body/fgis-root/div/fgis-lk/div/fgis-lk-selector/div[5]/div/div[2]/button").click()

    time.sleep(5)  # ожидание


# Функция внесения сведений о поверках
def update_info(d):

    # Переход в рабочую область
    d.get("http://10.250.74.17/roei/verification-measuring-instruments")

    time.sleep(5)  # ожидание

    # Нажатие на кнопку "Добавить"
    d.find_element(By.XPATH,
                   "/html/body/fgis-root/div/fgis-roei/fgis-roei-verification-measuring-instruments/div/div/div["
                   "1]/fgis-table-toolbar/section/div/div[1]/div/fgis-toolbar/div/div["
                   "1]/fgis-toolbar-button").click()

    time.sleep(2)  # ожидание

    # Внесение "Номер рез. поверки СИ" TODO: автоматизировать получение сведений из поля xls в send_keys
    d.find_element(By.XPATH, "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments-card"
                             "-edit/div/div/div/div/fgis-verification-measuring-instruments-card-edit-common"
                             "/fgis-card-block/div/div[2]/div/fgis-card-edit-row-two-columns["
                             "1]/fgis-card-edit-row[1]/div["
                             "2]/fgis-field-input/fgis-field-wrapper/div/div/input").send_keys('223737737')

    time.sleep(2)  # ожидание

    # Внесение "Тип поверяемого средства измерения" TODO: автоматизировать получение сведений из поля xls в send_keys
    d.find_element(By.XPATH, '/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring'
                             '-instruments-card-edit/div/div/div/div/fgis-verification-measuring'
                             '-instruments-card-edit-common/fgis-card-block/div/div['
                             '2]/div/fgis-card-edit-row-two-columns[1]/fgis-card-edit-row[2]/div['
                             '2]/fgis-field-input/fgis-field-wrapper/div/div/input').send_keys('TIP SR')

    time.sleep(2)  # ожидание

    # Внесение "Дата результатов поверки" TODO: автоматизировать получение сведений из поля xls в send_keys
    d.find_element(By.XPATH, "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments-card"
                             "-edit/div/div/div/div/fgis-verification-measuring-instruments-card-edit-common"
                             "/fgis-card-block/div/div[2]/div/fgis-card-edit-row-two-columns["
                             "2]/fgis-card-edit-row[1]/div["
                             "2]/fgis-field-calendar/fgis-field-wrapper/div/div/fgis-calendar/div/div/inp"
                             "ut").send_keys('01.03.2023')

    time.sleep(2)  # ожидание

    # Закрытие формы "Дата результатов поверки"
    d.find_element(By.XPATH,
                   "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments-card-edit/div/div"
                   "/div/div/fgis-verification-measuring-instruments-card-edit-common/fgis-card-block/div/div["
                   "2]/div/fgis-card-edit-row-two-columns[2]/fgis-card-edit-row[1]/div["
                   "2]/fgis-field-calendar/fgis-field-wrapper/div/div/fgis-calendar/div/div").click()

    # Выбор пригодности результата поверки - совершается с помощью 2 действий
    # Открытие списка выбора
    d.find_element(By.XPATH,
                   "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments-card-edit/div/div"
                   "/div/div/fgis-verification-measuring-instruments-card-edit-common/fgis-card-block/div/div["
                   "2]/div/fgis-card-edit-row[1]/div["
                   "2]/fgis-field-selectbox/fgis-field-wrapper/div/div/fgis-selectbox/div/div/div[1]").click()

    time.sleep(2)  # ожидание

    # TODO: прикрутить выбор из xls
    pri = True
    if pri:
        # Выбор элемент "Пригодно"
        d.find_element(By.XPATH, "/html/body/fgis-root/fgis-select-dropdown/div/div/div["
                                 "2]/fgis-virtual-list/div/div[2]/div[1]/fgis-virtual-list-item").click()

        time.sleep(2)  # ожидание

        # Внесение "Дата действия результатов поверки" TODO: автоматизировать получение сведений из поля xls в send_keys
        d.find_element(By.XPATH, "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments"
                                 "-card-edit/div/div/div/div/fgis-verification-measuring-instruments-card-edit"
                                 "-common/fgis-card-block/div/div[2]/div/fgis-card-edit-row-two-columns["
                                 "2]/fgis-card-edit-row[2]/div["
                                 "2]/fgis-field-calendar/fgis-field-wrapper/div/div/fgis-calendar/div/div/input"
                                 "").send_keys('28.02.2028')
        time.sleep(2)  # ожидание

        # Закрытие формы "Дата действия результатов поверки"
        d.find_element(By.XPATH, "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments"
                                 "-card-edit/div/div/div/div/fgis-verification-measuring-instruments-card-edit"
                                 "-common/fgis-card-block/div/div[2]/div/fgis-card-edit-row-two-columns["
                                 "2]/fgis-card-edit-row[2]/div["
                                 "2]/fgis-field-calendar/fgis-field-wrapper/div/div/fgis-calendar/div/div"
                                 "").click()

    else:
        # Выбор элемент "Непригодно"
        d.find_element(By.XPATH, "/html/body/fgis-root/fgis-select-dropdown/div/div/div["
                                 "2]/fgis-virtual-list/div/div[2]/div[2]/fgis-virtual-list-item").click()

    time.sleep(2)  # ожидание

    # Выбор лица проводившего поверку - совершается с помощью 2 действий
    # Открытие списка выбора
    d.find_element(By.XPATH, "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments-card"
                             "-edit/div/div/div/div/fgis-verification-measuring-instruments-card-edit-common"
                             "/fgis-card-block/div/div[2]/div/fgis-card-edit-row[2]/div["
                             "2]/fgis-field-selectbox/fgis-field-wrapper/div/div/fgis-selectbox/div/div/div["
                             "1]").click()

    time.sleep(2)  # ожидание

    # Внесение в поле фамилии поверителя TODO: прописать индивидуальное условие для ГИК
    # TODO: автоматизировать получение сведений из поля xls в send_keys
    d.find_element(By.XPATH, "/html/body/fgis-root/fgis-select-dropdown/div/div/div[1]/input").send_keys('Гипич')

    time.sleep(5)  # ожидание

    # Выбор фамилии из списка
    d.find_element(By.XPATH, "/html/body/fgis-root/fgis-select-dropdown/div/div/div["
                             "2]/fgis-virtual-list/div/div[2]/div[2]/fgis-virtual-list-item").click()


if __name__ == '__main__':
    # получение драйвера
    driver = browse()

    # прохождение авторизации в RA
    ra_authorization(driver)

    # TODO: Прикрутить цикл для прохода сведений по файлу xls
    while True:
        update_info(driver)

        time.sleep(1000)

        break

    # assert "No results found." not in driver.page_source
    driver.close()
