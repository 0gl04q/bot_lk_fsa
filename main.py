import time

from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException as Nse


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


# Функция авторизации в РА
def authorization_ra(d, org):
    # Переход на ФСА после входа в EСИА
    d.get("http://10.250.74.17/auth/?redirect_uri=http://10.250.74.17/logout")

    time.sleep(5)  # ожидание

    # Нажатие на кнопку "Вход в ЕСИА"
    d.find_element(By.XPATH, "/html/body/div/div[2]/div/div[2]/a[1]").click()

    # Переключение рабочей вкладки
    windows = driver.window_handles  # Список всех вкладок

    time.sleep(5)  # ожидание

    d.close()  # Закрываем первое окно
    d.switch_to.window(windows[1])  # Переключаем драйвер на рабочую область

    time.sleep(5)  # ожидание

    match org:
        case 'МС':
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

        case 'ATM':
            # TODO: дописать кейс для организации АТМ
            pass

    # Нажатие на кнопку входа в личный кабинет
    d.find_element(By.XPATH,
                   "/html/body/fgis-root/div/fgis-lk/div/fgis-lk-selector/div[5]/div/div[2]/button").click()

    time.sleep(5)  # ожидание


# Функция внесения сведений о поверках
def update_info(d, r):
    # Переход в рабочую область
    d.get("http://10.250.74.17/roei/verification-measuring-instruments")

    time.sleep(5)  # ожидание

    # Нажатие на кнопку "Добавить"
    d.find_element(By.XPATH,
                   "/html/body/fgis-root/div/fgis-roei/fgis-roei-verification-measuring-instruments/div/div/div["
                   "1]/fgis-table-toolbar/section/div/div[1]/div/fgis-toolbar/div/div["
                   "1]/fgis-toolbar-button").click()

    time.sleep(2)  # ожидание

    # Внесение "Номер рез. поверки СИ"
    d.find_element(By.XPATH, "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments-card"
                             "-edit/div/div/div/div/fgis-verification-measuring-instruments-card-edit-common"
                             "/fgis-card-block/div/div[2]/div/fgis-card-edit-row-two-columns["
                             "1]/fgis-card-edit-row[1]/div["
                             "2]/fgis-field-input/fgis-field-wrapper/div/div/input").send_keys(r[0].value)

    time.sleep(2)  # ожидание

    # Внесение "Тип поверяемого средства измерения"
    d.find_element(By.XPATH, '/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring'
                             '-instruments-card-edit/div/div/div/div/fgis-verification-measuring'
                             '-instruments-card-edit-common/fgis-card-block/div/div['
                             '2]/div/fgis-card-edit-row-two-columns[1]/fgis-card-edit-row[2]/div['
                             '2]/fgis-field-input/fgis-field-wrapper/div/div/input').send_keys(r[1].value)

    time.sleep(2)  # ожидание

    # Внесение "Дата результатов поверки"
    d.find_element(By.XPATH, "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments-card"
                             "-edit/div/div/div/div/fgis-verification-measuring-instruments-card-edit-common"
                             "/fgis-card-block/div/div[2]/div/fgis-card-edit-row-two-columns["
                             "2]/fgis-card-edit-row[1]/div["
                             "2]/fgis-field-calendar/fgis-field-wrapper/div/div/fgis-calendar/div/div/inp"
                             "ut").send_keys(r[2].value)

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

    # Выбор действий по условию
    if r[4].value == 'Пригодно':

        # Выбор элемент "Пригодно"
        d.find_element(By.XPATH, "/html/body/fgis-root/fgis-select-dropdown/div/div/div["
                                 "2]/fgis-virtual-list/div/div[2]/div[1]/fgis-virtual-list-item").click()

        time.sleep(2)  # ожидание

        # Внесение "Дата действия результатов поверки"
        d.find_element(By.XPATH, "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments"
                                 "-card-edit/div/div/div/div/fgis-verification-measuring-instruments-card-edit"
                                 "-common/fgis-card-block/div/div[2]/div/fgis-card-edit-row-two-columns["
                                 "2]/fgis-card-edit-row[2]/div["
                                 "2]/fgis-field-calendar/fgis-field-wrapper/div/div/fgis-calendar/div/div/input"
                                 "").send_keys(r[3].value)
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

    # Внесение в поле фамилии поверителя
    d.find_element(By.XPATH, "/html/body/fgis-root/fgis-select-dropdown/div/div/div[1]/input").send_keys(r[5].value)

    time.sleep(5)  # ожидание

    # Выбор фамилии из списка
    if r[5].value == 'Гипич':  # для Гипич отдельное условие так как есть совпадение фамилий
        try:
            d.find_element(By.XPATH,
                           "/html/body/fgis-root/fgis-select-dropdown/div/div/div[2]/fgis-virtual-list/div/div["
                           "2]/div[2]/fgis-virtual-list-item").click()
        except Nse:
            print('ошибка: NSE')
            d.refresh()
            return 'refresh'
    else:
        try:
            d.find_element(By.XPATH,
                           "/html/body/fgis-root/fgis-select-dropdown/div/div/div[2]/fgis-virtual-list/div/div["
                           "2]/div/fgis-virtual-list-item").click()
        except Nse:
            print('ошибка: NSE')
            d.refresh()
            return 'refresh'

    time.sleep(5)  # ожидание

    # нажатие на подтверждение заполненности формы
    d.find_element(By.XPATH, "/html/body/fgis-root/div/fgis-roei/fgis-verification-measuring-instruments-card-edit"
                             "/fgis-verification-measuring-instruments-card-edit-toolbar/div/fgis-toolbar/div/div["
                             "1]/fgis-toolbar-button").click()

    return 'Внесено'


# Функция публикации сведения в РА
def public_info(d, num_res):
    # Переход в рабочую область
    d.get("http://10.250.74.17/roei/verification-measuring-instruments")

    time.sleep(5)  # ожидание

    # Вставка в поле номера рез.пов в соответствующее поле отбора
    d.find_element(By.XPATH, '/html/body/fgis-root/div/fgis-roei/fgis-roei-verification-measuring-instruments/div'
                             '/fgis-roei-verification-measuring-instruments-advanced-search/fgis-filters-panel/fgis'
                             '-left-panel/div[2]/div[2]/div/div['
                             '1]/fgis-field-input/fgis-field-wrapper/div/div/input').send_keys(num_res)

    time.sleep(2)  # ожидание

    # Нажатие на кнопку найти
    d.find_element(By.XPATH, '/html/body/fgis-root/div/fgis-roei/fgis-roei-verification-measuring-instruments/div'
                             '/fgis-roei-verification-measuring-instruments-advanced-search/fgis-filters-panel/fgis'
                             '-left-panel/div[3]/div/button').click()

    time.sleep(10)  # ожидание

    # Выбор первого элемента записи
    d.find_element(By.XPATH, '/html/body/fgis-root/div/fgis-roei/fgis-roei-verification-measuring-instruments/div/div'
                             '/div[2]/fgis-table/div[2]/div/table/tbody/a/td[1]/div').click()

    time.sleep(2)  # ожидание

    # Нажатие на кнопку "Опубликовать"
    d.find_element(By.XPATH, '/html/body/fgis-root/div/fgis-roei/fgis-roei-verification-measuring-instruments/div/div'
                             '/div[1]/fgis-table-toolbar/section/div/div[1]/div/fgis-toolbar/div/div['
                             '4]/fgis-toolbar-button/button').click()

    time.sleep(2)  # ожидание

    # Нажатие на кнопку "Подтверждаю"
    d.find_element(By.XPATH, '/html/body/fgis-root/div/fgis-roei/fgis-roei-verification-measuring-instruments/fgis'
                             '-modal/div/div[1]/div[3]/div/button[1]').click()

    return 'Опубликовано'


if __name__ == '__main__':

    # TODO: прикрутить проход по папке
    func = lambda x: x if x != '' else 'МС 28.02.2023 - 28.02.2023.xlsx'

    file_name = f'file\\{func(input("Введите наименование файла: "))}'

    organization = file_name.split()[0]

    # получение драйвера
    driver = browse(organization)

    # прохождение авторизации в RA
    authorization_ra(driver, organization)

    # Открытие файла Excel
    wb = load_workbook(filename=file_name)
    sheet = wb.active

    # Проходим по строкам
    for row in sheet.iter_rows(min_row=2):
        if row[1].value:  # проверка на обязательное наличие сведений

            # проверка поля статуса Если статус пустой тогда переходим к созданию черновика
            if row[6].value is None:
                call = update_info(driver, row)  # Заполняем и сохраняем черновик

                # отработка исключения
                while call == 'refresh':
                    time.sleep(10)  # ожидание
                    driver = browse(organization)  # обновляем объект браузера
                    authorization_ra(driver, organization)  # повторно проходим круг с авторизацией
                    call = update_info(driver, row)  # Заполняем и сохраняем черновик

                row[6].value = call
                time.sleep(5)  # ожидание
                wb.save(file_name)  # сохраняем статус

            # Если статус = Внесено тогда приступаем к публикации сведения
            if row[6].value == 'Внесено':
                row[6].value = public_info(driver, row[0].value)  # публикуем запись
                time.sleep(5)  # ожидание
                wb.save(file_name)  # сохраняем статус
        else:
            break

    wb.close()

    driver.close()
