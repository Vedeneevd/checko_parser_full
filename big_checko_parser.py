import json
import os
import random
import time
from datetime import datetime, timedelta
import pandas as pd
import requests
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import logging

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('parser_monthly.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Конфигурация
BASE_URL = "https://checko.ru/search/advanced"
PAGE_LOAD_TIMEOUT = 30
MAX_RETRIES = 3
DELAY_BETWEEN_PAGES = 2  # Задержка между страницами в секундах
API_KEY = os.getenv('API_KEY')  # API ключ для rucaptcha
SMTPBZ_API_KEY = os.getenv('SMTPBZ_API_KEY')

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager


def setup_driver():
    """Настройка веб-драйвера для работы на VPS"""
    options = Options()

    # Настройки для работы в headless-режиме
    options.add_argument("--headless")  # Запуск без графического интерфейса
    options.add_argument("--disable-gpu")  # Отключение GPU
    options.add_argument("--no-sandbox")  # Отключение песочницы (для работы в контейнерах или VPS)
    options.add_argument("--disable-dev-shm-usage")  # Отключение ограничений памяти
    options.add_argument("start-maximized")  # Запуск в полноэкранном режиме (не обязательно)

    # Установка User-Agent для имитации обычного браузера
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )

    # Для предотвращения блокировки автоматизации
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    # Устанавливаем путь к браузеру (если нужно для нестандартных путей, например для Chromium)
    options.binary_location = '/usr/bin/chromium-browser'

    # Запуск веб-драйвера с использованием ChromeDriverManager для автоматической установки драйвера
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    # Для скрытия информации о WebDriver
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })

    return driver


def solve_recaptcha_v2(driver):
    """Полное решение reCAPTCHA v2 с отладкой"""
    print("Начинаем решение reCAPTCHA v2...")
    debug_screenshot(driver, "before_solving")

    try:
        # Получаем параметры капчи
        sitekey = driver.find_element(By.CSS_SELECTOR, 'div[data-sitekey]').get_attribute("data-sitekey")
        pageurl = driver.current_url
        print(f"Sitekey: {sitekey}, URL: {pageurl}")

        # 1. Создаем задачу в API
        payload = {
            "clientKey": API_KEY,
            "task": {
                "type": "RecaptchaV2TaskProxyless",
                "websiteURL": pageurl,
                "websiteKey": sitekey,
                "isInvisible": False
            },
            "softId": 3898
        }

        response = requests.post(
            "https://api.rucaptcha.com/createTask",
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=30
        )
        result = response.json()

        if result.get("errorId") != 0:
            error = result.get("errorDescription", "Неизвестная ошибка API")
            raise Exception(f"Ошибка API при создании задачи: {error}")

        task_id = result["taskId"]
        print(f"Задача создана, ID: {task_id}")
        debug_screenshot(driver, "task_created")

        # 2. Ожидаем решения
        start_time = time.time()
        while time.time() - start_time < 300:  # 5 минут максимум
            time.sleep(10)

            status_response = requests.post(
                "https://api.rucaptcha.com/getTaskResult",
                json={"clientKey": API_KEY, "taskId": task_id},
                headers={"Content-Type": "application/json"},
                timeout=30
            ).json()

            print(f"Статус решения: {json.dumps(status_response, indent=2)}")

            if status_response.get("status") == "ready":
                token = status_response["solution"]["gRecaptchaResponse"]
                print("Капча успешно решена!")

                # 3. Вводим токен
                driver.execute_script(f"""
                    var response = document.getElementById('g-recaptcha-response');
                    if (response) {{
                        response.style.display = '';
                        response.value = '{token}';
                    }} else {{
                        var input = document.createElement('input');
                        input.type = 'hidden';
                        input.id = 'g-recaptcha-response';
                        input.name = 'g-recaptcha-response';
                        input.value = '{token}';
                        document.body.appendChild(input);
                    }}
                """)
                debug_screenshot(driver, "after_token_input")
                time.sleep(2)

                # 4. Нажимаем кнопку через JS
                submit_btn = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "button[type='submit']"))
                )
                driver.execute_script("arguments[0].click();", submit_btn)
                print("Форма отправлена через JS")
                debug_screenshot(driver, "after_submit")
                time.sleep(3)

                return True

            elif status_response.get("errorId") != 0:
                error = status_response.get("errorDescription", "Неизвестная ошибка API")
                raise Exception(f"Ошибка API: {error}")

        raise Exception("Превышено время ожидания решения (5 минут)")

    except Exception as e:
        debug_screenshot(driver, "captcha_error")
        print(f"Ошибка при решении капчи: {str(e)}")
        return False


def handle_captcha(driver):
    """Полная обработка капчи с улучшенной логикой"""
    print("Обнаружена капча, начинаем обработку...")
    debug_screenshot(driver, "captcha_detected")

    try:
        # 1. Кликаем на чекбокс "Я не робот"
        checkbox_frame = WebDriverWait(driver, 20).until(
            EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe[title*='reCAPTCHA']"))
        )
        checkbox = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "recaptcha-checkbox"))
        )
        checkbox.click()
        print("Чекбокс 'Я не робот' нажат")
        driver.switch_to.default_content()
        debug_screenshot(driver, "after_checkbox_click")
        time.sleep(3)

        # 2. Решаем капчу через API
        if not solve_recaptcha_v2(driver):
            return False

        return True

    except Exception as e:
        debug_screenshot(driver, "captcha_handling_error")
        print(f"Ошибка при обработке капчи: {str(e)}")
        return False


def debug_screenshot(driver, name):
    """Сохранение скриншота для отладки"""
    if not os.path.exists('debug'):
        os.makedirs('debug')
    driver.save_screenshot(f'debug/{name}.png')


def apply_date_filters(driver, start_date, end_date):
    """Применение фильтров по дате регистрации с улучшенной обработкой"""
    logger.info(f"Применение фильтров: {start_date.strftime('%Y-%m-%d')} - {end_date.strftime('%Y-%m-%d')}")

    try:
        # Ожидаем загрузки страницы
        WebDriverWait(driver, PAGE_LOAD_TIMEOUT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "button[data-bs-target='#flush-collapse-1']"))
        )

        # Прокручиваем к верхней части страницы
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)

        # Находим кнопку "Дата регистрации"
        date_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-bs-target='#flush-collapse-1']"))
        )

        # Проверяем состояние кнопки (открыта/закрыта)
        is_collapsed = "collapsed" in date_button.get_attribute("class")

        # Если кнопка закрыта (collapsed), кликаем чтобы открыть
        if is_collapsed:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", date_button)
            date_button.click()
            time.sleep(2)

        # Ждем появления полей ввода дат
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "reg_date_from"))
        )

        # Очищаем поле "От" (3 разных способа на случай если один не сработает)
        date_from = driver.find_element(By.ID, "reg_date_from")
        driver.execute_script("arguments[0].removeAttribute('readonly')", date_from)
        date_from.clear()  # Способ 1: стандартный clear()
        date_from.send_keys(Keys.CONTROL + 'a')  # Способ 2: выделить все
        date_from.send_keys(Keys.DELETE)  # Способ 3: удалить
        time.sleep(1)

        # Очищаем поле "До" (аналогично)
        date_to = driver.find_element(By.ID, "reg_date_to")
        driver.execute_script("arguments[0].removeAttribute('readonly')", date_to)
        date_to.clear()
        date_to.send_keys(Keys.CONTROL + 'a')
        date_to.send_keys(Keys.DELETE)
        time.sleep(1)

        # Вводим новые даты
        date_from.send_keys(start_date.strftime("%Y-%m-%d"))
        time.sleep(1)
        date_to.send_keys(end_date.strftime("%Y-%m-%d"))
        time.sleep(1)

        # Прокручиваем к кнопке "Применить"
        apply_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//button[contains(@class, 'primary') and contains(., 'Применить')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", apply_button)
        time.sleep(1)

        # Кликаем кнопку "Применить"
        # После клика на кнопку "Применить"
        apply_button.click()
        time.sleep(6)

        # Проверяем капчу после применения фильтров
        if driver.find_elements(By.CSS_SELECTOR, "iframe[title*='reCAPTCHA']"):
            if not handle_captcha(driver):
                return False
        # Ждем загрузки результатов

        return True

    except Exception as e:
        logger.error(f"Ошибка при применении фильтров: {str(e)}")
        debug_screenshot(driver, "filter_error")
        return False


def get_all_company_links(driver):
    """Собираем ссылки на компании с учетом уже примененных фильтров"""
    all_links = []
    page_num = 1
    max_pages = 1  # Максимальное количество страниц
    processed_pages = set()

    while page_num <= max_pages:
        logger.info(f"Обработка страницы {page_num}")

        try:
            # Применяем фильтры для текущей страницы
            if page_num > 1:
                # В случае, если не первая страница, переходим на нужную страницу
                driver.get(f"{BASE_URL}?page={page_num}")
                time.sleep(2)

                # Проверяем наличие капчи
                if driver.find_elements(By.CSS_SELECTOR, "iframe[title*='reCAPTCHA']"):
                    if not handle_captcha(driver):
                        logger.error("Не удалось решить капчу при переходе на страницу")
                        break

            # Прокручиваем страницу до конца, чтобы кнопка "Далее" стала видимой
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)

            # Проверяем наличие капчи после прокрутки
            if driver.find_elements(By.CSS_SELECTOR, "iframe[title*='reCAPTCHA']"):
                if not handle_captcha(driver):
                    logger.error("Не удалось решить капчу после прокрутки")
                    break

            # Собираем все ссылки на компании на текущей странице
            soup = BeautifulSoup(driver.page_source, 'html.parser')

            # Проверяем наличие сообщения о том, что не найдено ни одного юридического лица
            no_results_message = soup.select_one("p.mt-4.text-center")
            if no_results_message and "Не найдено ни одного юридического лица" in no_results_message.text:
                logger.info("Не найдено ни одного юридического лица на последней странице.")
                break

            # Собираем ссылки на компании
            current_links = [f"https://checko.ru{a['href']}" for a in soup.select('a.link[href^="/company/"]')]

            # Проверяем новые ссылки и добавляем их
            new_links = [link for link in current_links if link not in all_links]
            all_links.extend(new_links)
            logger.info(f"Добавлено {len(new_links)} новых ссылок (Всего: {len(all_links)})")

            # Увеличиваем номер страницы для следующей итерации
            page_num += 1

        except Exception as e:
            logger.error(f"Ошибка на странице {page_num}: {str(e)}")
            debug_screenshot(driver, f"page_{page_num}_error")
            break

    logger.info(f"Сбор завершен. Всего ссылок: {len(all_links)}")
    return all_links



def get_person_info(soup, label):
    """Универсальная функция для поиска информации о директоре и учредителе"""
    try:
        # Определяем, что ищем (директора или учредителя)
        is_director = 'директор' in label.lower()

        # Поиск директора
        if is_director:
            director = None
            director_inn = None  # Добавляем переменную для ИНН директора

            # Находим секцию с директором
            director_section = soup.find('div', class_='fw-700', string=lambda t: t and 'директор' in t.lower())
            if not director_section:
                director_section = soup.find('strong', class_='fw-700', string=lambda t: t and 'директор' in t.lower())

            if director_section:
                # Ищем ссылку на имя директора
                director_tag = director_section.find_next('a', class_='link')
                if director_tag:
                    director = director_tag.get_text(strip=True)

                    # Находим ИНН рядом с директором
                    director_inn_tag = director_section.find_next('span', {'class': 'copy'})
                    if director_inn_tag:
                        director_inn = director_inn_tag.get_text(strip=True)

                else:
                    # Альтернативный вариант поиска, если структура отличается
                    parent_div = director_section.find_parent('div', class_='mb-3')
                    if parent_div:
                        director_tag = parent_div.find('a', class_='link')
                        if director_tag:
                            director = director_tag.get_text(strip=True)
                            director_inn_tag = parent_div.find('span', {'class': 'copy'})
                            if director_inn_tag:
                                director_inn = director_inn_tag.get_text(strip=True)

            return director, director_inn  # Возвращаем два значения: директор и ИНН

        # Поиск учредителя
        else:
            founder = None
            founder_inn = None

            # Попытка найти учредителя в секции с ID 'founders'
            founder_section = soup.find('section', id='founders')

            if founder_section:
                # Находим таблицу с учредителями
                founder_table = founder_section.find('table', class_='table table-md')

                if founder_table:
                    # Находим все строки в таблицы
                    rows = founder_table.find_all('tr')

                    if rows:
                        # Извлекаем первого учредителя (первую строку таблицы, пропуская заголовок)
                        first_row = rows[1]  # Пропускаем заголовок таблицы
                        columns = first_row.find_all('td')  # Получаем все столбцы в строке

                        if len(columns) >= 2:
                            # Извлекаем Ф. И. О. учредителя
                            founder_tag = columns[1].find('a')
                            if founder_tag:
                                founder = founder_tag.get_text(strip=True)

                                # Проверка на "Показать на карте"
                                if "Показать на карте" in founder:
                                    return '', ''  # Возвращаем пустые строки, если нашли "Показать на карте"

                                # Извлекаем ИНН учредителя
                                inn_div = columns[1].find_next('div')
                                if inn_div and "ИНН" in inn_div.text:
                                    founder_inn = inn_div.text.split()[-1]  # Получаем последний элемент, который будет ИНН
                            else:
                                logger.error("Не удалось найти имя учредителя в таблице.")
                        else:
                            logger.error("Не удалось найти достаточное количество столбцов для учредителя.")
                    else:
                        logger.error("В таблице учредителей нет строк.")
                else:
                    logger.error("Таблица учредителей не найдена в секции.")

            # Если учредитель не найден в секции, ищем его через стандартный поиск
            if not founder:
                # Стандартный способ поиска учредителя (как было раньше)
                founder_section = soup.find('strong', class_='fw-700', string='Учредитель')
                if not founder_section:
                    founder_section = soup.find('div', class_='fw-700', string='Учредитель')

                if founder_section:
                    # Ищем ссылку на учредителя рядом с заголовком
                    founder_tag = founder_section.find_next('a', class_='link')
                    if founder_tag:
                        founder = founder_tag.get_text(strip=True)
                    else:
                        # Если нет ссылки, проверяем структуру как в вашем примере
                        parent_div = founder_section.find_parent('div', class_='mb-3')
                        if parent_div:
                            # Проверяем, есть ли вложенные div (может быть адрес)
                            address_divs = parent_div.find_all('div', recursive=False)
                            if len(address_divs) > 0 and 'Субъект РФ' not in address_divs[0].get_text():
                                # Если это не адрес, то берем текст после заголовка
                                founder = parent_div.get_text(strip=True).replace('Учредитель', '').strip()
                            elif len(address_divs) > 0:
                                # Если это адрес, пропускаем
                                founder = None

            return founder, founder_inn  # Возвращаем Ф. И. О. и ИНН учредителя

    except Exception as e:
        logger.error(f"Ошибка при поиске {label}: {str(e)}")
        return '', ''  # Возвращаем два значения пустых строк для директора и ИНН


def get_first_okved(soup):
    """Функция для получения первого вида ОКВЭД"""
    try:
        x_section = soup.find('section', id='activity')
        if x_section:
            print('Секция найдена')
        # Находим таблицу с видами деятельности
            activity_table = x_section.find('table', class_='table table-sm table-striped')
            if not activity_table:
                logger.error("Таблица с видами деятельности не найдена.")
                return None, None

        # Находим все строки в таблице
            rows = activity_table.find_all('tr')

            if not rows:
                logger.error("В таблице нет строк с ОКВЭД.")
                return None, None

        # Извлекаем первый вид деятельности (первую строку таблицы)
            first_row = rows[0]  # Первая строка таблицы
            columns = first_row.find_all('td')  # Получаем все столбцы в строке

            okved_code = columns[0].text.strip()  # Код ОКВЭД (например, 01.21)
            activity_description = columns[1].text.strip()  # Описание (например, Выращивание винограда)

            return okved_code, activity_description
    except Exception as e:
        logger.error(f"Ошибка при извлечении ОКВЭД: {str(e)}")
        return None, None


def get_founder_inn(soup):
    """Функция для получения ИНН учредителя"""
    try:
        # Находим секцию с учредителями
        founder_section = soup.find('div', class_='tab-pane fade show active', id='founders-tab-1')

        if not founder_section:
            logger.error("Секция учредителей не найдена.")
            return None

        # Находим таблицу с учредителями
        founder_table = founder_section.find('table', class_='table table-md')

        if not founder_table:
            logger.error("Таблица с учредителями не найдена.")
            return None

        # Находим все строки в таблице
        rows = founder_table.find_all('tr')

        if not rows:
            logger.error("В таблице нет строк с учредителями.")
            return None

        # Извлекаем первого учредителя (первую строку таблицы)
        first_row = rows[1]  # Пропускаем заголовок таблицы, начинаем с первой строки с данными
        columns = first_row.find_all('td')  # Получаем все столбцы в строке

        if len(columns) < 2:
            logger.error("Неверная структура строки таблицы учредителей.")
            return None

        # Извлекаем ИНН учредителя, который содержится в следующем div
        inn_div = columns[1].find_next('div')
        if inn_div and "ИНН" in inn_div.text:
            founder_inn = inn_div.text.split()[-1]  # Получаем последний элемент, который будет ИНН
            return founder_inn
        else:
            logger.error("ИНН учредителя не найден.")
            return None
    except Exception as e:
        logger.error(f"Ошибка при извлечении ИНН учредителя: {str(e)}")
        return None


def parse_company_page(driver, url, existing_inns):
    """Парсинг данных компании с проверкой дубликатов по ИНН"""
    print(f"\nОбрабатываем компанию: {url}")
    try:
        driver.get(url)
        debug_screenshot(driver, f"company_page_{url.split('/')[-1]}")

        # Ожидаем либо данные, либо капчу
        WebDriverWait(driver, PAGE_LOAD_TIMEOUT).until(
            lambda d: d.find_elements(By.ID, "copy-inn") or
                      d.find_elements(By.CSS_SELECTOR, "iframe[title*='reCAPTCHA']")
        )

        # Если есть капча - обрабатываем
        if driver.find_elements(By.CSS_SELECTOR, "iframe[title*='reCAPTCHA']"):
            if not handle_captcha(driver):
                return None

        # Дожидаемся загрузки данных
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "copy-inn"))
        )

        # Парсинг данных
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # Прокручиваем страницу (может появиться капча)
        driver.execute_script("window.scrollTo(0, 5000);")
        time.sleep(1)

        # Проверяем капчу после прокрутки
        if driver.find_elements(By.CSS_SELECTOR, "iframe[title*='reCAPTCHA']"):
            if not handle_captcha(driver):
                return None

        # Основные данные
        inn = None
        try:
            inn = soup.find('strong', id='copy-inn').get_text(strip=True) if soup.find('strong',
                                                                                       id='copy-inn') else None
        except Exception as e:
            print(f"Ошибка при извлечении ИНН: {e}")

        # Проверка дубликата по ИНН
        if not inn:
            print("Пропускаем - нет ИНН")
            return None

        if inn in existing_inns:
            print(f"Пропускаем дубликат ИНН: {inn}")
            return None

        date = soup.find('div', string='Дата регистрации').find_next('div').get_text(strip=True) if soup.find('div',
                                                                                                              string='Дата регистрации') else None

        # Директор и учредитель
        director, director_inn = get_person_info(soup, 'Генеральный директор') or get_person_info(soup, 'Директор')
        founder, founder_inn = get_person_info(soup, 'Учредитель')

        # Телефоны
        phones = []
        phone_divs = soup.find_all('div', class_='col-12 col-lg-4')
        for div in phone_divs:
            if 'Телефон' in div.get_text():
                phone_links = div.find_all('a', href=lambda x: x and x.startswith('tel:'))
                for link in phone_links:
                    phone = link.get_text(strip=True)
                    if phone and phone not in phones:
                        phones.append(phone)

        phone = ', '.join(phones) if phones else None

        # Email
        email_tag = soup.find('a', href=lambda x: x and x.startswith('mailto:'))
        email = email_tag.get_text(strip=True) if email_tag else None

        # Получаем первый ОКВЭД
        okved_code, okved_description = get_first_okved(soup)

        # Извлекаем юридический адрес
        legal_address = None
        address_tag = soup.find('span', id='copy-address')
        if address_tag:
            legal_address = address_tag.get_text(strip=True)

        # Извлекаем уставной капитал
        charter_capital = None
        capital_tag = soup.find('div', string="Уставный капитал")
        if capital_tag:
            # Получаем следующий элемент, который содержит текст с уставным капиталом
            charter_capital = capital_tag.find_next('div').get_text(strip=True)

        # Проверяем обязательные поля
        if not inn:
            print("Пропускаем - нет ИНН")
            return None

        if not phone and not email:
            print("Пропускаем - нет ни телефона, ни email")
            return None

        # Формируем строку для таблицы
        current_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(
            f"Данные: ИНН={inn}, Дата={date}, Директор={director}, Учредитель={founder}, Телефон={phone}, Email={email}, ОКВЭД={okved_code} - {okved_description}, Юридический адрес={legal_address}, Уставной капитал={charter_capital}")

        return {
            'ИНН': inn,
            'Дата регистрации': date,
            'Ген. директор': director,
            'ИНН директора': director_inn,  # Добавляем ИНН директора
            'Учредитель': founder,
            'ИНН учредителя': founder_inn if founder_inn else '',  # Добавляем ИНН учредителя, если найден
            'Телефон': phone,
            'Email': email,
            'ОКВЭД': f"{okved_code} - {okved_description}",  # Добавляем ОКВЭД
            'Юридический адрес': legal_address,  # Добавляем юридический адрес
            'Уставной капитал': charter_capital,  # Добавляем уставной капитал
            'URL': url,
            'Дата добавления': current_date,
            'EmailSent': False  # Флаг отправки письма
        }

    except Exception as e:
        debug_screenshot(driver, f"parse_error_{url.split('/')[-1]}")
        print(f"Ошибка при парсинге компании: {str(e)}")
        return None


def save_to_excel(data, filepath):
    """Сохранение данных в Excel с проверкой дубликатов"""
    try:
        # Загрузка существующих данных, если файл уже есть
        if os.path.exists(filepath):
            existing_df = pd.read_excel(filepath)
            existing_inns = set(existing_df['ИНН'].dropna().astype(str))
        else:
            existing_df = pd.DataFrame()
            existing_inns = set()

        # Создаем DataFrame из новых данных
        new_df = pd.DataFrame(data)

        # Удаляем дубликаты среди новых данных
        new_df = new_df.drop_duplicates(subset=['ИНН'])

        # Фильтруем только новые записи
        new_df = new_df[~new_df['ИНН'].isin(existing_inns)]

        if new_df.empty:
            logger.info("Нет новых данных для сохранения")
            return

        # Объединяем с существующими данными
        final_df = pd.concat([existing_df, new_df], ignore_index=True)

        # Сохраняем результат
        final_df.to_excel(filepath, index=False)
        logger.info(f"Сохранено {len(new_df)} новых записей. Всего записей: {len(final_df)}")

    except Exception as e:
        logger.error(f"Ошибка при сохранении: {str(e)}")
        raise


def process_month(driver, start_date, end_date, existing_inns):
    """Обработка одного месяца"""
    month_name = start_date.strftime("%B %Y").lower()
    output_file = f"{month_name}.xlsx"
    all_data = []

    logger.info(f"\nНачинаем обработку месяца: {month_name}")

    # Переходим на страницу поиска
    driver.get(BASE_URL)
    time.sleep(3)

    # Применяем фильтры
    if not apply_date_filters(driver, start_date, end_date):
        return existing_inns, []

    # Собираем все ссылки на компании
    company_links = get_all_company_links(driver)
    logger.info(f"Найдено {len(company_links)} компаний за {month_name}")

    if not company_links:
        logger.info(f"Нет компаний за {month_name}, пропускаем")
        return existing_inns, []

    # Парсим данные компаний
    for i, link in enumerate(company_links, 1):
        company_data = parse_company_page(driver, link, existing_inns)
        if company_data:
            all_data.append(company_data)
            existing_inns.add(company_data['ИНН'])

        if i % 10 == 0:
            logger.info(f"Обработано {i}/{len(company_links)} компаний за {month_name}")

        time.sleep(random.uniform(1, 3))

    # Сохраняем данные
    if all_data:
        save_to_excel(all_data, output_file)
        logger.info(f"Сохранено {len(all_data)} компаний в файл {output_file}")
    else:
        logger.info(f"Нет новых компаний для сохранения за {month_name}")

    return existing_inns, all_data



def main():
    """Основная функция парсера"""
    driver = setup_driver()
    processed_count = 0
    emails_sent = 0
    all_inns = set()

    try:
        # Определяем месяцы для парсинга (с 1 января 2025 по 31 мая 2025)
        current_date = datetime(2025, 1, 1)
        end_date = datetime(2025, 5, 31)

        while current_date >= end_date:
            month_start = current_date.replace(day=1)
            month_end = (month_start + timedelta(days=32)).replace(day=1) - timedelta(days=1)

            # Обрабатываем месяц
            all_inns, month_data = process_month(driver, month_start, month_end, all_inns)
            processed_count += len(month_data)
            emails_sent += sum(1 for item in month_data if item['EmailSent'])

            # Переходим к предыдущему месяцу
            current_date = month_start - timedelta(days=1)

    finally:
        driver.quit()
        logger.info("Парсер завершил работу")
        logger.info(f"Обработано компаний: {processed_count}")
        logger.info(f"Отправлено писем: {emails_sent}")


if __name__ == "__main__":
    main()