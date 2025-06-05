import json
import os
import random
import time
from datetime import datetime, timedelta
import pandas as pd
import requests
from selenium import webdriver
from selenium.webdriver import ActionChains
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



def setup_driver():
    """Настройка веб-драйвера"""
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
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

        # Кликаем на кнопку "Дата регистрации"
        date_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-bs-target='#flush-collapse-1']"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", date_button)
        date_button.click()
        time.sleep(2)

        # Ждем появления полей ввода дат
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "reg_date_from"))
        )

        # Вводим дату "От"
        date_from = driver.find_element(By.ID, "reg_date_from")
        driver.execute_script("arguments[0].removeAttribute('readonly')", date_from)
        date_from.clear()
        date_from.send_keys(start_date.strftime("%Y-%m-%d"))
        time.sleep(1)

        # Вводим дату "До"
        date_to = driver.find_element(By.ID, "reg_date_to")
        driver.execute_script("arguments[0].removeAttribute('readonly')", date_to)
        date_to.clear()
        date_to.send_keys(end_date.strftime("%Y-%m-%d"))
        time.sleep(1)

        # Прокручиваем к кнопке "Применить"
        apply_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'primary') and contains(., 'Применить')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", apply_button)
        time.sleep(1)

        # Кликаем кнопку "Применить"
        apply_button.click()
        time.sleep(6)

        # Ждем загрузки результатов

        return True

    except Exception as e:
        logger.error(f"Ошибка при применении фильтров: {str(e)}")
        debug_screenshot(driver, "filter_error")
        return False


def get_all_company_links(driver, start_date, end_date):
    """Собираем ссылки на компании с учетом уже примененных фильтров"""
    all_links = []
    page_num = 558
    max_pages = 560  # Максимальное количество страниц

    while page_num <= max_pages:
        logger.info(f"Обработка страницы {page_num}")

        try:
            # Применяем фильтры для текущей страницы
            if page_num > 1:
                # В случае, если не первая страница, переходим на нужную страницу
                driver.get(f"{BASE_URL}?page={page_num}")
                time.sleep(2)

            # Прокручиваем страницу до конца, чтобы кнопка "Далее" стала видимой
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)

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
    person_section = soup.find('strong', class_='fw-700', string=label)
    if not person_section:
        person_section = soup.find('div', class_='fw-700', string=label)
    if person_section:
        person_tag = person_section.find_next('a', class_='link')
        if person_tag:
            return person_tag.get_text(strip=True)
        parent_div = person_section.find_parent('div', class_='mb-3')
        if parent_div:
            return parent_div.get_text(strip=True).replace(label, '').strip()
    return None


def parse_company_page(driver, url, existing_inns):
    """Парсинг данных компании с проверкой дубликатов по ИНН и немедленной отправкой письма"""
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

        driver.execute_script("window.scrollTo(0, 5000);")

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
        director = get_person_info(soup, 'Директор')
        founder = get_person_info(soup, 'Учредитель')

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
            f"Данные: ИНН={inn}, Дата={date}, Директор={director}, Учредитель={founder}, Телефон={phone}, Email={email}")

        return [inn, date, director, founder, phone, email, url, current_date]


    except Exception as e:
        debug_screenshot(driver, f"parse_error_{url.split('/')[-1]}")
        print(f"Ошибка при парсинге компании: {str(e)}")
        return None



def process_month(driver, start_date, end_date):
    """Обработка одного месяца с оптимизацией применения фильтров"""
    month_name = start_date.strftime("%B %Y").lower()
    output_file = f"{month_name}.xlsx"
    existing_inns = set()

    # Загружаем существующие данные, если файл уже есть
    if os.path.exists(output_file):
        try:
            existing_df = pd.read_excel(output_file)
            existing_inns = set(existing_df['ИНН'].dropna().astype(str))
            logger.info(f"Загружено {len(existing_inns)} существующих ИНН из файла {output_file}")
        except Exception as e:
            logger.error(f"Ошибка при загрузке файла {output_file}: {str(e)}")

    # Переходим на страницу поиска
    driver.get(BASE_URL)
    time.sleep(3)

    # Применяем фильтры
    if not apply_date_filters(driver, start_date, end_date):
        return False

    # Собираем все ссылки на компании
    company_links = get_all_company_links(driver, start_date, end_date)
    logger.info(f"Найдено {len(company_links)} компаний за {month_name}")

    if not company_links:
        logger.info(f"Нет компаний за {month_name}, пропускаем")
        return True

    # Парсим данные компаний
    companies_data = []
    for i, link in enumerate(company_links, 1):
        company_data = parse_company_page(driver, link, existing_inns)
        if company_data:
            companies_data.append(company_data)
            existing_inns.add(company_data['ИНН'])

        if i % 10 == 0:
            logger.info(f"Обработано {i}/{len(company_links)} компаний за {month_name}")

        time.sleep(random.uniform(1, 3))

    # Сохраняем данные
    if companies_data:
        df = pd.DataFrame(companies_data)
        df.to_excel(output_file, index=False)
        logger.info(f"Сохранено {len(df)} компаний в файл {output_file}")
    else:
        logger.info(f"Нет новых компаний для сохранения за {month_name}")

    return all_links

def load_existing_data(filepath):
    """Загрузка существующих данных из файла"""
    if os.path.exists(filepath):
        try:
            df = pd.read_excel(filepath)
            return df
        except Exception as e:
            print(f"Ошибка при загрузке файла: {str(e)}")
            return pd.DataFrame()
    return pd.DataFrame()



def save_to_excel(new_data, filepath):
    """Сохранение данных с надежной проверкой дубликатов"""
    try:
        # Загрузка существующих данных
        existing_df = load_existing_data(filepath)

        # Получаем список существующих ИНН для проверки дубликатов
        existing_inns = set(existing_df['ИНН'].dropna().unique()) if not existing_df.empty else set()

        # Создание DataFrame из новых данных
        new_df = pd.DataFrame(new_data,
                              columns=['ИНН', 'Дата регистрации', 'Ген. директор', 'Учредитель',
                                       'Телефон', 'Email', 'URL', 'Дата добавления', 'EmailSent'])

        # Удаление полностью пустых строк
        new_df = new_df.dropna(how='all')

        # Фильтрация только компаний с телефоном или email
        new_df = new_df[(new_df['Телефон'].notna()) | (new_df['Email'].notna())]

        # Удаление дубликатов среди новых данных
        new_df = new_df.drop_duplicates(subset=['ИНН', 'URL'])

        # Удаление записей, которые уже есть в существующих данных
        new_df = new_df[~new_df['ИНН'].isin(existing_inns)]

        if new_df.empty:
            logger.info("Нет новых данных для сохранения")
            return

        # Объединение с существующими данными
        final_df = pd.concat([existing_df, new_df], ignore_index=True)

        # Дополнительная проверка на дубликаты после объединения
        final_df = final_df.drop_duplicates(subset=['ИНН', 'URL'], keep='last')

        # Сохранение результата
        with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, index=False)

            # Форматирование
            worksheet = writer.sheets['Sheet1']
            worksheet.set_column('A:A', 15)  # ИНН
            worksheet.set_column('B:B', 15)  # Дата регистрации
            worksheet.set_column('C:C', 25)  # Ген. директор
            worksheet.set_column('D:D', 25)  # Учредитель
            worksheet.set_column('E:E', 20)  # Телефон
            worksheet.set_column('F:F', 25)  # Email
            worksheet.set_column('G:G', 40)  # URL
            worksheet.set_column('H:H', 20)  # Дата добавления
            worksheet.set_column('I:I', 20)  # EmailSent

        logger.info(f"Сохранено {len(new_df)} новых записей. Всего записей: {len(final_df)}")

    except Exception as e:
        logger.error(f"Ошибка при сохранении: {str(e)}")
        raise



def main():
    """Основная функция парсера с сохранением данных в Excel и переходом на начальную страницу после обработки месяца"""
    driver = setup_driver()

    # Определяем месяцы для парсинга (с мая 2025 по январь 2025)
    months_to_parse = []
    current_date = datetime(2025, 5, 1)
    end_date = datetime(2025, 1, 1)

    while current_date >= end_date:
        month_start = current_date.replace(day=1)
        month_end = (month_start + timedelta(days=32)).replace(day=1) - timedelta(days=1)
        months_to_parse.append((month_start, month_end))
        current_date = month_start - timedelta(days=1)

    all_data = []  # Список для хранения всех данных за месяц
    processed_count = 0
    emails_sent = 0

    # Обрабатываем каждый месяц
    for month_start, month_end in months_to_parse:
        month_name = month_start.strftime("%B %Y").lower()
        logger.info(f"\nНачинаем обработку месяца: {month_name}")

        success = False
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                # Обработка месяца
                company_links = process_month(driver, month_start, month_end)
                if company_links:
                    # Фильтруем новые ссылки, которые не были обработаны
                    for i, link in enumerate(company_links, 1):
                        print(f"Обработка компании {i}/{len(company_links)}: {link}")
                        data = parse_company_page(driver, link, processed_count)

                        if data:
                            all_data.append({
                                'ИНН': data[0],
                                'Дата регистрации': data[1],
                                'Ген. директор': data[2],
                                'Учредитель': data[3],
                                'Телефон': data[4],
                                'Email': data[5],
                                'URL': data[6],
                                'Дата добавления': data[7],
                                'EmailSent': data[8]  # Дата отправки письма (если было отправлено)
                            })
                            processed_count += 1
                            if data[8]:  # Если письмо было отправлено
                                emails_sent += 1

                        # Промежуточное сохранение каждые 20 компаний
                        if i % 20 == 0 and all_data:
                            print(f"\nПромежуточное сохранение после {i} компаний...")
                            save_to_excel(all_data, f"{month_name}_output.xlsx")
                            all_data = []  # Очищаем после сохранения

                        time.sleep(random.uniform(2, 5))

                    # Сохраняем все данные за месяц в Excel
                    if all_data:
                        print("\nФинальное сохранение результатов...")
                        save_to_excel(all_data, f"{month_name}_output.xlsx")
                        all_data = []  # Очищаем после сохранения

                # После завершения обработки месяца, возвращаемся на начальную страницу
                logger.info(f"Завершена обработка месяца {month_name}. Переход на начальную страницу.")
                driver.get(BASE_URL)  # Переход на начальную страницу
                time.sleep(3)  # Ждем немного, чтобы страница загрузилась

                success = True
                break
            except Exception as e:
                logger.error(f"Попытка {attempt} не удалась: {str(e)}")
                time.sleep(10 * attempt)

        if not success:
            logger.error(f"Не удалось обработать месяц {month_name} после {MAX_RETRIES} попыток")

    driver.quit()
    logger.info("Парсер завершил работу")
    logger.info(f"Обработано компаний: {processed_count}")
    logger.info(f"Отправлено писем: {emails_sent}")

if __name__ == "__main__":
    main()
