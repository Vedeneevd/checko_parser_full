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
    page_num = 1
    max_pages = 999  # Максимальное количество страниц

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


def parse_company_page(driver, url, existing_inns):
    """Парсинг данных компании"""
    logger.info(f"Обработка компании: {url}")

    try:
        driver.get(url)
        WebDriverWait(driver, PAGE_LOAD_TIMEOUT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#copy-inn, .company-not-found"))
        )

        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # Проверка на 404
        if soup.select('.company-not-found'):
            return None

        # Основные данные
        inn = soup.find('strong', id='copy-inn').get_text(strip=True) if soup.find('strong', id='copy-inn') else None
        if not inn or inn in existing_inns:
            return None

        name = soup.select_one('h1.company-title').get_text(strip=True) if soup.select_one('h1.company-title') else None
        address = soup.select_one('div.company-address').get_text(strip=True) if soup.select_one(
            'div.company-address') else None
        reg_date = soup.find('div', string='Дата регистрации').find_next('div').get_text(strip=True) if soup.find('div',
                                                                                                                  string='Дата регистрации') else None

        # Телефоны и email
        phones = []
        phone_section = soup.find('strong', string='Телефоны')
        if phone_section:
            phone_div = phone_section.find_next('div')
            if phone_div:
                phones = [a.get_text(strip=True) for a in
                          phone_div.find_all('a', href=lambda x: x and x.startswith('tel:'))]

        email_tag = soup.find('a', href=lambda x: x and x.startswith('mailto:'))
        email = email_tag.get_text(strip=True) if email_tag else None

        return {
            'Название': name,
            'ИНН': inn,
            'Адрес': address,
            'Дата регистрации': reg_date,
            'Телефоны': ', '.join(phones) if phones else None,
            'Email': email,
            'URL': url
        }

    except Exception as e:
        logger.error(f"Ошибка при парсинге компании {url}: {str(e)}")
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

    return True


def main():
    """Основная функция парсера"""
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

    # Обрабатываем каждый месяц
    for month_start, month_end in months_to_parse:
        month_name = month_start.strftime("%B %Y").lower()
        logger.info(f"\nНачинаем обработку месяца: {month_name}")

        success = False
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                success = process_month(driver, month_start, month_end)
                if success:
                    break
            except Exception as e:
                logger.error(f"Попытка {attempt} не удалась: {str(e)}")
                time.sleep(10 * attempt)

        if not success:
            logger.error(f"Не удалось обработать месяц {month_name} после {MAX_RETRIES} попыток")

    driver.quit()
    logger.info("Парсер завершил работу")


if __name__ == "__main__":
    main()
