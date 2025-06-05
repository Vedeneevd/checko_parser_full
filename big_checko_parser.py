import time
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

# SMTPBZ_API_KEY = os.getenv('SMTPBZ_API_KEY')  # API ключ для smtp.bz
BASE_URL = "https://checko.ru/search/advanced"
START_PAGE = 1
END_PAGE = 10  # Всего 10 страниц
OUTPUT_FILE = "companies_data.xlsx"
PAGE_LOAD_TIMEOUT = 60
MAX_RETRIES = 100

def setup_driver():
    """Настройка веб-драйвера с уникальным каталогом данных"""
    options = webdriver.ChromeOptions()

    # Основные настройки
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    # Для работы на сервере
    # options.add_argument("--headless")
    # options.add_argument("--no-sandbox")
    # options.add_argument("--disable-dev-shm-usage")

    # Уникальный каталог данных для каждой сессии
    user_data_dir = f"/tmp/chrome_{int(time.time())}"
    options.add_argument(f"--user-data-dir={user_data_dir}")

    # Улучшенный User-Agent
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

    service = Service(ChromeDriverManager().install())
    try:
        driver = webdriver.Chrome(service=service, options=options)
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": """
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
            });
            """
        })
        return driver
    except Exception as e:
        print(f"Ошибка при создании драйвера: {str(e)}")
        raise



def search_filters_add(driver):
    url = driver.get(BASE_URL)
    time.sleep(5)
    driver.execute_script("window.scrollBy(0,5000)")
    time.sleep(2)
    actions = ActionChains(driver)
    contacts = driver.find_element(By.XPATH, "//*[@id='select-accordion']/div[9]")
    actions.move_to_element(contacts).perform()
    time.sleep(2)
    contacts_button = driver.find_element(By.XPATH, '//*[@id="flush-heading-9"]/button')
    if contacts_button:
        contacts_button.click()
        print('Кнопка контакты успешно нажата')
        time.sleep(2)
    else:
        print('Кнопки контакты не найдена')

    mobile_check_box = driver.find_element(By.XPATH, '//*[@id="flush-collapse-9"]/div/div[1]/div/div/div/div[1]')
    time.sleep(3)
    if mobile_check_box:
        mobile_check_box.click()
        print('Кнопка Указан телефон успешно нажата')
        time.sleep(3)
    else:
        print('Кнопка телефона не найдена')

    email_button = driver.find_element(By.XPATH, '//*[@id="flush-collapse-9"]/div/div[2]/div/div/div/div[1]')
    time.sleep(2)
    if email_button:
        email_button.click()
        print('Кнопка Указан email успешно нажата')
        time.sleep(1)
    else:
        print('Кнопка Указан email не найдена')
    time.sleep(2)
    submit_filter_button = driver.find_element(By.XPATH,'//*[@id="vue"]/div/button[1]')
    actions.move_to_element(submit_filter_button).perform()
    if submit_filter_button:
        submit_filter_button.click()
        print('Фильтры применены')
    else:
        print('Фильтры не применены')
    time.sleep(10)
    return driver


def main():
    """Основная функция сбора данных с немедленной отправкой писем"""
    print(f"\n=== Запуск парсера {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
    driver = setup_driver()
    add_filter = search_filters_add(driver)
    print('Фильтры успешно применены')

if __name__ == "__main__":
    main()

