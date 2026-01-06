"""
Агент синхронизации CISLink → ЛК PROTECO
Версия 1.4 - для GitHub Actions
"""

import os
import re
import json
import time
import logging
from datetime import datetime
from typing import Optional

import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

CONFIG = {
    'cislink_url': 'https://b2b.cislinkdts.com',
    'cislink_login': os.getenv('CISLINK_LOGIN'),
    'cislink_password': os.getenv('CISLINK_PASSWORD'),
    'api_url': os.getenv('API_URL'),
    'api_key': os.getenv('API_KEY'),
    'debug_mode': os.getenv('DEBUG_MODE', 'False').lower() == 'true',
    'timeout': 30
}

class CISLinkScraper:
    def __init__(self):
        self.driver = None
        self.wait = None

    def init_browser(self):
        logger.info("Инициализация браузера...")
        options = Options()
        if not CONFIG['debug_mode']:
            options.add_argument('--headless=new')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option('excludeSwitches', ['enable-automation'])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument('--lang=ru-RU')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-infobars')
        options.add_argument('--remote-debugging-port=9222')

        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        self.wait = WebDriverWait(self.driver, CONFIG['timeout'])
        logger.info("Браузер запущен")

    def login(self) -> bool:
        logger.info("Авторизация в CISLink...")
        try:
            self.driver.get(CONFIG['cislink_url'])
            time.sleep(3)
            login_field = self.wait.until(EC.presence_of_element_located((By.ID, "txtLogin")))
            password_field = self.driver.find_element(By.ID, "txtPassword")
            login_field.clear()
            time.sleep(0.5)
            login_field.send_keys(CONFIG['cislink_login'])
            password_field.clear()
            time.sleep(0.5)
            password_field.send_keys(CONFIG['cislink_password'])
            login_button = self.driver.find_element(By.ID, "btnEnter")
            login_button.click()
            time.sleep(5)
            current_url = self.driver.current_url
            logger.info(f"URL после входа: {current_url}")
            if "Default.aspx" in current_url or "Dictionary" in current_url:
                logger.info("Авторизация успешна!")
                return True
            return False
        except Exception as e:
            logger.error(f"Ошибка авторизации: {e}")
            return False

    def navigate_to_reports(self) -> bool:
        logger.info("Переход на страницу отчетов...")
        try:
            self.driver.get(f"{CONFIG['cislink_url']}/Dictionary/Default.aspx")
            time.sleep(3)
            try:
                select_all = self.driver.find_element(By.ID, "cbDistrs")
                if not select_all.is_selected():
                    select_all.click()
                    time.sleep(1)
            except NoSuchElementException:
                pass
            self.driver.get(f"{CONFIG['cislink_url']}/Reports/UploadHistory.aspx")
            time.sleep(3)
            return True
        except Exception as e:
            logger.error(f"Ошибка навигации: {e}")
            return False

    def parse_date(self, date_str: str) -> Optional[str]:
        if not date_str or date_str.strip() == '':
            return None
        try:
            date_str = date_str.strip()
            if ' ' in date_str:
                dt = datetime.strptime(date_str, '%d.%m.%Y %H:%M')
                return dt.strftime('%Y-%m-%d %H:%M:%S')
            else:
                dt = datetime.strptime(date_str, '%d.%m.%Y')
                return dt.strftime('%Y-%m-%d')
        except ValueError:
            return None

    def parse_int(self, value: str) -> Optional[int]:
        if not value or value.strip() == '':
            return None
        try:
            return int(value.strip())
        except ValueError:
            return None

    def scrape_reports(self) -> list:
        logger.info("Сбор данных из таблицы...")
        reports = []
        try:
            time.sleep(2)
            tables = self.driver.find_elements(By.TAG_NAME, "table")
            main_table = max(tables, key=lambda t: len(t.find_elements(By.TAG_NAME, "tr")), default=None)
            if not main_table:
                return []
            rows = main_table.find_elements(By.TAG_NAME, "tr")[1:]
            for row in rows:
                try:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) < 10:
                        continue
                    upload_status_text = cells[1].text.strip().lower()
                    upload_status = 'error' if 'неудачн' in upload_status_text else ('success' if 'удачн' in upload_status_text else 'error')
                    stock_max_date = self.parse_date(cells[9].text)
                    distr_id = self.parse_int(cells[4].text)
                    if distr_id:
                        report = {
                            'distr_id': distr_id,
                            'distr_code': cells[3].text.strip(),
                            'distr_name': cells[5].text.strip(),
                            'city': cells[6].text.strip(),
                            'upload_datetime': self.parse_date(cells[0].text),
                            'upload_status': 'error' if upload_status == 'success' and not stock_max_date else upload_status,
                            'connection_type': cells[11].text.strip() if len(cells) > 11 else '',
                            'error_file_type': cells[2].text.strip() if upload_status == 'error' else '',
                            'doc_max_date': self.parse_date(cells[7].text),
                            'doc_period': self.parse_int(cells[8].text),
                            'stock_max_date': stock_max_date,
                            'stock_period': self.parse_int(cells[10].text) if len(cells) > 10 else None,
                            'errors': None
                        }
                        reports.append(report)
                except:
                    continue
            logger.info(f"Собрано {len(reports)} записей")
        except Exception as e:
            logger.error(f"Ошибка сбора: {e}")
        return reports

    def close(self):
        if self.driver:
            self.driver.quit()

class APIClient:
    def __init__(self):
        self.url = CONFIG['api_url']
        self.api_key = CONFIG['api_key']

    def send_reports(self, reports: list) -> dict:
        try:
            response = requests.post(self.url, json={'api_key': self.api_key, 'reports': reports}, timeout=60)
            return response.json()
        except Exception as e:
            return {'success': False, 'error': str(e)}

def main():
    logger.info("Агент CISLink v1.4")
    if not all([CONFIG['cislink_login'], CONFIG['cislink_password'], CONFIG['api_url'], CONFIG['api_key']]):
        logger.error("Не заданы переменные окружения!")
        exit(1)
    scraper = CISLinkScraper()
    try:
        scraper.init_browser()
        if not scraper.login() or not scraper.navigate_to_reports():
            exit(1)
        reports = scraper.scrape_reports()
        if reports:
            result = APIClient().send_reports(reports)
            logger.info(f"Результат: {result}")
            if not result.get('success'):
                exit(1)
    finally:
        scraper.close()

if __name__ == '__main__':
    main()
