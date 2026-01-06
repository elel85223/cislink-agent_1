"""
Агент синхронизации CISLink → ЛК PROTECO
Версия 1.5 - с парсингом детальных ошибок
"""

import os
import re
import json
import time
import logging
from datetime import datetime
from typing import Optional, Dict, List, Any

import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
    ElementClickInterceptedException
)
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
    'timeout': 30,
    'max_error_text_length': 2000,
    'max_error_examples': 5
}

# ID элементов CISLink
SELECTORS = {
    'table': 'ctl00_ContentPlaceHolder1_gvUploads',
    'error_link_template': 'ctl00_ContentPlaceHolder1_gvUploads_ctl{row_id}_lnkView',
    'error_details': 'ctl00_ContentPlaceHolder1_lblDetails',
    'close_button_xpath': "//input[@type='button' and @value='Закрыть']"
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

    def reload_reports_page(self) -> bool:
        """Перезагрузить страницу отчётов для избежания StaleElement"""
        try:
            self.driver.get(f"{CONFIG['cislink_url']}/Reports/UploadHistory.aspx")
            time.sleep(3)
            self.wait.until(EC.presence_of_element_located((By.ID, SELECTORS['table'])))
            return True
        except Exception as e:
            logger.error(f"Ошибка перезагрузки страницы: {e}")
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

    def parse_error_structure(self, raw_text: str, raw_html: str) -> Dict[str, Any]:
        """Парсинг структурированных данных из текста ошибки"""
        result = {
            'raw_text': raw_text[:CONFIG['max_error_text_length']] if raw_text else '',
            'errors': []
        }

        if not raw_text:
            return result

        # Паттерн для извлечения шага и файла
        step_pattern = r'Шаг\s+(\d+):\s*(.+?)(?=\(pd|$)'
        file_pattern = r'\((pd\w+)(?:\.txt|\.dbf)?\)'

        step_match = re.search(step_pattern, raw_text)
        file_matches = re.findall(file_pattern, raw_text)

        error_info = {
            'step': f"Шаг {step_match.group(1)}" if step_match else None,
            'file': file_matches[0] if file_matches else None,
            'message': step_match.group(2).strip() if step_match else raw_text[:200],
            'fields': [],
            'count': 0,
            'is_truncated': '(список неполный)' in raw_text or '...' in raw_text,
            'examples': []
        }

        # Парсинг таблицы с примерами из HTML
        if raw_html and '<table' in raw_html:
            try:
                # Извлекаем заголовки таблицы
                header_pattern = r'<tr[^>]*>\s*<td[^>]*>([^<]+)</td>'
                headers = re.findall(r'<td[^>]*>([^<]+)</td>', raw_html.split('</tr>')[0] if '</tr>' in raw_html else '')

                if headers:
                    error_info['fields'] = [h.strip() for h in headers if h.strip()]

                # Извлекаем строки данных
                rows = raw_html.split('</tr>')[1:]  # Пропускаем заголовок
                example_count = 0

                for row in rows:
                    if example_count >= CONFIG['max_error_examples']:
                        break
                    cells = re.findall(r'<td[^>]*>([^<]*)</td>', row)
                    if cells and len(cells) == len(error_info['fields']):
                        example = {}
                        for i, field in enumerate(error_info['fields']):
                            example[field] = cells[i].strip()
                        if any(example.values()):
                            error_info['examples'].append(example)
                            example_count += 1

                error_info['count'] = len(rows) - 1  # Примерное количество ошибок

            except Exception as e:
                logger.debug(f"Ошибка парсинга таблицы: {e}")

        result['errors'].append(error_info)
        return result

    def fetch_error_details(self, row_index: int) -> Optional[Dict[str, Any]]:
        """Получить детали ошибки для конкретной строки"""
        try:
            # Формируем ID ссылки на ошибку (row_index + 2, с ведущим нулём)
            row_id = str(row_index + 2).zfill(2)
            error_link_id = SELECTORS['error_link_template'].format(row_id=row_id)

            logger.debug(f"Ищем ссылку Error с ID: {error_link_id}")

            # Пробуем найти ссылку Error несколькими способами
            error_link = None

            # Способ 1: По точному ID
            try:
                error_link = self.driver.find_element(By.ID, error_link_id)
            except NoSuchElementException:
                pass

            # Способ 2: По частичному ID в строке таблицы
            if not error_link:
                try:
                    table = self.driver.find_element(By.ID, SELECTORS['table'])
                    rows = table.find_elements(By.TAG_NAME, "tr")
                    if row_index + 1 < len(rows):
                        row = rows[row_index + 1]  # +1 потому что первая строка - заголовок
                        links = row.find_elements(By.TAG_NAME, "a")
                        for link in links:
                            if 'lnkView' in (link.get_attribute('id') or ''):
                                error_link = link
                                break
                            if link.text.strip().lower() == 'error':
                                error_link = link
                                break
                except Exception:
                    pass

            if not error_link:
                logger.debug(f"Ссылка Error не найдена для строки {row_index}")
                return None

            # Кликаем на ссылку Error
            try:
                self.driver.execute_script("arguments[0].scrollIntoView(true);", error_link)
                time.sleep(0.5)
                error_link.click()
            except ElementClickInterceptedException:
                self.driver.execute_script("arguments[0].click();", error_link)

            time.sleep(2)

            # Ждём появления popup с деталями
            try:
                details_element = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, SELECTORS['error_details']))
                )

                raw_text = details_element.text.strip()
                raw_html = details_element.get_attribute('innerHTML')

                logger.info(f"Получены детали ошибки для строки {row_index}: {raw_text[:100]}...")

                # Закрываем popup
                self.close_error_popup()

                # Парсим структурированные данные
                return self.parse_error_structure(raw_text, raw_html)

            except TimeoutException:
                logger.warning(f"Popup с ошибкой не появился для строки {row_index}")
                self.close_error_popup()
                return None

        except Exception as e:
            logger.error(f"Ошибка получения деталей для строки {row_index}: {e}")
            self.close_error_popup()
            return None

    def close_error_popup(self):
        """Закрыть popup с ошибкой"""
        try:
            # Способ 1: Кнопка "Закрыть"
            close_button = self.driver.find_element(By.XPATH, SELECTORS['close_button_xpath'])
            close_button.click()
            time.sleep(1)
            return
        except NoSuchElementException:
            pass

        try:
            # Способ 2: Любая кнопка с текстом "Закрыть" или "Close"
            buttons = self.driver.find_elements(By.TAG_NAME, "input")
            for btn in buttons:
                if btn.get_attribute('value') in ['Закрыть', 'Close', 'OK']:
                    btn.click()
                    time.sleep(1)
                    return
        except Exception:
            pass

        try:
            # Способ 3: Нажатие Escape
            from selenium.webdriver.common.keys import Keys
            self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
            time.sleep(1)
        except Exception:
            pass

    def scrape_reports(self) -> list:
        """
        Двухпроходный алгоритм сбора данных:
        1. Первый проход: собираем базовые данные и запоминаем строки с ошибками
        2. Второй проход: для каждой строки с ошибкой перезагружаем страницу и парсим popup
        """
        logger.info("Сбор данных из таблицы...")
        reports = []
        error_rows = []  # Индексы строк с ошибками для второго прохода

        try:
            # === ПЕРВЫЙ ПРОХОД: сбор базовых данных ===
            time.sleep(2)

            try:
                main_table = self.driver.find_element(By.ID, SELECTORS['table'])
            except NoSuchElementException:
                # Fallback: ищем самую большую таблицу
                tables = self.driver.find_elements(By.TAG_NAME, "table")
                main_table = max(tables, key=lambda t: len(t.find_elements(By.TAG_NAME, "tr")), default=None)

            if not main_table:
                logger.error("Таблица отчётов не найдена")
                return []

            rows = main_table.find_elements(By.TAG_NAME, "tr")[1:]  # Пропускаем заголовок
            logger.info(f"Найдено {len(rows)} строк в таблице")

            for row_index, row in enumerate(rows):
                try:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) < 10:
                        continue

                    upload_status_text = cells[1].text.strip().lower()
                    is_error = 'неудачн' in upload_status_text
                    upload_status = 'error' if is_error else ('success' if 'удачн' in upload_status_text else 'error')

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

                        # Проверяем наличие ссылки Error
                        if is_error:
                            try:
                                links = row.find_elements(By.TAG_NAME, "a")
                                has_error_link = any(
                                    'lnkView' in (link.get_attribute('id') or '') or
                                    link.text.strip().lower() == 'error'
                                    for link in links
                                )
                                if has_error_link:
                                    error_rows.append((len(reports) - 1, row_index))
                            except Exception:
                                pass

                except Exception as e:
                    logger.debug(f"Ошибка обработки строки {row_index}: {e}")
                    continue

            logger.info(f"Первый проход: собрано {len(reports)} записей, {len(error_rows)} с ошибками")

            # === ВТОРОЙ ПРОХОД: парсинг детальных ошибок ===
            if error_rows:
                logger.info(f"Второй проход: парсинг {len(error_rows)} ошибок...")

                for report_index, row_index in error_rows:
                    try:
                        # Перезагружаем страницу перед каждым парсингом ошибки
                        if not self.reload_reports_page():
                            logger.warning(f"Не удалось перезагрузить страницу для строки {row_index}")
                            continue

                        error_details = self.fetch_error_details(row_index)

                        if error_details:
                            reports[report_index]['errors'] = error_details
                            logger.info(f"Ошибка для {reports[report_index]['distr_name']}: получена")
                        else:
                            logger.debug(f"Детали ошибки не найдены для строки {row_index}")

                    except Exception as e:
                        logger.error(f"Ошибка парсинга деталей для строки {row_index}: {e}")
                        continue

                logger.info(f"Второй проход завершён")

            logger.info(f"Всего собрано {len(reports)} записей")

        except Exception as e:
            logger.error(f"Ошибка сбора данных: {e}")

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
            response = requests.post(
                self.url,
                json={'api_key': self.api_key, 'reports': reports},
                timeout=60
            )
            return response.json()
        except Exception as e:
            return {'success': False, 'error': str(e)}


def main():
    logger.info("Агент CISLink v1.5 (с парсингом ошибок)")

    if not all([CONFIG['cislink_login'], CONFIG['cislink_password'], CONFIG['api_url'], CONFIG['api_key']]):
        logger.error("Не заданы переменные окружения!")
        exit(1)

    scraper = CISLinkScraper()
    try:
        scraper.init_browser()

        if not scraper.login():
            logger.error("Авторизация не удалась")
            exit(1)

        if not scraper.navigate_to_reports():
            logger.error("Навигация не удалась")
            exit(1)

        reports = scraper.scrape_reports()

        if reports:
            # Подсчитываем статистику
            with_errors = sum(1 for r in reports if r.get('errors'))
            logger.info(f"Отчётов с детальными ошибками: {with_errors}")

            result = APIClient().send_reports(reports)
            logger.info(f"Результат отправки: {result}")

            if not result.get('success'):
                exit(1)
        else:
            logger.warning("Нет данных для отправки")

    finally:
        scraper.close()


if __name__ == '__main__':
    main()
