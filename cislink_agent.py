"""
Агент синхронизации CISLink → ЛК PROTECO
Версия 2.0 - с парсингом детальных ошибок
"""

import os
import re
import json
import time
import logging
from datetime import datetime
from typing import Optional, Dict, List

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

    def parse_error_details(self, row_index: int) -> Optional[Dict]:
        """
        Кликает на ссылку Error в строке таблицы и парсит детали ошибки из popup.
        
        Args:
            row_index: Индекс строки в таблице (начиная с 0 для первой строки данных)
        
        Returns:
            Словарь с деталями ошибки или None
        """
        try:
            # Формируем ID ссылки Error для данной строки
            # Формат: ctl00_ContentPlaceHolder1_gvUploads_ctl{XX}_lnkView
            # где XX = row_index + 2 (начинается с ctl02)
            row_num = str(row_index + 2).zfill(2)
            link_id = f"ctl00_ContentPlaceHolder1_gvUploads_ctl{row_num}_lnkView"
            
            logger.info(f"Ищем ссылку Error с ID: {link_id}")
            
            try:
                error_link = self.driver.find_element(By.ID, link_id)
            except NoSuchElementException:
                # Пробуем альтернативный поиск по тексту "Error"
                logger.info("Поиск ссылки Error по тексту...")
                rows = self.driver.find_elements(By.CSS_SELECTOR, "table tr")
                if row_index + 1 < len(rows):
                    row = rows[row_index + 1]
                    try:
                        error_link = row.find_element(By.LINK_TEXT, "Error")
                    except NoSuchElementException:
                        logger.warning(f"Ссылка Error не найдена для строки {row_index}")
                        return None
                else:
                    return None
            
            # Кликаем на ссылку Error
            logger.info(f"Кликаем на ссылку Error...")
            self.driver.execute_script("arguments[0].click();", error_link)
            
            # Ждем появления popup с деталями ошибки
            time.sleep(2)
            
            # Ищем элемент с деталями ошибки
            details_element = None
            try:
                details_element = self.wait.until(
                    EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_lblDetails"))
                )
            except TimeoutException:
                # Пробуем альтернативные селекторы
                try:
                    details_element = self.driver.find_element(By.CSS_SELECTOR, ".pnlBackGround span")
                except NoSuchElementException:
                    pass
            
            if not details_element:
                logger.warning("Элемент с деталями ошибки не найден")
                return None
            
            # Получаем HTML содержимое
            details_html = details_element.get_attribute('innerHTML')
            details_text = details_element.text
            
            logger.info(f"Получен текст ошибки ({len(details_text)} символов)")
            
            # Парсим детали ошибки
            error_data = self._parse_error_html(details_html, details_text)
            
            # Закрываем popup (ищем кнопку "Закрыть")
            try:
                close_button = self.driver.find_element(By.CSS_SELECTOR, "input[value='Закрыть']")
                close_button.click()
                time.sleep(1)
            except NoSuchElementException:
                # Пробуем закрыть по Escape или кликом вне popup
                try:
                    self.driver.find_element(By.TAG_NAME, "body").send_keys("\x1b")
                    time.sleep(0.5)
                except:
                    pass
            
            return error_data
            
        except Exception as e:
            logger.error(f"Ошибка при парсинге деталей ошибки: {e}")
            return None

    def _parse_error_html(self, html: str, text: str) -> Dict:
        """
        Парсит HTML с деталями ошибки и извлекает структурированные данные.
        
        Returns:
            Словарь с ключами:
            - raw_text: полный текст ошибки
            - errors: список распознанных ошибок
        """
        result = {
            'raw_text': text.strip(),
            'errors': []
        }
        
        # Извлекаем заголовок ошибки (текст в <b>)
        header_match = re.search(r'<b>([^<]+)</b>', html)
        error_header = header_match.group(1).strip() if header_match else ''
        
        # Определяем шаг и тип ошибки из заголовка
        step_match = re.search(r'Шаг\s*(\d+)', error_header)
        step = step_match.group(0) if step_match else None
        
        # Определяем файл с ошибкой
        file_type = None
        if 'pdrest' in error_header.lower():
            file_type = 'pdrest'
        elif 'pdfact' in error_header.lower():
            file_type = 'pdfact'
        elif 'pdcatal' in error_header.lower():
            file_type = 'pdcatal'
        elif 'pdclient' in error_header.lower():
            file_type = 'pdclient'
        elif 'pddoc' in error_header.lower():
            file_type = 'pddoc'
        elif 'pdwh' in error_header.lower():
            file_type = 'pdwh'
        elif 'pdseria' in error_header.lower():
            file_type = 'pdseria'
        
        # Определяем затронутые поля
        affected_fields = []
        field_patterns = [
            r'поля?\s+(\w+(?:,\s*\w+)*)',
            r'in_qty|out_qty|beg_rest|end_rest|quantity|amount|vat',
            r'code|whcode|clientcode|doc_number|serial_no'
        ]
        for pattern in field_patterns:
            matches = re.findall(pattern, error_header, re.IGNORECASE)
            affected_fields.extend(matches)
        
        # Парсим таблицу с примерами ошибок
        examples = []
        table_match = re.search(r'<table[^>]*>(.*?)</table>', html, re.DOTALL | re.IGNORECASE)
        if table_match:
            table_html = table_match.group(1)
            rows = re.findall(r'<tr[^>]*>(.*?)</tr>', table_html, re.DOTALL | re.IGNORECASE)
            
            headers = []
            for i, row in enumerate(rows):
                cells = re.findall(r'<td[^>]*>(.*?)</td>', row, re.DOTALL | re.IGNORECASE)
                cells = [re.sub(r'<[^>]+>', '', cell).strip() for cell in cells]
                
                if i == 0:
                    # Первая строка - заголовки
                    headers = cells
                else:
                    # Остальные - данные
                    if cells and len(cells) == len(headers):
                        example = dict(zip(headers, cells))
                        examples.append(example)
        
        # Считаем количество ошибок
        error_count = len(examples) if examples else None
        
        # Проверяем наличие "(список неполный)"
        is_truncated = '(список неполный)' in text or '...' in text
        
        # Формируем структурированную ошибку
        error_entry = {
            'step': step,
            'file': file_type,
            'message': error_header,
            'fields': list(set(affected_fields)) if affected_fields else None,
            'count': error_count,
            'is_truncated': is_truncated,
            'examples': examples[:5] if examples else None  # Максимум 5 примеров
        }
        
        result['errors'].append(error_entry)
        
        return result

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
            
            for row_index, row in enumerate(rows):
                try:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) < 10:
                        continue
                    
                    upload_status_text = cells[1].text.strip().lower()
                    upload_status = 'error' if 'неудачн' in upload_status_text else ('success' if 'удачн' in upload_status_text else 'error')
                    stock_max_date = self.parse_date(cells[9].text)
                    distr_id = self.parse_int(cells[4].text)
                    
                    # Проверяем наличие ссылки Error в последней колонке
                    has_error_link = False
                    error_details = None
                    
                    if upload_status == 'error':
                        # Ищем ссылку Error в строке
                        try:
                            error_link = row.find_element(By.LINK_TEXT, "Error")
                            has_error_link = True
                        except NoSuchElementException:
                            # Пробуем найти по частичному тексту
                            try:
                                error_link = row.find_element(By.PARTIAL_LINK_TEXT, "Error")
                                has_error_link = True
                            except NoSuchElementException:
                                pass
                        
                        # Если есть ссылка Error, парсим детали
                        if has_error_link:
                            logger.info(f"Найдена ошибка для дистрибьютора {cells[5].text.strip()}, парсим детали...")
                            error_details = self.parse_error_details(row_index)
                    
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
                            'errors': error_details
                        }
                        reports.append(report)
                        
                except Exception as e:
                    logger.warning(f"Ошибка обработки строки {row_index}: {e}")
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
            response = requests.post(
                self.url, 
                json={'api_key': self.api_key, 'reports': reports}, 
                timeout=60
            )
            return response.json()
        except Exception as e:
            return {'success': False, 'error': str(e)}


def main():
    logger.info("=" * 50)
    logger.info("Агент CISLink v2.0 (с парсингом ошибок)")
    logger.info("=" * 50)
    
    if not all([CONFIG['cislink_login'], CONFIG['cislink_password'], CONFIG['api_url'], CONFIG['api_key']]):
        logger.error("Не заданы переменные окружения!")
        exit(1)
    
    scraper = CISLinkScraper()
    try:
        scraper.init_browser()
        
        if not scraper.login():
            logger.error("Не удалось авторизоваться")
            exit(1)
            
        if not scraper.navigate_to_reports():
            logger.error("Не удалось перейти к отчетам")
            exit(1)
        
        reports = scraper.scrape_reports()
        
        if reports:
            # Логируем статистику
            errors_count = sum(1 for r in reports if r['upload_status'] == 'error')
            with_details = sum(1 for r in reports if r.get('errors'))
            logger.info(f"Всего записей: {len(reports)}, с ошибками: {errors_count}, с деталями: {with_details}")
            
            # Отправляем в API
            result = APIClient().send_reports(reports)
            logger.info(f"Результат отправки: {result}")
            
            if not result.get('success'):
                exit(1)
        else:
            logger.warning("Нет данных для отправки")
            
    finally:
        scraper.close()
    
    logger.info("Синхронизация завершена успешно")


if __name__ == '__main__':
    main()
