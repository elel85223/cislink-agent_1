#!/usr/bin/env python3
"""
CISLink Agent v2.1 - Синхронизация данных дистрибьюторов
С исправленным парсингом ошибок и обработкой stale elements
"""

import os
import re
import json
import time
import logging
import requests
from datetime import datetime
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

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# Конфигурация
CISLINK_URL = "https://b2b.cislinkdts.com"
CISLINK_LOGIN = os.environ.get("CISLINK_LOGIN")
CISLINK_PASSWORD = os.environ.get("CISLINK_PASSWORD")
API_URL = os.environ.get("API_URL")
API_KEY = os.environ.get("API_KEY")


class CISLinkAgent:
    """Агент для сбора данных из CISLink"""
    
    def __init__(self):
        self.driver = None
        self.wait = None
        
    def setup_browser(self):
        """Настройка браузера"""
        logger.info("Инициализация браузера...")
        
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, 15)
        
        logger.info("Браузер запущен")
        
    def login(self):
        """Авторизация в CISLink"""
        logger.info("Авторизация в CISLink...")
        
        self.driver.get(f"{CISLINK_URL}/Account/Login.aspx")
        time.sleep(2)
        
        # Ввод логина
        login_field = self.wait.until(
            EC.presence_of_element_located((By.ID, "MainContent_LoginUser_UserName"))
        )
        login_field.clear()
        login_field.send_keys(CISLINK_LOGIN)
        
        # Ввод пароля
        password_field = self.driver.find_element(By.ID, "MainContent_LoginUser_Password")
        password_field.clear()
        password_field.send_keys(CISLINK_PASSWORD)
        
        # Клик на кнопку входа
        login_button = self.driver.find_element(By.ID, "MainContent_LoginUser_LoginButton")
        login_button.click()
        
        time.sleep(3)
        
        # Проверка успешной авторизации
        current_url = self.driver.current_url
        logger.info(f"URL после входа: {current_url}")
        
        if "Login" in current_url:
            raise Exception("Не удалось авторизоваться")
            
        logger.info("Авторизация успешна!")
        
    def navigate_to_reports(self):
        """Переход на страницу отчетов"""
        logger.info("Переход на страницу отчетов...")
        
        self.driver.get(f"{CISLINK_URL}/Reports/UploadHistory.aspx")
        time.sleep(3)
        
        # Ждем загрузки таблицы
        self.wait.until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_gvUploads"))
        )
        
    def get_table_rows(self):
        """Получение строк таблицы (свежие элементы)"""
        table = self.driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_gvUploads")
        return table.find_elements(By.TAG_NAME, "tr")[1:]  # Пропускаем заголовок
    
    def parse_error_details_for_row(self, row_index):
        """
        Парсинг деталей ошибки для конкретной строки.
        Перезагружает страницу и находит строку заново для избежания stale element.
        """
        try:
            # Перезагружаем страницу чтобы получить свежие элементы
            self.driver.get(f"{CISLINK_URL}/Reports/UploadHistory.aspx")
            time.sleep(2)
            
            # Ждем загрузки таблицы
            self.wait.until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_gvUploads"))
            )
            
            # Получаем свежие строки
            rows = self.get_table_rows()
            if row_index >= len(rows):
                logger.warning(f"Строка {row_index} не найдена после перезагрузки")
                return None
                
            row = rows[row_index]
            
            # Ищем ссылку Error в строке
            error_link = None
            
            # Способ 1: по ID
            row_num = str(row_index + 2).zfill(2)
            link_id = f"ctl00_ContentPlaceHolder1_gvUploads_ctl{row_num}_lnkView"
            try:
                error_link = self.driver.find_element(By.ID, link_id)
                logger.info(f"Найдена ссылка Error по ID: {link_id}")
            except NoSuchElementException:
                pass
            
            # Способ 2: поиск по тексту в строке
            if not error_link:
                try:
                    links = row.find_elements(By.TAG_NAME, "a")
                    for link in links:
                        if "Error" in link.text or "lnkView" in link.get_attribute("id") or "":
                            error_link = link
                            logger.info("Найдена ссылка Error по тексту")
                            break
                except:
                    pass
            
            if not error_link:
                logger.warning(f"Ссылка Error не найдена для строки {row_index}")
                return None
            
            # Кликаем на ссылку Error
            logger.info("Кликаем на ссылку Error...")
            try:
                # Скроллим к элементу
                self.driver.execute_script("arguments[0].scrollIntoView(true);", error_link)
                time.sleep(0.5)
                
                # Пробуем обычный клик
                error_link.click()
            except ElementClickInterceptedException:
                # Если не получилось, используем JavaScript
                self.driver.execute_script("arguments[0].click();", error_link)
            
            # Ждем появления popup с деталями
            time.sleep(2)
            
            # Ищем элемент с текстом ошибки разными способами
            error_text = ""
            error_html = ""
            
            # Способ 1: по ID
            selectors = [
                "ctl00_ContentPlaceHolder1_lblDetails",
                "lblDetails",
                "ContentPlaceHolder1_lblDetails"
            ]
            
            for selector in selectors:
                try:
                    element = self.driver.find_element(By.ID, selector)
                    error_text = element.text
                    error_html = element.get_attribute("innerHTML")
                    if error_text:
                        logger.info(f"Найден текст ошибки по ID {selector}")
                        break
                except NoSuchElementException:
                    continue
            
            # Способ 2: по CSS селекторам
            if not error_text:
                css_selectors = [
                    ".pnlBackGround span",
                    ".pnlBackGround div",
                    "[id*='lblDetails']",
                    "[id*='Details']",
                    ".modal-body",
                    ".popup-content"
                ]
                
                for css in css_selectors:
                    try:
                        elements = self.driver.find_elements(By.CSS_SELECTOR, css)
                        for el in elements:
                            text = el.text.strip()
                            if text and len(text) > 20 and ("Шаг" in text or "найден" in text or "ошибк" in text.lower()):
                                error_text = text
                                error_html = el.get_attribute("innerHTML")
                                logger.info(f"Найден текст ошибки по CSS: {css}")
                                break
                        if error_text:
                            break
                    except:
                        continue
            
            # Способ 3: поиск по XPath
            if not error_text:
                try:
                    # Ищем любой элемент с текстом про шаг или ошибку
                    xpath = "//*[contains(text(), 'Шаг') or contains(text(), 'найден') or contains(text(), 'ошибк')]"
                    elements = self.driver.find_elements(By.XPATH, xpath)
                    for el in elements:
                        text = el.text.strip()
                        if len(text) > 50:
                            error_text = text
                            error_html = el.get_attribute("innerHTML")
                            logger.info("Найден текст ошибки по XPath")
                            break
                except:
                    pass
            
            # Способ 4: получить весь HTML popup и извлечь текст
            if not error_text:
                try:
                    popup_selectors = [".pnlBackGround", "[class*='popup']", "[class*='modal']", "[id*='pnl']"]
                    for sel in popup_selectors:
                        try:
                            popup = self.driver.find_element(By.CSS_SELECTOR, sel)
                            full_text = popup.text
                            if "Шаг" in full_text or "найден" in full_text:
                                error_text = full_text
                                error_html = popup.get_attribute("innerHTML")
                                logger.info(f"Найден текст ошибки в popup: {sel}")
                                break
                        except:
                            continue
                except:
                    pass
            
            logger.info(f"Получен текст ошибки ({len(error_text)} символов)")
            
            if error_text:
                logger.info(f"Первые 200 символов: {error_text[:200]}")
            
            # Закрываем popup
            self._close_popup()
            
            # Парсим структурированные данные
            if error_text:
                return self._parse_error_content(error_text, error_html)
            
            return None
            
        except Exception as e:
            logger.error(f"Ошибка при парсинге деталей: {e}")
            # Пытаемся закрыть popup в любом случае
            try:
                self._close_popup()
            except:
                pass
            return None
    
    def _close_popup(self):
        """Закрытие popup окна"""
        try:
            # Способ 1: кнопка Закрыть
            close_buttons = self.driver.find_elements(By.CSS_SELECTOR, "input[value='Закрыть'], input[value='Close'], button.close, .btn-close")
            for btn in close_buttons:
                try:
                    btn.click()
                    time.sleep(0.5)
                    return
                except:
                    continue
            
            # Способ 2: клик вне popup
            try:
                overlay = self.driver.find_element(By.CSS_SELECTOR, ".modal-backdrop, .overlay, .pnlBackGround")
                self.driver.execute_script("arguments[0].click();", overlay)
                time.sleep(0.5)
                return
            except:
                pass
            
            # Способ 3: Escape
            from selenium.webdriver.common.keys import Keys
            self.driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            time.sleep(0.5)
            
        except Exception as e:
            logger.warning(f"Не удалось закрыть popup: {e}")
    
    def _parse_error_content(self, text, html):
        """Парсинг содержимого ошибки в структурированный формат"""
        result = {
            "raw_text": text,
            "errors": []
        }
        
        # Извлекаем номер шага
        step_match = re.search(r'Шаг\s*(\d+)', text)
        step = step_match.group(0) if step_match else None
        
        # Определяем файл с ошибкой
        file_type = None
        file_patterns = {
            'pdrest': 'остатков',
            'pdfact': 'продаж', 
            'pdcatal': 'справочник товаров',
            'pdclient': 'справочник клиентов',
            'pddoc': 'документов',
            'pdwh': 'справочник складов',
            'pdseria': 'справочник серий'
        }
        
        for pattern, name in file_patterns.items():
            if pattern in text.lower():
                file_type = pattern
                break
        
        # Извлекаем затронутые поля
        fields = []
        field_patterns = [
            r'поля?\s+(\w+)',
            r'полей?\s+(\w+(?:\s*,\s*\w+)*)',
            r'значени[ея]\s+(\w+)'
        ]
        for pattern in field_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                fields.extend([f.strip() for f in match.split(',')])
        
        # Подсчет количества ошибок
        count = None
        count_match = re.search(r'(\d+)\s*(?:строк|записей|значений)', text)
        if count_match:
            count = int(count_match.group(1))
        
        # Проверяем усечение списка
        is_truncated = "(список неполный)" in text or "и ещё" in text
        
        # Парсим примеры из таблицы (если есть HTML)
        examples = []
        if html and '<table' in html.lower():
            examples = self._parse_error_table(html)
        
        error_entry = {
            "step": step,
            "file": file_type,
            "fields": list(set(fields)) if fields else None,
            "message": text[:500],  # Ограничиваем длину
            "count": count,
            "is_truncated": is_truncated,
            "examples": examples[:5] if examples else None  # Максимум 5 примеров
        }
        
        result["errors"].append(error_entry)
        
        return result
    
    def _parse_error_table(self, html):
        """Парсинг таблицы с примерами ошибок из HTML"""
        examples = []
        
        try:
            # Простой парсинг таблицы через regex
            # Ищем строки таблицы
            row_pattern = r'<tr[^>]*>(.*?)</tr>'
            cell_pattern = r'<td[^>]*>(.*?)</td>'
            header_pattern = r'<t[hd][^>]*>(.*?)</t[hd]>'
            
            rows = re.findall(row_pattern, html, re.IGNORECASE | re.DOTALL)
            
            if len(rows) < 2:
                return examples
            
            # Первая строка - заголовки
            headers = re.findall(header_pattern, rows[0], re.IGNORECASE | re.DOTALL)
            headers = [re.sub(r'<[^>]+>', '', h).strip() for h in headers]
            
            # Остальные строки - данные
            for row in rows[1:6]:  # Максимум 5 строк
                cells = re.findall(cell_pattern, row, re.IGNORECASE | re.DOTALL)
                cells = [re.sub(r'<[^>]+>', '', c).strip() for c in cells]
                
                if cells and len(cells) == len(headers):
                    example = dict(zip(headers, cells))
                    examples.append(example)
                    
        except Exception as e:
            logger.warning(f"Ошибка парсинга таблицы: {e}")
        
        return examples
    
    def scrape_reports(self):
        """Сбор данных из таблицы отчетов"""
        logger.info("Сбор данных из таблицы...")
        
        reports = []
        rows_with_errors = []  # Индексы строк с ошибками
        
        # Первый проход: собираем основные данные
        rows = self.get_table_rows()
        
        for idx, row in enumerate(rows):
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) < 12:
                    continue
                
                # Определяем статус
                status_text = cells[1].text.strip()
                is_success = "Удачная" in status_text
                has_error_link = "Error" in row.text or "lnkView" in row.get_attribute("innerHTML")
                
                # Запоминаем индексы строк с ошибками для второго прохода
                if not is_success and has_error_link:
                    rows_with_errors.append(idx)
                
                report = {
                    "upload_datetime": cells[0].text.strip(),
                    "upload_status": "success" if is_success else "error",
                    "error_file_type": cells[2].text.strip() if not is_success else None,
                    "distr_code": cells[3].text.strip(),
                    "distr_id": cells[4].text.strip(),
                    "distr_name": cells[5].text.strip(),
                    "city": cells[6].text.strip(),
                    "doc_max_date": cells[7].text.strip() or None,
                    "doc_period": cells[8].text.strip() or None,
                    "stock_max_date": cells[9].text.strip() or None,
                    "stock_period": cells[10].text.strip() or None,
                    "connection_type": cells[11].text.strip() if len(cells) > 11 else "API",
                    "errors": None,
                    "_row_index": idx  # Сохраняем индекс для второго прохода
                }
                
                reports.append(report)
                
            except StaleElementReferenceException:
                logger.warning(f"Stale element на строке {idx}, пропускаем")
                continue
            except Exception as e:
                logger.warning(f"Ошибка обработки строки {idx}: {e}")
                continue
        
        logger.info(f"Собрано {len(reports)} записей, с ошибками: {len(rows_with_errors)}")
        
        # Второй проход: парсим детали ошибок
        errors_parsed = 0
        for report in reports:
            if report["_row_index"] in rows_with_errors:
                logger.info(f"Парсим ошибку для дистрибьютора {report['distr_name']}...")
                error_details = self.parse_error_details_for_row(report["_row_index"])
                if error_details:
                    report["errors"] = error_details
                    errors_parsed += 1
                    logger.info(f"Успешно спарсили детали ошибки")
            
            # Удаляем служебное поле
            del report["_row_index"]
        
        logger.info(f"Всего записей: {len(reports)}, с ошибками: {len(rows_with_errors)}, с деталями: {errors_parsed}")
        
        return reports
    
    def send_to_api(self, reports):
        """Отправка данных в API"""
        if not reports:
            logger.warning("Нет данных для отправки")
            return False
        
        try:
            response = requests.post(
                API_URL,
                json={"reports": reports},
                headers={
                    "Content-Type": "application/json",
                    "X-API-Key": API_KEY
                },
                timeout=30
            )
            
            result = response.json()
            logger.info(f"Результат отправки: {result}")
            
            return result.get("success", False)
            
        except Exception as e:
            logger.error(f"Ошибка отправки в API: {e}")
            return False
    
    def close(self):
        """Закрытие браузера"""
        if self.driver:
            self.driver.quit()
    
    def run(self):
        """Основной метод запуска агента"""
        logger.info("=" * 50)
        logger.info("Агент CISLink v2.1 (с исправленным парсингом ошибок)")
        logger.info("=" * 50)
        
        try:
            self.setup_browser()
            self.login()
            self.navigate_to_reports()
            
            reports = self.scrape_reports()
            
            if reports:
                success = self.send_to_api(reports)
                if success:
                    logger.info("Синхронизация завершена успешно")
                else:
                    logger.error("Ошибка при отправке данных")
            else:
                logger.warning("Не удалось собрать данные")
                
        except Exception as e:
            logger.error(f"Критическая ошибка: {e}")
            raise
        finally:
            self.close()


if __name__ == "__main__":
    agent = CISLinkAgent()
    agent.run()
