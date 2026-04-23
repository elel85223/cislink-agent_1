"""
Агент автоматической привязки непривязанных товаров в CISLink
Версия 1.2 - добавлен шаг выбора всех дистрибьюторов перед переходом в справочник

Логика работы:
1. Логинится в CISLink
2. Заходит на страницу выбора дистрибьюторов (Dictionary/Default.aspx)
   и нажимает "Выбрать все" (чекбокс cbDistrs)
3. Открывает страницу непривязанных товаров (reportId=13, contentId=3)
4. Собирает список товаров (до MAX_ITEMS_PER_RUN) с ссылками на карточки и артикулами
5. Для каждого товара:
   - Открывает карточку
   - Вбивает артикул дистрибьютора в поле #inpTextCode ("Номенклатура Артикул")
   - Триггерит blur - система запускает валидацию (функция enter_code)
   - Если #lblTextCodeError стал видим - артикул не распознан, пропуск
   - Если #btnSave стал видим и #inpManfCode заполнен - жмем Сохранить
6. Отправляет итоговый отчет в API
"""

import os
import time
import logging
from datetime import datetime
from typing import Optional, List, Dict, Any

import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementClickInterceptedException,
    StaleElementReferenceException
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
    'distributors_page_url': 'https://b2b.cislinkdts.com/Dictionary/Default.aspx',
    'unlinked_page_url': 'https://b2b.cislinkdts.com/Dictionary/DGrid.aspx?reportId=13&contentId=3',
    'cislink_login': os.getenv('CISLINK_LOGIN'),
    'cislink_password': os.getenv('CISLINK_PASSWORD'),
    'api_url': os.getenv('API_URL'),
    'api_key': os.getenv('API_KEY'),
    'debug_mode': os.getenv('DEBUG_MODE', 'False').lower() == 'true',
    'timeout': 30,
    'max_items_per_run': int(os.getenv('MAX_ITEMS_PER_RUN', '50')),
    'validation_wait_seconds': 3,
    'save_wait_seconds': 3,
}

# Точные id элементов, полученные по результатам разведки HTML-разметки
SELECTORS = {
    # Страница выбора дистрибьюторов
    'select_all_distrs_checkbox': 'cbDistrs',
    # Страница списка непривязанных товаров
    'list_table': 'ctl00_ContentPlaceHolder1_ucDGrid_gvList',
    'page_size_dropdown': 'ctl00_ContentPlaceHolder1_ucDGrid_ddlPageSize',
    'item_type_dropdown': 'ctl00_ContentPlaceHolder1_ucDGrid_ddlItemType',
    # Ссылки на карточку в строках: id шаблона ctl00_ContentPlaceHolder1_ucDGrid_gvList_ctl{NN}_hlLabel1
    'row_link_suffix': '_hlLabel1',
    # Поля на карточке товара
    'card_article_input': 'inpTextCode',    # Номенклатура Артикул (ввод)
    'card_manf_id_input': 'inpManfCode',    # Номенклатура ID (заполняется системой)
    'card_ean_input': 'inpEAN',             # Штрихкод
    'card_product_select': 'ddlProducts',   # Номенклатура Наименование (select)
    'card_direction_select': 'ddl2',        # Направление
    'card_line_select': 'ddl1',             # Линейка
    'card_save_button': 'btnSave',
    'card_error_label': 'lblTextCodeError',
    'card_form': 'aspnetForm',
}


class CISLinkLinker:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.results: List[Dict[str, Any]] = []

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
        self.driver.execute_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
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
                logger.info("Авторизация успешна")
                return True
            logger.error(f"Неожиданный URL после авторизации: {current_url}")
            return False
        except Exception as e:
            logger.error(f"Ошибка авторизации: {e}")
            return False

    def select_all_distributors(self) -> bool:
        """
        Переходит на страницу Dictionary/Default.aspx и отмечает чекбокс 'Выбрать все'.
        Это обязательный шаг - без выбранных дистрибьюторов справочник товаров пуст.
        """
        logger.info("Переход на страницу дистрибьюторов и выбор всех...")
        try:
            self.driver.get(CONFIG['distributors_page_url'])
            time.sleep(3)
            try:
                checkbox = self.wait.until(
                    EC.presence_of_element_located((By.ID, SELECTORS['select_all_distrs_checkbox']))
                )
            except TimeoutException:
                logger.error(f"Чекбокс #{SELECTORS['select_all_distrs_checkbox']} не найден")
                return False

            if checkbox.is_selected():
                logger.info("Чекбокс 'Выбрать все' уже активен - ничего не делаем")
                return True

            try:
                checkbox.click()
            except ElementClickInterceptedException:
                self.driver.execute_script("arguments[0].click();", checkbox)
            time.sleep(2)

            checkbox = self.driver.find_element(By.ID, SELECTORS['select_all_distrs_checkbox'])
            if checkbox.is_selected():
                logger.info("Все дистрибьюторы выбраны")
                return True
            logger.warning("После клика чекбокс не активен - пробуем еще раз через JS")
            self.driver.execute_script("arguments[0].click();", checkbox)
            time.sleep(2)
            checkbox = self.driver.find_element(By.ID, SELECTORS['select_all_distrs_checkbox'])
            return checkbox.is_selected()
        except Exception as e:
            logger.error(f"Ошибка выбора всех дистрибьюторов: {e}")
            return False

    def open_unlinked_page(self) -> bool:
        logger.info("Открываем страницу непривязанных товаров...")
        try:
            self.driver.get(CONFIG['unlinked_page_url'])
            time.sleep(3)
            self.wait.until(EC.presence_of_element_located((By.ID, SELECTORS['list_table'])))
            return True
        except TimeoutException:
            logger.error("Таблица непривязанных товаров не появилась за таймаут")
            return False
        except Exception as e:
            logger.error(f"Ошибка открытия страницы непривязанных товаров: {e}")
            return False

    def collect_unlinked_items(self) -> List[Dict[str, str]]:
        """
        Собирает список непривязанных товаров со страницы.
        Использует селектор ссылок вида a[id$='_hlLabel1'] внутри таблицы gvList.
        Структура колонок (по отчету разведки):
            td[0] - № строки
            td[1] - Название товара дистрибьютора (ссылка)
            td[2] - Номенклатура Наименование
            td[3] - Номенклатура ID
            td[4] - Артикул дистрибьютора
            td[5] - Штрих-код дистрибьютора
            td[6] - Код товара дистрибьютора
            td[7] - кнопка-иконка Редактировать
            td[8] - кнопка-иконка Удалить
        """
        logger.info("Собираем данные о непривязанных товарах...")
        items: List[Dict[str, str]] = []
        try:
            table = self.driver.find_element(By.ID, SELECTORS['list_table'])
            link_elements = table.find_elements(
                By.CSS_SELECTOR,
                f"a[id^='{SELECTORS['list_table']}_ctl'][id$='{SELECTORS['row_link_suffix']}']"
            )
            logger.info(f"Найдено ссылок на карточки: {len(link_elements)}")

            for link in link_elements:
                try:
                    row = link.find_element(By.XPATH, "./ancestor::tr[1]")
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) < 7:
                        continue
                    product_name = link.text.strip()
                    detail_url = link.get_attribute("href") or ""
                    article = cells[4].text.strip()
                    distr_code = cells[6].text.strip()

                    if not article or not detail_url:
                        logger.debug(f"Пропуск строки без артикула/ссылки: {product_name}")
                        continue

                    items.append({
                        'product_name': product_name,
                        'article': article,
                        'detail_url': detail_url,
                        'distr_code': distr_code,
                    })

                    if len(items) >= CONFIG['max_items_per_run']:
                        logger.info(
                            f"Достигнут лимит {CONFIG['max_items_per_run']} товаров за запуск "
                            f"(всего на странице {len(link_elements)})"
                        )
                        break
                except Exception as e:
                    logger.debug(f"Ошибка обработки строки: {e}")
                    continue
            logger.info(f"Собрано {len(items)} непривязанных товаров для обработки")
            return items
        except NoSuchElementException:
            logger.warning("Таблица gvList не найдена на странице")
            return []
        except Exception as e:
            logger.error(f"Ошибка сбора непривязанных товаров: {e}")
            return []

    def process_item(self, item: Dict[str, str]) -> Dict[str, Any]:
        """
        Обрабатывает один товар: открывает карточку, вбивает артикул,
        дергает blur (триггерит валидацию через enter_code),
        проверяет результат по видимости #lblTextCodeError и #btnSave,
        сохраняет или пропускает.
        """
        result = {
            'product_name': item['product_name'],
            'article': item['article'],
            'distr_code': item.get('distr_code', ''),
            'detail_url': item['detail_url'],
            'status': 'unknown',
            'message': '',
            'nomenclature_id': '',
            'processed_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }
        try:
            logger.info(f"Обработка: {item['product_name']} (артикул {item['article']})")
            self.driver.get(item['detail_url'])
            time.sleep(2)

            try:
                article_input = self.wait.until(
                    EC.presence_of_element_located((By.ID, SELECTORS['card_article_input']))
                )
            except TimeoutException:
                result['status'] = 'error'
                result['message'] = f"Поле #{SELECTORS['card_article_input']} не найдено на карточке"
                return result

            # Устанавливаем значение через JS (обходит проверку интерактивности Selenium)
            # и явно триггерим события input/change/blur, чтобы сработал enter_code.
            set_value_script = """
                var el = document.getElementById(arguments[0]);
                if (!el) return false;
                el.focus();
                el.value = arguments[1];
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.blur();
                return true;
            """
            ok = self.driver.execute_script(
                set_value_script,
                SELECTORS['card_article_input'],
                item['article']
            )
            if not ok:
                result['status'] = 'error'
                result['message'] = f"Не удалось установить значение в #{SELECTORS['card_article_input']}"
                return result
            time.sleep(CONFIG['validation_wait_seconds'])

            # Проверка ошибки валидации
            error_text = self._read_error_label()
            if error_text:
                result['status'] = 'skipped_invalid_article'
                result['message'] = f"Ошибка валидации: {error_text}"
                logger.info(f"  -> пропуск: {error_text}")
                return result

            nomenclature_id = self._read_input_value(SELECTORS['card_manf_id_input'])
            save_visible = self._is_element_visible(SELECTORS['card_save_button'])

            if not nomenclature_id or not save_visible:
                result['status'] = 'skipped_no_match'
                result['message'] = (
                    f"ID не подтянулся (nomenclature_id='{nomenclature_id}', "
                    f"save_button_visible={save_visible})"
                )
                logger.info(f"  -> пропуск: {result['message']}")
                return result

            result['nomenclature_id'] = nomenclature_id
            logger.info(f"  -> ID найден: {nomenclature_id}, сохраняем")

            if self._click_save_button():
                time.sleep(CONFIG['save_wait_seconds'])
                result['status'] = 'linked'
                result['message'] = 'Товар успешно привязан'
                logger.info(f"  -> сохранено")
            else:
                result['status'] = 'error'
                result['message'] = 'Не удалось нажать кнопку Сохранить'
            return result
        except Exception as e:
            logger.error(f"Ошибка обработки товара {item.get('product_name')}: {e}")
            result['status'] = 'error'
            result['message'] = f'Исключение: {e}'
            return result

    def _read_input_value(self, element_id: str) -> str:
        try:
            el = self.driver.find_element(By.ID, element_id)
            return (el.get_attribute('value') or '').strip()
        except NoSuchElementException:
            return ''
        except Exception:
            return ''

    def _is_element_visible(self, element_id: str) -> bool:
        """Проверяет реальную видимость элемента (is_displayed + computed style display)."""
        try:
            el = self.driver.find_element(By.ID, element_id)
            if not el.is_displayed():
                return False
            display = self.driver.execute_script(
                "return window.getComputedStyle(arguments[0]).display;", el
            )
            return display != 'none'
        except NoSuchElementException:
            return False
        except Exception:
            return False

    def _read_error_label(self) -> str:
        """Читает текст из #lblTextCodeError, если он видим. Пустая строка = нет ошибки."""
        try:
            el = self.driver.find_element(By.ID, SELECTORS['card_error_label'])
            display = self.driver.execute_script(
                "return window.getComputedStyle(arguments[0]).display;", el
            )
            if display == 'none':
                return ''
            text = (el.text or '').strip()
            if not text:
                text = (el.get_attribute('innerText') or '').strip()
            return text
        except NoSuchElementException:
            return ''
        except Exception as e:
            logger.debug(f"Ошибка чтения label ошибки: {e}")
            return ''

    def _click_save_button(self) -> bool:
        """Клик по #btnSave."""
        try:
            save_btn = self.driver.find_element(By.ID, SELECTORS['card_save_button'])
            if not self._is_element_visible(SELECTORS['card_save_button']):
                logger.warning("Кнопка Сохранить не видима")
                return False
            try:
                self.driver.execute_script("arguments[0].scrollIntoView(true);", save_btn)
                time.sleep(0.3)
                save_btn.click()
            except ElementClickInterceptedException:
                self.driver.execute_script("arguments[0].click();", save_btn)
            return True
        except NoSuchElementException:
            logger.error("Кнопка #btnSave не найдена в DOM")
            return False
        except Exception as e:
            logger.error(f"Ошибка клика Сохранить: {e}")
            return False

    def run_linking(self):
        items = self.collect_unlinked_items()
        if not items:
            logger.info("Непривязанных товаров не найдено")
            return
        for idx, item in enumerate(items, start=1):
            logger.info(f"--- [{idx}/{len(items)}] ---")
            res = self.process_item(item)
            self.results.append(res)

    def close(self):
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass


class APIClient:
    def __init__(self):
        self.url = CONFIG['api_url']
        self.api_key = CONFIG['api_key']

    def send_results(self, results: List[Dict[str, Any]]) -> dict:
        if not self.url or not self.api_key:
            logger.warning("API_URL или API_KEY не заданы - пропуск отправки")
            return {'success': False, 'error': 'api_not_configured'}
        try:
            payload = {
                'api_key': self.api_key,
                'source': 'link_products_agent',
                'run_datetime': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'reports': results,
            }
            response = requests.post(self.url, json=payload, timeout=60)
            try:
                return response.json()
            except Exception:
                return {'success': response.ok, 'status_code': response.status_code}
        except Exception as e:
            return {'success': False, 'error': str(e)}


def summarize(results: List[Dict[str, Any]]):
    total = len(results)
    linked = sum(1 for r in results if r['status'] == 'linked')
    skipped_invalid = sum(1 for r in results if r['status'] == 'skipped_invalid_article')
    skipped_no_match = sum(1 for r in results if r['status'] == 'skipped_no_match')
    errors = sum(1 for r in results if r['status'] == 'error')
    logger.info("=" * 60)
    logger.info(f"Итого обработано: {total}")
    logger.info(f"  Привязано: {linked}")
    logger.info(f"  Пропущено (некорректный артикул): {skipped_invalid}")
    logger.info(f"  Пропущено (ID не подтянулся): {skipped_no_match}")
    logger.info(f"  Ошибки: {errors}")
    logger.info("=" * 60)


def main():
    logger.info("Агент привязки товаров CISLink v1.2")
    if not all([CONFIG['cislink_login'], CONFIG['cislink_password']]):
        logger.error("Не заданы CISLINK_LOGIN / CISLINK_PASSWORD")
        exit(1)

    linker = CISLinkLinker()
    try:
        linker.init_browser()
        if not linker.login():
            logger.error("Авторизация не удалась")
            exit(1)
        if not linker.select_all_distributors():
            logger.error("Не удалось выбрать всех дистрибьюторов")
            exit(1)
        if not linker.open_unlinked_page():
            logger.error("Не удалось открыть страницу непривязанных товаров")
            exit(1)
        linker.run_linking()
        summarize(linker.results)
        if linker.results:
            api_resp = APIClient().send_results(linker.results)
            logger.info(f"Ответ API: {api_resp}")
    finally:
        linker.close()


if __name__ == '__main__':
    main()
