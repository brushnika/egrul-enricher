# -*- coding: utf-8 -*-
"""
СКРИПТ 2: ДОПОЛНЕНИЕ ДАННЫХ ИЗ ЕГРЮЛ (с автоматической сменой прокси)
Использует библиотеку free-proxy для получения рабочих прокси
"""

import time
import random
import re
import os
import sys
import json
import requests
from datetime import datetime
from typing import Dict, Optional, List
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from urllib3.exceptions import InsecureRequestWarning

# Отключаем предупреждения о SSL
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

# Пытаемся импортировать библиотеку для прокси
try:
    from fp.fp import FreeProxy
    PROXY_LIBRARY_AVAILABLE = True
    print("✅ Библиотека free-proxy установлена")
except ImportError:
    PROXY_LIBRARY_AVAILABLE = False
    print("❌ Библиотека free-proxy НЕ установлена")
    print("   Установите: pip install free-proxy")

# =====================================================================
# НАСТРОЙКИ
# =====================================================================

INPUT_EXCEL_PATH = "Брянская_область_Кожевенные_и_обувные_производства_2026-06-29_01-46.xlsx"
COLUMN_NAME_HEADER = "НАИМ"
COLUMN_ADDRESS_HEADER = "Адрес"
OUTPUT_DIR = "."

# Настройки прокси
USE_PROXY = True  # Включить использование прокси
PROXY_TYPE = 'http'  # Тип прокси: http, https, socks4, socks5
PROXY_COUNTRY = None  # Страна прокси (None - любая) например: ['US', 'RU', 'DE']
PROXY_ANONYMOUS = True  # Только анонимные прокси

# Настройки времени
TIMEOUT = 60
PAUSE_MIN = 8.0
PAUSE_MAX = 15.0
HUMAN_TYPING_DELAY = 0.1

# =====================================================================


def random_pause():
    """Случайная пауза между запросами."""
    time.sleep(random.uniform(PAUSE_MIN, PAUSE_MAX))


def human_like_pause():
    """Короткая пауза как у человека."""
    time.sleep(random.uniform(1.0, 3.0))


def get_random_user_agent() -> str:
    """Возвращает случайный User-Agent."""
    user_agents = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/121.0',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.1 Safari/605.1.15',
        'Mozilla/5.0 (iPhone; CPU iPhone OS 17_2 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.2 Mobile/15E148 Safari/604.1',
    ]
    return random.choice(user_agents)


def get_proxy_from_library() -> Optional[str]:
    """
    Получает рабочий прокси из библиотеки free-proxy.
    Возвращает строку с прокси или None.
    """
    if not PROXY_LIBRARY_AVAILABLE or not USE_PROXY:
        return None
    
    try:
        print("    → Получение прокси из библиотеки...")
        
        # Настройка параметров для FreeProxy
        proxy_params = {
            'type': PROXY_TYPE,
            'rand': True,  # Случайный прокси
            'timeout': 5,  # Таймаут проверки
            'https': True,  # Поддержка HTTPS
        }
        
        # Добавляем страну, если указана
        if PROXY_COUNTRY:
            proxy_params['country_id'] = PROXY_COUNTRY
        
        # Добавляем анонимность
        if PROXY_ANONYMOUS:
            proxy_params['anon'] = True
        
        # Получаем прокси
        proxy = FreeProxy(**proxy_params).get()
        
        if proxy:
            print(f"    ✅ Получен прокси: {proxy}")
            return proxy
        else:
            print("    ❌ Не удалось получить прокси")
            return None
            
    except Exception as e:
        print(f"    ❌ Ошибка получения прокси: {e}")
        return None


def get_proxy_from_file() -> Optional[str]:
    """
    Получает прокси из файла (запасной вариант).
    Файл proxies.txt должен содержать по одному прокси на строку.
    """
    try:
        proxy_file = "proxies.txt"
        if os.path.exists(proxy_file):
            with open(proxy_file, 'r') as f:
                proxies = [line.strip() for line in f if line.strip()]
                if proxies:
                    proxy = random.choice(proxies)
                    print(f"    ✅ Получен прокси из файла: {proxy}")
                    return proxy
    except Exception as e:
        print(f"    ❌ Ошибка чтения файла с прокси: {e}")
    
    return None


def create_human_session() -> Optional[requests.Session]:
    """
    Создает сессию с прокси и имитацией человека.
    """
    try:
        session = requests.Session()
        
        # Случайный User-Agent
        session.headers.update({
            'User-Agent': get_random_user_agent(),
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
            'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'none',
            'Sec-Fetch-User': '?1',
            'Cache-Control': 'max-age=0',
        })
        
        # Получаем прокси
        proxy = None
        if USE_PROXY:
            # Сначала пробуем получить из библиотеки
            proxy = get_proxy_from_library()
            
            # Если не получилось, пробуем из файла
            if not proxy:
                print("    → Пробуем получить прокси из файла...")
                proxy = get_proxy_from_file()
        
        # Настраиваем прокси
        if proxy:
            proxy_url = f"http://{proxy}"
            session.proxies = {
                'http': proxy_url,
                'https': proxy_url,
            }
        
        # Имитация человека
        print("    → Открытие сайта (имитация человека)...")
        human_like_pause()
        
        response = session.get(
            "https://egrul.nalog.ru/index.html",
            timeout=TIMEOUT,
            verify=False,
            allow_redirects=True
        )
        
        if response.status_code == 200:
            print(f"    ✅ Сессия создана (статус: {response.status_code})")
            return session
        else:
            print(f"    ❌ Ошибка: статус {response.status_code}")
            return None
            
    except Exception as e:
        print(f"    ❌ Ошибка создания сессии: {e}")
        return None


def clean_company_name(name: str) -> str:
    """Очищает название компании для поиска в ЕГРЮЛ."""
    if not name or name == "Н/Д":
        return ""
    
    cleaned = re.sub(r'(?i)\b(ООО|ЗАО|ПАО|ОАО|АО|ИП|ЧП)\b', '', name)
    cleaned = cleaned.replace('«', '').replace('»', '').replace('"', '').replace('"', '')
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    
    return cleaned


def search_egrul_with_selenium(driver, company_name: str, proxy: Optional[str] = None) -> Dict:
    """
    Поиск в ЕГРЮЛ через Selenium с имитацией человека.
    """
    result = {
        "Юр. наименование (ЕГРЮЛ)": "Н/Д",
        "ИНН": "Н/Д",
        "ОГРН": "Н/Д",
        "Должность руководителя": "Н/Д",
        "ФИО руководителя": "Н/Д",
        "Примечание ЕГРЮЛ": "",
    }
    
    clean_name = clean_company_name(company_name)
    if not clean_name:
        result["Примечание ЕГРЮЛ"] = "Пустое название"
        return result
    
    print(f"    Поиск: {clean_name}")
    
    try:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        
        # Открываем сайт
        print("    → Открытие сайта...")
        driver.get("https://egrul.nalog.ru/")
        human_like_pause()
        
        # Ждем загрузки
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input#query"))
        )
        
        # Вводим запрос
        print("    → Ввод запроса...")
        search_input = driver.find_element(By.CSS_SELECTOR, "input#query")
        search_input.click()
        human_like_pause()
        search_input.clear()
        human_like_pause()
        
        for char in clean_name:
            search_input.send_keys(char)
            time.sleep(random.uniform(0.03, HUMAN_TYPING_DELAY))
        
        human_like_pause()
        
        # Нажимаем кнопку
        print("    → Отправка запроса...")
        search_btn = driver.find_element(By.CSS_SELECTOR, "button#btnSearch")
        search_btn.click()
        
        time.sleep(random.uniform(3, 5))
        
        # Проверяем результаты
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.results-table, div.res-row, div.res"))
            )
            print("    ✅ Результаты загружены")
        except:
            print("    ❌ Результаты не загружены")
            result["Примечание ЕГРЮЛ"] = "Результаты не загружены"
            return result
        
        # Находим результаты
        rows = driver.find_elements(By.CSS_SELECTOR, "div.res-row, tr.results-table-row, div.res")
        if not rows:
            print("    ❌ Записей не найдено")
            result["Примечание ЕГРЮЛ"] = "Совпадений не найдено"
            return result
        
        first_row = rows[0]
        print("    ✅ Найдена запись")
        
        # Извлекаем данные
        row_text = first_row.text
        
        inn_match = re.search(r"ИНН[:\s]*(\d{10,12})", row_text)
        ogrn_match = re.search(r"ОГРН[:\s]*(\d{13,15})", row_text)
        
        if inn_match:
            result["ИНН"] = inn_match.group(1)
            print(f"    ИНН: {result['ИНН']}")
        if ogrn_match:
            result["ОГРН"] = ogrn_match.group(1)
            print(f"    ОГРН: {result['ОГРН']}")
        
        try:
            name_elem = driver.find_element(By.CSS_SELECTOR, ".res-cell-name, .name, .res-name")
            result["Юр. наименование (ЕГРЮЛ)"] = name_elem.text.strip()
            print(f"    Название: {result['Юр. наименование (ЕГРЮЛ)']}")
        except:
            pass
        
        # Получаем детали
        print("    → Получение деталей...")
        try:
            first_row.click()
            time.sleep(random.uniform(2, 4))
        except:
            pass
        
        # Ищем данные на странице
        try:
            page_text = driver.find_element(By.TAG_NAME, "body").text
            
            fio_patterns = [
                r'генеральный директор[:\s]*([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)',
                r'директор[:\s]*([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)',
                r'руководитель[:\s]*([А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+)',
                r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+',
            ]
            
            for pattern in fio_patterns:
                match = re.search(pattern, page_text, re.IGNORECASE)
                if match:
                    fio = match.group(1) if '(' in pattern else match.group(0)
                    if len(fio.split()) >= 3:
                        result["ФИО руководителя"] = fio
                        print(f"    ✅ Найдено ФИО: {fio}")
                        break
            
            pos_pattern = r'(Генеральный директор|Директор|Президент|Управляющий|Руководитель)'
            pos_match = re.search(pos_pattern, page_text, re.IGNORECASE)
            if pos_match:
                result["Должность руководителя"] = pos_match.group(1).capitalize()
                print(f"    ✅ Найдена должность: {result['Должность руководителя']}")
            
            if result["ФИО руководителя"] == "Н/Д":
                result["Примечание ЕГРЮЛ"] = "Найдена запись, но не удалось извлечь ФИО"
                
        except Exception as e:
            print(f"    ❌ Ошибка при извлечении данных: {e}")
        
        return result
        
    except Exception as e:
        result["Примечание ЕГРЮЛ"] = f"Ошибка: {str(e)[:100]}"
        print(f"    ❌ Ошибка: {e}")
        return result


def format_recipient(position: str, name: str, fio: str) -> str:
    """Формирует фразу 'Кому адресовано письмо'."""
    if "Н/Д" in (position, name, fio):
        return "Н/Д"
    
    return f"{position} {name} {fio}"


def save_result(workbook, input_path: str, output_dir: str) -> str:
    """Сохраняет результат в новый файл."""
    os.makedirs(output_dir, exist_ok=True)
    
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    output_path = os.path.join(output_dir, f"{base_name}_ЕГРЮЛ_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx")
    
    workbook.save(output_path)
    return output_path


def main():
    print("=" * 70)
    print("СКРИПТ 2: ДОПОЛНЕНИЕ ДАННЫХ ИЗ ЕГРЮЛ (с прокси)")
    print(f"Исходный файл: {INPUT_EXCEL_PATH}")
    print("=" * 70)
    
    if not os.path.exists(INPUT_EXCEL_PATH):
        print(f"ОШИБКА: файл не найден: {INPUT_EXCEL_PATH}")
        sys.exit(1)
    
    # Проверяем Selenium
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        from selenium.webdriver.chrome.options import Options
        from webdriver_manager.chrome import ChromeDriverManager
        print("✅ Selenium и драйверы установлены")
    except ImportError:
        print("❌ Selenium не установлен. Установите: pip install selenium webdriver-manager")
        sys.exit(1)
    
    # Настройка браузера
    print("\nНастройка браузера...")
    options = Options()
    
    prefs = {
        "profile.managed_default_content_settings.images": 2,
        "disk-cache-size": 4096,
    }
    options.add_experimental_option("prefs", prefs)
    
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-notifications")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    
    # Получаем прокси для браузера
    browser_proxy = None
    if USE_PROXY and PROXY_LIBRARY_AVAILABLE:
        browser_proxy = get_proxy_from_library()
        if browser_proxy:
            options.add_argument(f'--proxy-server={browser_proxy}')
            print(f"   Браузер использует прокси: {browser_proxy}")
        else:
            print("   ⚠️ Не удалось получить прокси для браузера")
    
    # Создаем драйвер
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        print("✅ Браузер запущен")
    except Exception as e:
        print(f"❌ Ошибка запуска браузера: {e}")
        sys.exit(1)
    
    # Загружаем Excel
    print("\nЗагрузка Excel-файла...")
    wb = load_workbook(INPUT_EXCEL_PATH)
    ws = wb.active
    
    headers = [cell.value for cell in ws[1]]
    
    if COLUMN_NAME_HEADER not in headers:
        print(f"ОШИБКА: столбец '{COLUMN_NAME_HEADER}' не найден")
        print(f"Доступные столбцы: {headers}")
        driver.quit()
        sys.exit(1)
    
    name_col = headers.index(COLUMN_NAME_HEADER) + 1
    addr_col = headers.index(COLUMN_ADDRESS_HEADER) + 1 if COLUMN_ADDRESS_HEADER in headers else None
    
    # Добавляем новые столбцы
    new_headers = [
        "Юр. наименование (ЕГРЮЛ)",
        "ИНН",
        "ОГРН",
        "Должность руководителя",
        "ФИО руководителя",
        "Кому адресовано письмо",
        "Примечание ЕГРЮЛ",
    ]
    
    start_col = ws.max_column + 1
    for i, header in enumerate(new_headers):
        cell = ws.cell(row=1, column=start_col + i)
        cell.value = header
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    total_rows = ws.max_row
    print(f"Найдено компаний для обработки: {total_rows - 1}")
    print(f"Пауза между компаниями: {PAUSE_MIN}-{PAUSE_MAX} секунд")
    print(f"Использование прокси: {'включено' if USE_PROXY else 'выключено'}")
    
    processed = 0
    found = 0
    
    try:
        for row in range(2, total_rows + 1):
            company_name = ws.cell(row=row, column=name_col).value
            
            if not company_name or company_name == "Н/Д":
                for i in range(len(new_headers)):
                    ws.cell(row=row, column=start_col + i, value="Н/Д")
                continue
            
            print(f"\n[{row - 1}/{total_rows - 1}] Обработка: {company_name}")
            
            # Проверяем, нужно ли сменить прокси
            if USE_PROXY and browser_proxy and processed % 3 == 0:
                # Меняем прокси каждые 3 компании
                print("    → Смена прокси...")
                new_proxy = get_proxy_from_library()
                if new_proxy:
                    # Пересоздаем драйвер с новым прокси
                    driver.quit()
                    options.add_argument(f'--proxy-server={new_proxy}')
                    service = Service(ChromeDriverManager().install())
                    driver = webdriver.Chrome(service=service, options=options)
                    browser_proxy = new_proxy
                    print(f"    ✅ Прокси сменен на: {browser_proxy}")
            
            # Поиск в ЕГРЮЛ
            egrul_data = search_egrul_with_selenium(driver, company_name, browser_proxy)
            
            recipient = format_recipient(
                egrul_data["Должность руководителя"],
                egrul_data["Юр. наименование (ЕГРЮЛ)"] if egrul_data["Юр. наименование (ЕГРЮЛ)"] != "Н/Д" else company_name,
                egrul_data["ФИО руководителя"]
            )
            
            values = [
                egrul_data["Юр. наименование (ЕГРЮЛ)"],
                egrul_data["ИНН"],
                egrul_data["ОГРН"],
                egrul_data["Должность руководителя"],
                egrul_data["ФИО руководителя"],
                recipient,
                egrul_data["Примечание ЕГРЮЛ"],
            ]
            
            for i, value in enumerate(values):
                ws.cell(row=row, column=start_col + i, value=value)
            
            processed += 1
            if egrul_data["ФИО руководителя"] != "Н/Д":
                found += 1
                print(f"    ✅ Найдено: {egrul_data['ФИО руководителя']}")
            else:
                print(f"    ❌ Не найдено: {egrul_data['Примечание ЕГРЮЛ']}")
            
            if processed % 3 == 0:
                temp_path = save_result(wb, INPUT_EXCEL_PATH, OUTPUT_DIR)
                print(f"    💾 Промежуточное сохранение: {temp_path}")
            
            print(f"    ⏳ Пауза {PAUSE_MIN}-{PAUSE_MAX} сек...")
            random_pause()
            
    except KeyboardInterrupt:
        print("\n⚠️ Прервано пользователем. Сохраняем промежуточный результат...")
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
    finally:
        driver.quit()
        print("✅ Браузер закрыт")
    
    output_path = save_result(wb, INPUT_EXCEL_PATH, OUTPUT_DIR)
    
    print("\n" + "=" * 70)
    print("СТАТИСТИКА")
    print("=" * 70)
    print(f"Обработано компаний: {processed}")
    print(f"Найдено руководителей: {found}")
    print(f"Файл сохранен: {output_path}")
    print("=" * 70)


if __name__ == "__main__":
    main()