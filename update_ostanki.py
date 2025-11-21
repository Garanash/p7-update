#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для обновления остатков из почты в файл Сборка Москва.xlsx
"""

import imaplib
import email
import os
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from datetime import datetime
import re

# Настройки почты
EMAIL_LOGIN = "almazgeobur.it@gmail.com"
EMAIL_PASSWORD = "almazgeobur2013"
IMAP_SERVER = "imap.gmail.com"
IMAP_PORT = 993

# Пути к файлам
MAIN_FILE = "Сборка Москва.xlsx"
TEMP_BOT_FILE = "temp_bot_file.xlsx"
SHEET_OSTANKI = "остатки"


def connect_to_email():
    """Подключение к почте через IMAP"""
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_LOGIN, EMAIL_PASSWORD)
        mail.select('inbox')
        return mail
    except Exception as e:
        print(f"Ошибка подключения к почте: {e}")
        return None


def find_latest_excel_attachment(mail):
    """Поиск последнего письма с Excel вложением"""
    try:
        # Поиск всех непрочитанных писем
        status, messages = mail.search(None, 'UNSEEN')
        if status != 'OK':
            print("Не удалось найти письма")
            return None
        
        email_ids = messages[0].split()
        
        # Если нет непрочитанных, ищем последние письма (последние 50)
        if not email_ids:
            status, messages = mail.search(None, 'ALL')
            email_ids = messages[0].split()[-50:]  # Берем последние 50 писем
        
        # Ищем последнее письмо с Excel вложением
        for email_id in reversed(email_ids):
            status, msg_data = mail.fetch(email_id, '(RFC822)')
            if status != 'OK':
                continue
            
            email_body = msg_data[0][1]
            email_message = email.message_from_bytes(email_body)
            
            # Проверяем вложения
            for part in email_message.walk():
                if part.get_content_disposition() == 'attachment':
                    filename = part.get_filename()
                    if filename and (filename.endswith('.xlsx') or filename.endswith('.xls')):
                        # Проверяем, что это файл "для бота" (поиск по ключевым словам)
                        filename_lower = filename.lower()
                        if any(keyword in filename_lower for keyword in ['бот', 'bot', 'xlsx']):
                            print(f"Найден файл: {filename}")
                            return part, email_id
        
        return None
    except Exception as e:
        print(f"Ошибка при поиске вложения: {e}")
        import traceback
        traceback.print_exc()
        return None


def download_attachment(part, save_path):
    """Скачивание вложения"""
    try:
        with open(save_path, 'wb') as f:
            f.write(part.get_payload(decode=True))
        print(f"Файл скачан: {save_path}")
        return True
    except Exception as e:
        print(f"Ошибка при скачивании файла: {e}")
        return False


def update_ostanki_sheet(bot_file_path, main_file_path):
    """Обновление листа 'остатки' в основном файле"""
    try:
        # Читаем данные из файла бота
        bot_df = pd.read_excel(bot_file_path, sheet_name=0)
        print(f"Прочитано строк из файла бота: {len(bot_df)}")
        print(f"Столбцы в файле бота: {list(bot_df.columns)}")
        
        # Загружаем основной файл
        wb = load_workbook(main_file_path)
        
        # Проверяем существование листа 'остатки' или создаем его
        if SHEET_OSTANKI in wb.sheetnames:
            ws = wb[SHEET_OSTANKI]
            ws.delete_rows(1, ws.max_row)  # Удаляем старые данные
        else:
            ws = wb.create_sheet(SHEET_OSTANKI)
        
        # Записываем заголовки
        headers = list(bot_df.columns)
        for col_idx, header in enumerate(headers, start=1):
            ws.cell(row=1, column=col_idx, value=header)
        
        # Записываем данные
        for row_idx, row_data in enumerate(bot_df.values, start=2):
            for col_idx, value in enumerate(row_data, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)
        
        wb.save(main_file_path)
        print(f"Лист '{SHEET_OSTANKI}' обновлен")
        return bot_df
    except Exception as e:
        print(f"Ошибка при обновлении листа остатки: {e}")
        import traceback
        traceback.print_exc()
        return None


def find_catalog_columns(ws):
    """Поиск столбцов с номерами по каталогу"""
    catalog_col = None
    catalog_agb_col = None
    ostanki_col = None
    
    # Ищем заголовки в первой строке
    for col in range(1, ws.max_column + 1):
        cell_value = str(ws.cell(row=1, column=col).value or "").lower()
        # Ищем основной номер по каталогу (без АГБ)
        if 'номер' in cell_value and 'каталог' in cell_value:
            if 'агб' in cell_value:
                catalog_agb_col = col
            else:
                catalog_col = col
        # Ищем столбец остатков
        elif 'остат' in cell_value:
            ostanki_col = col
    
    return catalog_col, catalog_agb_col, ostanki_col


def create_ostanki_dict(ostanki_df):
    """Создание словаря остатков по номерам каталога"""
    ostanki_dict = {}
    
    # Ищем столбцы с номерами каталога и остатками
    catalog_cols = []
    ostanki_value_col = None
    
    # Ищем столбцы с номерами каталога и остатками
    # Приоритет: "Номер" (числовой код), "Номенклатура.код" (код вида АГ-00003268), "Номенклатура" (название)
    
    for col in ostanki_df.columns:
        col_lower = str(col).lower().strip()
        col_exact = str(col).strip()
        
        # Проверяем точное совпадение со столбцом "Номер"
        # Убираем все пробелы для сравнения
        col_clean = col_lower.replace(' ', '').replace('\t', '').replace('\n', '').replace('\r', '')
        
        # Явная проверка столбца "Номер"
        if col_clean == 'номер' or col_exact == 'Номер':
            # Это столбец "Номер" - числовой код
            catalog_cols.insert(0, col)  # Высокий приоритет
        elif 'номенклатура' in col_lower and 'код' in col_lower:
            # Это столбец "Номенклатура.код"
            catalog_cols.append(col)
        elif 'номенклатура' in col_lower and 'код' not in col_lower:
            # Это столбец "Номенклатура" (название)
            catalog_cols.append(col)
        # Ищем столбец с остатками/количеством
        elif 'остат' in col_lower or ('количество' in col_lower and 'код' not in col_lower) or 'кол-во' in col_lower:
            ostanki_value_col = col
    
    if not catalog_cols:
        print(f"Не найдены столбцы с номерами каталога в листе остатки")
        print(f"Столбцы: {list(ostanki_df.columns)}")
        return ostanki_dict
    
    if not ostanki_value_col:
        print(f"Не найден столбец с остатками в листе остатки")
        print(f"Столбцы: {list(ostanki_df.columns)}")
        return ostanki_dict
    
    print(f"Найден столбец с остатками: {ostanki_value_col}")
    print(f"Найдены столбцы с номерами каталога (в порядке приоритета): {catalog_cols}")
    
    # Создаем словарь: номер каталога -> сумма остатков
    # Используем все найденные столбцы для создания словаря
    for _, row in ostanki_df.iterrows():
        ostanki_value = row[ostanki_value_col] if pd.notna(row[ostanki_value_col]) else 0
        try:
            ostanki_value = float(ostanki_value)
        except:
            ostanki_value = 0
        
        if ostanki_value == 0:
            continue
        
        # Проверяем все столбцы с номерами каталога
        for catalog_col in catalog_cols:
            catalog_num = str(row[catalog_col]).strip() if pd.notna(row[catalog_col]) else ""
            # Пропускаем пустые значения и NaN
            if catalog_num and catalog_num.lower() != 'nan' and catalog_num.lower() != 'none':
                # Суммируем остатки для одинаковых номеров каталога
                if catalog_num in ostanki_dict:
                    ostanki_dict[catalog_num] += ostanki_value
                else:
                    ostanki_dict[catalog_num] = ostanki_value
                break  # Используем только первый найденный номер
    
    print(f"Создан словарь остатков для {len(ostanki_dict)} позиций")
    return ostanki_dict


def update_ostanki_in_all_sheets(main_file_path, ostanki_dict):
    """Обновление столбца остатки на всех листах"""
    try:
        wb = load_workbook(main_file_path)
        updated_sheets = []
        
        for sheet_name in wb.sheetnames:
            if sheet_name.lower() == SHEET_OSTANKI.lower():
                continue  # Пропускаем лист остатки
            
            ws = wb[sheet_name]
            catalog_col, catalog_agb_col, ostanki_col = find_catalog_columns(ws)
            
            if not ostanki_col:
                # Если столбца остатки нет, создаем его
                ostanki_col = ws.max_column + 1
                ws.cell(row=1, column=ostanki_col, value="Остатки")
            
            if catalog_col or catalog_agb_col:
                updated_count = 0
                # Обновляем остатки для каждой строки
                for row in range(2, ws.max_row + 1):
                    catalog_num = None
                    
                    # Пробуем взять номер из основного каталога
                    if catalog_col:
                        cell_value = ws.cell(row=row, column=catalog_col).value
                        if cell_value and str(cell_value).strip().lower() != 'none':
                            catalog_num = str(cell_value).strip()
                    
                    # Если не нашли, пробуем из каталога АГБ
                    if not catalog_num and catalog_agb_col:
                        cell_value = ws.cell(row=row, column=catalog_agb_col).value
                        if cell_value and str(cell_value).strip().lower() != 'none':
                            catalog_num = str(cell_value).strip()
                    
                    # Обновляем остатки
                    if catalog_num and catalog_num.lower() != 'nan' and catalog_num:
                        if catalog_num in ostanki_dict:
                            ws.cell(row=row, column=ostanki_col, value=ostanki_dict[catalog_num])
                            updated_count += 1
                        else:
                            # Если номер есть, но остатков нет, ставим 0
                            ws.cell(row=row, column=ostanki_col, value=0)
                
                if updated_count > 0:
                    updated_sheets.append(f"{sheet_name} ({updated_count} строк)")
                    print(f"Обновлен лист '{sheet_name}': {updated_count} строк")
        
        wb.save(main_file_path)
        print(f"Всего обновлено листов: {len(updated_sheets)}")
        return updated_sheets
    except Exception as e:
        print(f"Ошибка при обновлении листов: {e}")
        import traceback
        traceback.print_exc()
        return []


def main(use_local_file=None):
    """Основная функция
    
    Args:
        use_local_file: Если указан путь к локальному файлу, используется он вместо почты
    """
    print("=" * 50)
    print("Начало обновления остатков")
    print("=" * 50)
    
    mail = None
    
    if use_local_file:
        # Используем локальный файл для тестирования
        print(f"\n1. Использование локального файла: {use_local_file}")
        if not os.path.exists(use_local_file):
            print(f"Файл не найден: {use_local_file}")
            return
        import shutil
        shutil.copy(use_local_file, TEMP_BOT_FILE)
        print(f"Файл скопирован в {TEMP_BOT_FILE}")
    else:
        # Подключаемся к почте
        print("\n1. Подключение к почте...")
        mail = connect_to_email()
        if not mail:
            print("Не удалось подключиться к почте")
            print("\nВНИМАНИЕ: Для Gmail может потребоваться пароль приложения!")
            print("Создайте пароль приложения в настройках Google аккаунта:")
            print("https://myaccount.google.com/apppasswords")
            return
        
        # Ищем последнее вложение
        print("\n2. Поиск файла в почте...")
        attachment_data = find_latest_excel_attachment(mail)
        if not attachment_data:
            print("Файл не найден в почте")
            mail.close()
            mail.logout()
            return
        
        part, email_id = attachment_data
        
        # Скачиваем файл
        print("\n3. Скачивание файла...")
        if not download_attachment(part, TEMP_BOT_FILE):
            print("Не удалось скачать файл")
            mail.close()
            mail.logout()
            return
    
    # Обновляем лист остатки
    print("\n4. Обновление листа 'остатки'...")
    ostanki_df = update_ostanki_sheet(TEMP_BOT_FILE, MAIN_FILE)
    if ostanki_df is None:
        print("Не удалось обновить лист остатки")
        if os.path.exists(TEMP_BOT_FILE):
            os.remove(TEMP_BOT_FILE)
        if mail:
            mail.close()
            mail.logout()
        return
    
    # Создаем словарь остатков
    print("\n5. Создание словаря остатков...")
    ostanki_dict = create_ostanki_dict(ostanki_df)
    
    # Обновляем остатки на всех листах
    print("\n6. Обновление остатков на всех листах...")
    updated_sheets = update_ostanki_in_all_sheets(MAIN_FILE, ostanki_dict)
    
    # Удаляем временный файл
    if os.path.exists(TEMP_BOT_FILE):
        os.remove(TEMP_BOT_FILE)
        print(f"\nВременный файл удален: {TEMP_BOT_FILE}")
    
    # Закрываем соединение с почтой
    if mail:
        mail.close()
        mail.logout()
    
    print("\n" + "=" * 50)
    print("Обновление завершено успешно!")
    print("=" * 50)
    if updated_sheets:
        print("\nОбновленные листы:")
        for sheet_info in updated_sheets:
            print(f"  - {sheet_info}")


if __name__ == "__main__":
    import sys
    # Можно запустить с аргументом для использования локального файла
    # python update_ostanki.py "для бота (XLSX).xlsx"
    if len(sys.argv) > 1:
        local_file = sys.argv[1]
        main(use_local_file=local_file)
    else:
        main()

