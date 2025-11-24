#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import imaplib
import email
import os
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from datetime import datetime
import re
import time
import subprocess
import psutil

EMAIL_LOGIN = "almazgeobur.it@mail.ru"
EMAIL_PASSWORD = "Ba9uV5zDx6rE1fs6PgsV"
IMAP_SERVER = "imap.mail.ru"
IMAP_PORT = 993

MAIN_FILE = "Сборка Москва.xlsx"
TEMP_BOT_FILE = "temp_bot_file.xlsx"
SHEET_OSTANKI = "остатки"


def connect_to_email():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_LOGIN, EMAIL_PASSWORD)
        mail.select('inbox')
        return mail
    except Exception as e:
        print(f"Ошибка подключения к почте: {e}")
        return None


def find_latest_excel_attachment(mail):
    try:
        status, messages = mail.search(None, 'UNSEEN')
        if status != 'OK':
            print("Не удалось найти письма")
            return None
        
        email_ids = messages[0].split()
        print(f"Найдено непрочитанных писем: {len(email_ids)}")
        
        if not email_ids:
            print("Непрочитанных писем нет, ищем последние 50 писем...")
            status, messages = mail.search(None, 'ALL')
            if status == 'OK':
                all_ids = messages[0].split()
                email_ids = all_ids[-50:] if len(all_ids) > 50 else all_ids
                print(f"Найдено всего писем: {len(all_ids)}, проверяем последние {len(email_ids)}")
        
        excel_files_found = []
        for email_id in reversed(email_ids):
            try:
                status, msg_data = mail.fetch(email_id, '(RFC822)')
                if status != 'OK':
                    continue
                
                email_body = msg_data[0][1]
                email_message = email.message_from_bytes(email_body)
                
                subject = email_message.get('Subject', 'Без темы')
                print(f"Проверяем письмо: {subject[:50]}...")
                
                for part in email_message.walk():
                    if part.get_content_disposition() == 'attachment':
                        filename = part.get_filename()
                        if filename:
                            try:
                                from email.header import decode_header
                                decoded = decode_header(filename)[0]
                                if decoded[1]:
                                    filename = decoded[0].decode(decoded[1])
                                else:
                                    filename = decoded[0] if isinstance(decoded[0], str) else decoded[0].decode('utf-8')
                            except:
                                pass
                            
                            print(f"  Найдено вложение: {filename}")
                            if filename and (filename.endswith('.xlsx') or filename.endswith('.xls')):
                                excel_files_found.append((filename, part, email_id))
                                filename_lower = filename.lower()
                                if any(keyword in filename_lower for keyword in ['бот', 'bot', 'xlsx']):
                                    print(f"[OK] Найден файл для бота: {filename}")
                                    return part, email_id
            except Exception as e:
                print(f"Ошибка при обработке письма {email_id}: {e}")
                continue
        
        if excel_files_found:
            filename, part, email_id = excel_files_found[-1]
            print(f"Файл 'для бота' не найден, используем последний Excel файл: {filename}")
            return part, email_id
        
        print("Excel файлы не найдены в письмах")
        return None
    except Exception as e:
        print(f"Ошибка при поиске вложения: {e}")
        import traceback
        traceback.print_exc()
        return None


def download_attachment(part, save_path):
    try:
        with open(save_path, 'wb') as f:
            f.write(part.get_payload(decode=True))
        print(f"Файл скачан: {save_path}")
        return True
    except Exception as e:
        print(f"Ошибка при скачивании файла: {e}")
        return False


def update_ostanki_sheet(bot_file_path, main_file_path):
    try:
        bot_df = pd.read_excel(bot_file_path, sheet_name=0)
        print(f"Прочитано строк из файла бота: {len(bot_df)}")
        print(f"Столбцы в файле бота: {list(bot_df.columns)}")
        
        for col in bot_df.columns:
            col_str = str(col).strip()
            if 'артикул' in col_str.lower():
                print(f"Нормализация столбца '{col}' (удаление пробелов)")
                bot_df[col] = bot_df[col].astype(str).str.replace(' ', '').str.replace('\t', '').str.replace('\n', '').str.replace('\r', '')
                break
        
        wb = load_workbook(main_file_path)
        
        existing_sheet = None
        sheets_to_remove = []
        
        for sheet_name in wb.sheetnames:
            sheet_lower = sheet_name.lower().strip()
            if sheet_lower == SHEET_OSTANKI.lower():
                existing_sheet = sheet_name
            elif sheet_lower.startswith(SHEET_OSTANKI.lower()) and sheet_lower != SHEET_OSTANKI.lower():
                sheets_to_remove.append(sheet_name)
        
        for sheet_name in sheets_to_remove:
            print(f"Удаление дубликата листа: '{sheet_name}'")
            wb.remove(wb[sheet_name])
        
        if existing_sheet:
            ws = wb[existing_sheet]
            if existing_sheet != SHEET_OSTANKI:
                ws.title = SHEET_OSTANKI
            ws.delete_rows(1, ws.max_row)
            print(f"Обновление существующего листа '{SHEET_OSTANKI}'")
        else:
            ws = wb.create_sheet(SHEET_OSTANKI)
            print(f"Создание нового листа '{SHEET_OSTANKI}'")
        
        headers = list(bot_df.columns)
        for col_idx, header in enumerate(headers, start=1):
            ws.cell(row=1, column=col_idx, value=header)
        
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
    catalog_col = None
    catalog_agb_col = None
    ostanki_col = None
    kolvo_col = None
    
    for col in range(1, ws.max_column + 1):
        cell_value = str(ws.cell(row=1, column=col).value or "").lower()
        if 'номер' in cell_value and 'каталог' in cell_value:
            if 'агб' in cell_value:
                catalog_agb_col = col
            else:
                catalog_col = col
        elif 'кол-во' in cell_value and 'станок' in cell_value:
            kolvo_col = col
        elif 'остат' in cell_value and 'склад' in cell_value:
            ostanki_col = col
    
    if kolvo_col and not ostanki_col:
        ostanki_col = kolvo_col + 1
        ws.insert_cols(ostanki_col)
        ws.cell(row=1, column=ostanki_col, value="Остатки на складах")
    
    return catalog_col, catalog_agb_col, ostanki_col


def create_ostanki_dict(ostanki_df):
    ostanki_dict = {}
    
    catalog_cols = []
    ostanki_value_col = None
    
    for col in ostanki_df.columns:
        col_lower = str(col).lower().strip()
        col_exact = str(col).strip()
        
        if (col_exact == 'Номер' or 
            col_lower == 'номер' or 
            (col_lower.startswith('номер') and len(col_lower) <= 10 and 'каталог' not in col_lower and 'код' not in col_lower)):
            if col not in catalog_cols:
                catalog_cols.insert(0, col)
        elif 'номенклатура' in col_lower and 'код' in col_lower:
            if col not in catalog_cols:
                catalog_cols.append(col)
        elif 'номенклатура' in col_lower and 'код' not in col_lower:
            if col not in catalog_cols:
                catalog_cols.append(col)
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
    
    number_col = None
    for i, col in enumerate(ostanki_df.columns):
        col_str = str(col).strip()
        if (col_str == 'Номер' or 
            col_str.lower().replace(' ', '') == 'номер' or
            (i == 1 and 'номер' in col_str.lower() and 'каталог' not in col_str.lower())):
            number_col = col
            if col not in catalog_cols:
                catalog_cols.insert(0, col)
            break
    
    if number_col:
        print(f"[OK] Найден критически важный столбец 'Номер': {number_col}")
    else:
        if len(ostanki_df.columns) > 1:
            potential_num_col = ostanki_df.columns[1]
            print(f"[INFO] Пробуем использовать столбец по позиции [1]: '{potential_num_col}'")
            if potential_num_col not in catalog_cols:
                catalog_cols.insert(0, potential_num_col)
                number_col = potential_num_col
    
    print(f"Найдены столбцы с номерами каталога (в порядке приоритета): {catalog_cols}")
    
    artikul_col = None
    for col in ostanki_df.columns:
        if 'артикул' in str(col).lower():
            artikul_col = col
            print(f"Найден столбец 'Артикул': {artikul_col}, нормализация...")
            ostanki_df[col] = ostanki_df[col].astype(str).str.replace(' ', '').str.replace('\t', '').str.replace('\n', '').str.replace('\r', '')
            break
    
    for _, row in ostanki_df.iterrows():
        ostanki_value = row[ostanki_value_col] if pd.notna(row[ostanki_value_col]) else 0
        try:
            ostanki_value = float(ostanki_value)
        except:
            ostanki_value = 0
        
        if ostanki_value == 0:
            continue
        
        catalog_num = None
        for catalog_col in catalog_cols:
            val = row[catalog_col]
            if pd.notna(val):
                catalog_num = str(val).strip()
                catalog_num = catalog_num.replace(' ', '').replace('\t', '').replace('\n', '').replace('\r', '').lower()
                if catalog_num and catalog_num not in ['nan', 'none', '']:
                    if catalog_num in ostanki_dict:
                        ostanki_dict[catalog_num] += ostanki_value
                    else:
                        ostanki_dict[catalog_num] = ostanki_value
                    break
    
    print(f"Создан словарь остатков для {len(ostanki_dict)} позиций")
    return ostanki_dict


def close_file_sessions(file_path):
    try:
        file_path_abs = os.path.abspath(file_path)
        print(f"\nПроверка открытых сеансов файла: {file_path_abs}")
        
        processes_to_close = []
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                proc_info = proc.info
                proc_name = proc_info['name'].lower()
                
                if 'p7' in proc_name or 'excel' in proc_name or 'spreadsheet' in proc_name or 'calc' in proc_name:
                    try:
                        open_files = proc.open_files()
                        for file_info in open_files:
                            if file_info and hasattr(file_info, 'path'):
                                try:
                                    if os.path.abspath(file_info.path) == file_path_abs:
                                        processes_to_close.append((proc_info['pid'], proc_info['name']))
                                        break
                                except:
                                    pass
                    except (psutil.AccessDenied, psutil.NoSuchProcess):
                        pass
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
        
        if processes_to_close:
            print(f"Найдено {len(processes_to_close)} процессов, открывающих файл:")
            for pid, name in processes_to_close:
                print(f"  - PID {pid}: {name}")
            
            print("Закрытие процессов...")
            for pid, name in processes_to_close:
                try:
                    proc = psutil.Process(pid)
                    proc.terminate()
                    print(f"  Закрыт процесс {name} (PID {pid})")
                except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
                    print(f"  Не удалось закрыть процесс {name} (PID {pid}): {e}")
            
            time.sleep(2)
            
            for pid, name in processes_to_close:
                try:
                    proc = psutil.Process(pid)
                    if proc.is_running():
                        proc.kill()
                        print(f"  Принудительно закрыт процесс {name} (PID {pid})")
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            
            time.sleep(1)
        else:
            print("Файл не открыт в других процессах")
        
        max_wait = 30
        wait_time = 0
        while wait_time < max_wait:
            try:
                with open(file_path, 'r+b') as f:
                    pass
                print("Файл освобожден и готов к обновлению")
                return True
            except (PermissionError, IOError):
                wait_time += 1
                if wait_time % 5 == 0:
                    print(f"Ожидание освобождения файла... ({wait_time}/{max_wait} сек)")
                time.sleep(1)
        
        print(f"Предупреждение: файл не освобожден за {max_wait} секунд, продолжаем...")
        return False
    except Exception as e:
        print(f"Ошибка при закрытии сеансов файла: {e}")
        import traceback
        traceback.print_exc()
        return False


def update_ostanki_in_all_sheets(main_file_path, ostanki_dict):
    try:
        wb = load_workbook(main_file_path)
        updated_sheets = []
        
        for sheet_name in wb.sheetnames:
            if sheet_name.lower() == SHEET_OSTANKI.lower():
                continue
            
            ws = wb[sheet_name]
            catalog_col, catalog_agb_col, ostanki_col = find_catalog_columns(ws)
            
            if not ostanki_col:
                ostanki_col = ws.max_column + 1
                ws.cell(row=1, column=ostanki_col, value="Остатки на складах")
            
            if catalog_col or catalog_agb_col:
                updated_count = 0
                for row in range(2, ws.max_row + 1):
                    catalog_num = None
                    
                    if catalog_col:
                        cell_value = ws.cell(row=row, column=catalog_col).value
                        if cell_value and str(cell_value).strip().lower() != 'none':
                            catalog_num = str(cell_value).strip()
                    
                    if not catalog_num and catalog_agb_col:
                        cell_value = ws.cell(row=row, column=catalog_agb_col).value
                        if cell_value and str(cell_value).strip().lower() != 'none':
                            catalog_num = str(cell_value).strip()
                    
                    if catalog_num and catalog_num.lower() not in ['nan', 'none', '']:
                        catalog_num_clean = str(catalog_num).strip().replace(' ', '').replace('\t', '').replace('\n', '').replace('\r', '').lower()
                        
                        found_value = None
                        if catalog_num_clean in ostanki_dict:
                            found_value = ostanki_dict[catalog_num_clean]
                        else:
                            for key in ostanki_dict.keys():
                                key_normalized = str(key).replace(' ', '').replace('\t', '').replace('\n', '').replace('\r', '').lower()
                                if catalog_num_clean in key_normalized or key_normalized in catalog_num_clean:
                                    found_value = ostanki_dict[key]
                                    break
                        
                        try:
                            cell = ws.cell(row=row, column=ostanki_col)
                            if hasattr(cell, 'value') and hasattr(cell, 'coordinate'):
                                if found_value is not None:
                                    cell.value = found_value
                                    updated_count += 1
                                else:
                                    cell.value = 0
                        except AttributeError:
                            continue
                
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
    print("=" * 50)
    print("Начало обновления остатков")
    print("=" * 50)
    
    mail = None
    
    if use_local_file:
        print(f"\n1. Использование локального файла: {use_local_file}")
        if not os.path.exists(use_local_file):
            print(f"Файл не найден: {use_local_file}")
            return
        import shutil
        shutil.copy(use_local_file, TEMP_BOT_FILE)
        print(f"Файл скопирован в {TEMP_BOT_FILE}")
    else:
        print("\n1. Подключение к почте...")
        mail = connect_to_email()
        if not mail:
            print("Не удалось подключиться к почте")
            return
        
        print("\n2. Поиск файла в почте...")
        attachment_data = find_latest_excel_attachment(mail)
        if not attachment_data:
            print("Файл не найден в почте")
            mail.close()
            mail.logout()
            return
        
        part, email_id = attachment_data
        
        print("\n3. Скачивание файла...")
        if not download_attachment(part, TEMP_BOT_FILE):
            print("Не удалось скачать файл")
            mail.close()
            mail.logout()
            return
    
    print("\n4. Закрытие сеансов файла перед обновлением...")
    close_file_sessions(MAIN_FILE)
    
    print("\n5. Обновление листа 'остатки'...")
    ostanki_df = update_ostanki_sheet(TEMP_BOT_FILE, MAIN_FILE)
    if ostanki_df is None:
        print("Не удалось обновить лист остатки")
        if os.path.exists(TEMP_BOT_FILE):
            os.remove(TEMP_BOT_FILE)
        if mail:
            mail.close()
            mail.logout()
        return
    
    print("\n6. Создание словаря остатков...")
    ostanki_dict = create_ostanki_dict(ostanki_df)
    
    print("\n7. Обновление остатков на всех листах...")
    close_file_sessions(MAIN_FILE)
    updated_sheets = update_ostanki_in_all_sheets(MAIN_FILE, ostanki_dict)
    
    if os.path.exists(TEMP_BOT_FILE):
        os.remove(TEMP_BOT_FILE)
        print(f"\nВременный файл удален: {TEMP_BOT_FILE}")
    
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
    if len(sys.argv) > 1:
        local_file = sys.argv[1]
        main(use_local_file=local_file)
    else:
        main()

