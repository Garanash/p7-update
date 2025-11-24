#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import schedule
import time
import subprocess
import sys
import os
from datetime import datetime

def run_update():
    print(f"\n{'='*60}")
    print(f"Запуск обновления остатков: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*60}\n")
    
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        script_path = os.path.join(script_dir, "update_ostanki.py")
        
        result = subprocess.run(
            [sys.executable, script_path],
            cwd=script_dir,
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        
        print(result.stdout)
        if result.stderr:
            print("Ошибки:", result.stderr)
        
        print(f"\nЗавершено: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"{'='*60}\n")
        
    except Exception as e:
        print(f"Ошибка при запуске скрипта: {e}")
        import traceback
        traceback.print_exc()


def main():
    print("=" * 60)
    print("Планировщик обновления остатков")
    print("=" * 60)
    print(f"Запуск каждый день в 19:20")
    print(f"Текущее время: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("Для остановки нажмите Ctrl+C")
    print("=" * 60)
    
    schedule.every().day.at("19:20").do(run_update)
    
    while True:
        schedule.run_pending()
        time.sleep(60)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nПланировщик остановлен пользователем")
        sys.exit(0)

