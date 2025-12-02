#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Тест импортов для диагностики проблем запуска"""

import sys

print("Проверка импортов...")
print(f"Python версия: {sys.version}")
print()

errors = []

# Проверяем стандартные библиотеки
try:
    import sqlite3
    print("✓ sqlite3")
except ImportError as e:
    errors.append(f"✗ sqlite3: {e}")
    print(f"✗ sqlite3: {e}")

try:
    import tkinter as tk
    print("✓ tkinter")
except ImportError as e:
    errors.append(f"✗ tkinter: {e}")
    print(f"✗ tkinter: {e}")

try:
    import pandas as pd
    print("✓ pandas")
except ImportError as e:
    errors.append(f"✗ pandas: {e}")
    print(f"✗ pandas: {e}")

try:
    from tkcalendar import DateEntry
    print("✓ tkcalendar")
except ImportError as e:
    errors.append(f"✗ tkcalendar: {e}")
    print(f"✗ tkcalendar: {e}")

try:
    import ctypes
    print("✓ ctypes")
except ImportError as e:
    errors.append(f"✗ ctypes: {e}")
    print(f"✗ ctypes: {e}")

try:
    from pathlib import Path
    print("✓ pathlib")
except ImportError as e:
    errors.append(f"✗ pathlib: {e}")
    print(f"✗ pathlib: {e}")

print()
if errors:
    print("ОШИБКИ ИМПОРТА:")
    for error in errors:
        print(f"  {error}")
    print()
    print("Решение: Установите недостающие модули:")
    print("  pip install -r requirements.txt")
else:
    print("Все импорты успешны!")


