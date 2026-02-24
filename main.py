#!/usr/bin/env python3
"""
Фототаблица - приложение для создания таблиц с фотографиями в Word.
"""

import sys
import os

# Добавляем путь к проекту
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from src.gui import PhotoTableApp


def main():
    """Точка входа в приложение."""
    root = tk.Tk()
    app = PhotoTableApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
