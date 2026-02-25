#!/usr/bin/env python3
"""Точка входа в приложение."""

import os
import tkinter as tk
from src.gui import PhotoTableApp


def main():
    """Основная функция запуска приложения."""
    root = tk.Tk()

    # Путь к иконке
    icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")

    # Устанавливаем иконку если файл существует
    if os.path.exists(icon_path):
        try:
            root.iconbitmap(default=icon_path)
        except Exception as e:
            print(f"Не удалось загрузить иконку: {e}")

    app = PhotoTableApp(root)  # pylint: disable=unused-variable
    root.mainloop()


if __name__ == "__main__":
    main()
