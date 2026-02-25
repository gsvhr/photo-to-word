"""
Скрипт для сборки приложения в один EXE-файл.
Установите зависимости: pip install -r requirements.txt
Запуск: python build_exe.py
"""

import os
import PyInstaller.__main__


def build():
    """Сборка приложения."""

    # Проверяем наличие иконки
    icon_path = "icon.ico"
    if not os.path.exists(icon_path):
        print("ВНИМАНИЕ: Файл icon.ico не найден. Приложение будет без иконки.")
        icon_path = None

    # Базовые параметры сборки
    args = [
        "main.py",
        "--name=Фототаблица 1.0",
        "--onefile",
        "--windowed",
        "--add-data=src;src",
        "--clean",
        "--noconfirm",
    ]

    # Добавляем иконку если есть
    if icon_path:
        args.append(f"--icon={icon_path}")

    # Добавляем метаданные если есть файл версии
    if os.path.exists("version.txt"):
        args.append("--version-file=version.txt")

    print("Начинаем сборку...")
    print(f"Команда: pyinstaller {' '.join(args)}")

    # Запускаем сборку
    PyInstaller.__main__.run(args)

    print("\n" + "=" * 50)
    print("Сборка завершена!")
    print(f"EXE-файл находится в папке: {os.path.abspath('dist')}")
    print("=" * 50)


if __name__ == "__main__":
    build()
