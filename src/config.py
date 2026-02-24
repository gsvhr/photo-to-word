"""Все настройки приложения в одном месте."""

import os
import json
from dataclasses import dataclass, field
from typing import List, Tuple, Dict, Any
from pathlib import Path


@dataclass
class AppConfig:
    """Главный класс конфигурации приложения."""

    # Настройки окна
    window_title: str = "Фототаблица"
    window_min_width: int = 800
    window_min_height: int = 600

    # Настройки изображений
    supported_formats: List[str] = field(
        default_factory=lambda: [
            ".jpg",
            ".jpeg",
            ".png",
            ".bmp",
            ".gif",
            ".tiff",
            ".webp",
            ".avif",
        ]
    )
    thumbnail_size: Tuple[int, int] = (150, 150)

    # Настройки генерации Word
    word_document_defaults: Dict[str, Any] = field(
        default_factory=lambda: {
            "page_margins": {
                "top": 1.0,  # см
                "left": 1.0,  # см
                "right": 1.0,  # см
                "bottom": 1.0,  # см
            },
            "font_name": "Times New Roman",
            "font_size": 10,
            "image_width_portrait": 400,  # пикселей
            "image_width_landscape": 600,  # пикселей (одинаково для обеих ориентаций)
            "jpeg_quality": 85,  # %
            "table_width_portrait": 16,  # см
            "table_width_landscape": 24,  # см
            "caption_template": "Рис. {number}. {filename}",
        }
    )

    # Настройки интерфейса
    colors: Dict[str, str] = field(
        default_factory=lambda: {
            "status_ready": "#2ecc71",  # зеленый
            "status_processing": "#f39c12",  # оранжевый
            "status_success": "#27ae60",  # зеленый
            "status_error": "#e74c3c",  # красный
            "thumbnail_hover": "#f0f0f0",  # светло-серый
            "thumbnail_selected": "#e3f2fd",  # голубой
        }
    )

    # Тексты статусов
    status_messages: Dict[str, str] = field(
        default_factory=lambda: {
            "ready": "Готов к работе",
            "processing": "Обработка {current} из {total}",
            "success": "Готово! Файл сохранен",
            "error": "Ошибка!",
        }
    )

    def __post_init__(self):
        """Инициализация после создания."""
        # Определяем путь к папке с настройками
        self.config_dir = Path(os.environ.get("APPDATA", "")) / "PhotoTable"
        self.config_file = self.config_dir / "config.json"

        # Создаем папку если её нет
        self.config_dir.mkdir(parents=True, exist_ok=True)

        # Загружаем настройки
        self.user_settings = self.load_settings()

    def load_settings(self) -> dict:
        """Загрузить настройки из файла."""
        default_settings = {
            "last_path": str(Path.home()),
            "window_geometry": {},
            "table_width_portrait": 16.0,
            "table_width_landscape": 25.0,
            "jpeg_quality": 85,
            "orientation": "portrait",
        }

        if self.config_file.exists():
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                    # Обновляем только существующие ключи
                    for key in default_settings.keys():
                        if key in saved:
                            default_settings[key] = saved[key]
                print(f"Настройки загружены из: {self.config_file}")
            except Exception as e:
                print(f"Ошибка загрузки настроек: {e}")

        return default_settings

    def save_settings(self):
        """Сохранить настройки в файл."""
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(self.user_settings, f, indent=2, ensure_ascii=False)
            print(f"Настройки сохранены в: {self.config_file}")
        except Exception as e:
            print(f"Ошибка сохранения настроек: {e}")

    # Методы для работы с настройками
    def get_last_path(self) -> str:
        """Получить последний использованный путь."""
        return self.user_settings.get("last_path", str(Path.home()))

    def set_last_path(self, path: str):
        """Сохранить последний использованный путь."""
        self.user_settings["last_path"] = path
        self.save_settings()

    def get_window_geometry(self) -> dict:
        """Получить геометрию окна."""
        return self.user_settings.get("window_geometry", {})

    def set_window_geometry(self, geometry: dict):
        """Сохранить геометрию окна."""
        self.user_settings["window_geometry"] = geometry
        self.save_settings()

    def get_table_width(self, orientation: str) -> float:
        """Получить ширину таблицы для ориентации."""
        key = (
            "table_width_portrait"
            if orientation == "portrait"
            else "table_width_landscape"
        )
        return self.user_settings.get(key, 16.0 if orientation == "portrait" else 25.0)

    def set_table_width(self, orientation: str, width: float):
        """Сохранить ширину таблицы для ориентации."""
        key = (
            "table_width_portrait"
            if orientation == "portrait"
            else "table_width_landscape"
        )
        self.user_settings[key] = width
        self.save_settings()

    def get_jpeg_quality(self) -> int:
        """Получить качество JPEG."""
        return self.user_settings.get("jpeg_quality", 85)

    def set_jpeg_quality(self, quality: int):
        """Сохранить качество JPEG."""
        self.user_settings["jpeg_quality"] = quality
        self.save_settings()

    def get_orientation(self) -> str:
        """Получить последнюю использованную ориентацию."""
        return self.user_settings.get("orientation", "portrait")

    def set_orientation(self, orientation: str):
        """Сохранить последнюю использованную ориентацию."""
        self.user_settings["orientation"] = orientation
        self.save_settings()

    def get_image_width(self, orientation: str) -> int:
        """Получить ширину изображения в зависимости от ориентации."""
        return self.word_document_defaults["image_width_portrait"]

    def get_columns_count(self, orientation: str) -> int:
        """Получить количество колонок (всегда 2 для обеих ориентаций)."""
        return 2

    def get_rows_per_page(self, orientation: str) -> int:
        """Получить количество строк на странице."""
        if orientation == "landscape":
            return 2  # Альбомная: 2 строки = 4 фото
        return 4  # Портретная: 4 строки = 8 фото

    def get_caption_text(self, number: int, filename: str) -> str:
        """Сформировать текст подписи по шаблону с ограничением длины."""
        template = self.word_document_defaults["caption_template"]

        # Ограничиваем имя файла до 35 символов
        if len(filename) > 35:
            filename = filename[:32] + "..."

        return template.format(number=number, filename=filename)
