"""Обработка изображений: загрузка, создание миниатюр, изменение размера."""

import os
from typing import List, Tuple, Optional, Dict
from pathlib import Path
from PIL import Image, ImageOps, UnidentifiedImageError
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class ImageProcessor:
    """Класс для обработки изображений."""

    def __init__(self, config):
        self.config = config
        self.images = []  # Список путей к загруженным изображениям
        self.thumbnails = []  # Список миниатюр
        self.rotations = (
            []
        )  # Список углов поворота для каждого изображения (0, 90, 180, 270)
        self.original_sizes = []  # Оригинальные размеры для определения ориентации

    def load_images(self, file_paths: List[str]) -> List[str]:
        """Загрузить изображения из файлов."""
        loaded = []
        for file_path in file_paths:
            if self._is_supported(file_path):
                try:
                    # Проверяем, что файл действительно изображение
                    with Image.open(file_path) as img:
                        img.verify()

                    if file_path not in self.images:
                        self.images.append(file_path)
                        self.rotations.append(0)  # Начальный поворот 0
                        loaded.append(file_path)
                        logger.info(f"Загружено: {file_path}")
                except (UnidentifiedImageError, IOError) as e:
                    logger.error(f"Ошибка загрузки {file_path}: {e}")

        # Создаем миниатюры для новых изображений
        if loaded:
            self._create_thumbnails()

        return loaded

    def _is_supported(self, file_path: str) -> bool:
        """Проверить, поддерживается ли формат файла."""
        ext = Path(file_path).suffix.lower()
        return ext in self.config.supported_formats

    def _create_thumbnails(self):
        """Создать миниатюры для всех изображений с учетом поворота."""
        self.thumbnails = []
        self.original_sizes = []

        for i, img_path in enumerate(self.images):
            try:
                with Image.open(img_path) as img:
                    # Сохраняем оригинальный размер для определения ориентации
                    self.original_sizes.append(img.size)

                    # Применяем поворот если нужно
                    if self.rotations[i] != 0:
                        img = img.rotate(self.rotations[i], expand=True)

                    # Создаем миниатюру с сохранением пропорций
                    img.thumbnail(self.config.thumbnail_size, Image.Resampling.LANCZOS)

                    # Создаем новое изображение с белым фоном нужного размера
                    thumb = Image.new("RGB", self.config.thumbnail_size, "white")

                    # Вставляем изображение по центру
                    offset_x = (self.config.thumbnail_size[0] - img.size[0]) // 2
                    offset_y = (self.config.thumbnail_size[1] - img.size[1]) // 2
                    thumb.paste(img, (offset_x, offset_y))

                    self.thumbnails.append(thumb)
                    logger.info(f"Создана миниатюра для: {img_path}")
            except Exception as e:
                logger.error(f"Ошибка создания миниатюры для {img_path}: {e}")
                # Добавляем заглушку
                thumb = Image.new("RGB", self.config.thumbnail_size, "#cccccc")
                self.thumbnails.append(thumb)
                self.original_sizes.append((100, 100))  # Заглушка

    def rotate_image(self, index: int) -> bool:
        """Повернуть изображение на 90 градусов по часовой стрелке."""
        if 0 <= index < len(self.images):
            # Увеличиваем угол поворота на 90 градусов
            self.rotations[index] = (self.rotations[index] + 90) % 360
            # Пересоздаем миниатюры
            self._create_thumbnails()
            logger.info(
                f"Изображение {index} повернуто на {self.rotations[index]} градусов"
            )
            return True
        return False

    def remove_images(self, indices: List[int]):
        """Удалить изображения по индексам."""
        # Удаляем в обратном порядке, чтобы не сбивались индексы
        for idx in sorted(indices, reverse=True):
            if 0 <= idx < len(self.images):
                del self.images[idx]
                del self.thumbnails[idx]
                del self.rotations[idx]
                del self.original_sizes[idx]

        logger.info(f"Удалено {len(indices)} изображений")

    def clear_all(self):
        """Очистить все изображения."""
        self.images.clear()
        self.thumbnails.clear()
        self.rotations.clear()
        self.original_sizes.clear()
        logger.info("Все изображения очищены")

    def process_for_word(
        self, image_path: str, target_width: int, quality: int, rotation: int = 0
    ) -> Image.Image:
        """Обработать изображение для вставки в Word с учетом поворота."""
        with Image.open(image_path) as img:
            # Удаляем EXIF и другие метаданные, конвертируем в RGB
            if img.mode in ("RGBA", "LA", "P"):
                # Конвертируем в RGB с белым фоном для прозрачности
                background = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == "P":
                    img = img.convert("RGBA")
                if img.mode == "RGBA":
                    background.paste(img, mask=img.split()[3])
                else:
                    background.paste(img)
                img = background
            else:
                img = img.convert("RGB")

            # Применяем поворот если нужно
            if rotation != 0:
                img = img.rotate(rotation, expand=True)

            # Изменяем размер с сохранением пропорций
            ratio = target_width / img.size[0]
            target_height = int(img.size[1] * ratio)
            img = img.resize((target_width, target_height), Image.Resampling.LANCZOS)

            return img

    def get_image_count(self) -> int:
        """Получить количество загруженных изображений."""
        return len(self.images)

    def get_image_paths(self) -> List[str]:
        """Получить список путей к изображениям."""
        return self.images.copy()

    def get_thumbnail(self, index: int) -> Optional[Image.Image]:
        """Получить миниатюру по индексу."""
        if 0 <= index < len(self.thumbnails):
            return self.thumbnails[index]
        return None

    def get_rotation(self, index: int) -> int:
        """Получить угол поворота для изображения."""
        if 0 <= index < len(self.rotations):
            return self.rotations[index]
        return 0

    def get_original_size(self, index: int) -> Tuple[int, int]:
        """Получить оригинальный размер изображения."""
        if 0 <= index < len(self.original_sizes):
            return self.original_sizes[index]
        return (100, 100)  # Значение по умолчанию

    def is_landscape(self, index: int) -> bool:
        """Проверить, является ли изображение ландшафтным (ширина > высоты)."""
        width, height = self.get_original_size(index)
        return width > height
