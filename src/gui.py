"""Графический интерфейс приложения."""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
import logging
from datetime import datetime
from pathlib import Path
from PIL import ImageTk  # type: ignore

from src.config import AppConfig
from src.image_processor import ImageProcessor
from src.word_generator import WordGenerator

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class ThumbnailTile(ttk.Frame):
    """Отдельная плитка с фото, чекбоксом, номером и кнопкой поворота."""

    def __init__(
        self,
        parent,
        index,
        image_path,
        thumbnail,
        on_select,
        on_hover,
        on_rotate,
        config,
        image_processor,
    ):
        super().__init__(parent)
        self.index = index
        self.image_path = image_path
        self.config = config
        self.image_processor = image_processor
        self.selected = tk.BooleanVar(value=False)
        self.on_select = on_select
        self.on_hover = on_hover
        self.on_rotate = on_rotate
        self.tooltip = None
        self.tooltip_after_id = None
        self.root = self.winfo_toplevel()

        # Контейнер для миниатюры
        self.image_frame = ttk.Frame(self, relief="solid", borderwidth=1)
        self.image_frame.pack(padx=2, pady=2, fill="both", expand=True)

        # Миниатюра
        self.photo = ImageTk.PhotoImage(thumbnail)
        self.image_label = ttk.Label(self.image_frame, image=self.photo)
        self.image_label.pack()

        # Индикатор ориентации (в левом верхнем углу)
        self.create_orientation_indicator()

        # Верхняя правая панель с номером и чекбоксом
        self.top_frame = ttk.Frame(self.image_frame)
        self.top_frame.place(relx=1.0, rely=0.0, anchor="ne")

        # Номер фото
        self.number_label = ttk.Label(
            self.top_frame,
            text=f"#{index + 1}",
            font=("Arial", 8, "bold"),
            foreground="white",
            background="#3498db",
        )
        self.number_label.pack(side="left", padx=2)

        # Чекбокс
        self.checkbox = ttk.Checkbutton(
            self.top_frame, variable=self.selected, command=self._on_checkbox_click
        )
        self.checkbox.pack(side="left")

        # Кнопка поворота (в правом нижнем углу)
        self.create_rotate_button()

        # Привязка событий
        self.image_label.bind("<Enter>", self._on_enter)
        self.image_label.bind("<Leave>", self._on_leave)
        self.image_label.bind("<Button-1>", self._on_click)

    def create_orientation_indicator(self):
        """Создать индикатор ориентации в левом верхнем углу."""
        self.orientation_frame = tk.Frame(self.image_frame, bg="white")
        self.orientation_frame.place(relx=0.0, rely=0.0, anchor="nw")

        # Определяем ориентацию
        is_landscape = self.image_processor.is_landscape(self.index)

        # Создаем иконку в зависимости от ориентации
        if is_landscape:
            # Горизонтальная иконка (широкий прямоугольник)
            icon_canvas = tk.Canvas(
                self.orientation_frame,
                width=16,
                height=12,
                highlightthickness=0,
                bg="white",
            )
            icon_canvas.create_rectangle(
                2, 2, 14, 10, outline="#27ae60", fill="#2ecc71", width=1
            )
            icon_canvas.create_text(8, 6, text="↔", font=("Arial", 8), fill="white")
        else:
            # Вертикальная иконка (высокий прямоугольник)
            icon_canvas = tk.Canvas(
                self.orientation_frame,
                width=12,
                height=16,
                highlightthickness=0,
                bg="white",
            )
            icon_canvas.create_rectangle(
                2, 2, 10, 14, outline="#e67e22", fill="#f39c12", width=1
            )
            icon_canvas.create_text(6, 8, text="↕", font=("Arial", 8), fill="white")

        icon_canvas.pack()

        # Добавляем статический тултип
        self.create_static_tooltip(
            icon_canvas, "Горизонтальное фото" if is_landscape else "Вертикальное фото"
        )

    def create_rotate_button(self):
        """Создать кнопку поворота в правом нижнем углу."""
        self.rotate_frame = tk.Frame(self.image_frame, bg="white")
        self.rotate_frame.place(relx=1.0, rely=1.0, anchor="se")

        # Создаем кнопку с иконкой поворота
        rotate_canvas = tk.Canvas(
            self.rotate_frame,
            width=20,
            height=20,
            highlightthickness=0,
            bg="white",
            cursor="hand2",
        )

        # Рисуем иконку поворота (круг со стрелкой)
        rotate_canvas.create_oval(
            3, 3, 17, 17, outline="#3498db", fill="#ecf0f1", width=1
        )
        rotate_canvas.create_text(
            10, 10, text="↻", font=("Arial", 12, "bold"), fill="#3498db"
        )

        # Добавляем индикатор текущего поворота
        rotation = self.image_processor.get_rotation(self.index)
        if rotation > 0:
            # Показываем маленький индикатор
            rotate_canvas.create_oval(14, 2, 18, 6, fill="#e74c3c", outline="")
            rotate_canvas.create_text(
                16, 4, text=f"{rotation//90}", font=("Arial", 6), fill="white"
            )

        rotate_canvas.pack()

        # Привязываем события для динамического тултипа
        rotate_canvas.bind(
            "<Enter>",
            lambda e: self._show_tooltip(e, "Повернуть на 90° (по часовой стрелке)"),
        )
        rotate_canvas.bind("<Leave>", self._hide_tooltip)

        # Привязываем событие клика
        rotate_canvas.bind("<Button-1>", self._on_rotate_click)

    def create_static_tooltip(self, widget, text):
        """Создать простой статический тултип без задержки."""

        def enter(event):
            self.tooltip = tk.Toplevel()
            self.tooltip.wm_overrideredirect(True)

            x = event.x_root + 10
            y = event.y_root + 10
            self.tooltip.wm_geometry(f"+{x}+{y}")

            label = ttk.Label(
                self.tooltip,
                text=text,
                background="#ffffe0",
                relief="solid",
                borderwidth=1,
                padding=2,
            )
            label.pack()

        def leave(event):
            if self.tooltip:
                self.tooltip.destroy()
                self.tooltip = None

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def _show_tooltip(self, event, text):
        """Показать всплывающую подсказку с задержкой."""
        if self.tooltip_after_id:
            self.root.after_cancel(self.tooltip_after_id)

        self.tooltip_text = text
        self.tooltip_x = event.x_root
        self.tooltip_y = event.y_root

        self.tooltip_after_id = self.root.after(500, self._create_dynamic_tooltip)

    def _create_dynamic_tooltip(self):
        """Создать динамический тултип."""
        self._hide_tooltip()

        self.tooltip = tk.Toplevel()
        self.tooltip.wm_overrideredirect(True)

        x = self.tooltip_x + 10
        y = self.tooltip_y + 10

        screen_width = self.tooltip.winfo_screenwidth()
        screen_height = self.tooltip.winfo_screenheight()

        if x + 150 > screen_width:
            x = self.tooltip_x - 160
        if y + 30 > screen_height:
            y = self.tooltip_y - 40

        self.tooltip.wm_geometry(f"+{x}+{y}")

        label = ttk.Label(
            self.tooltip,
            text=self.tooltip_text,
            background="#ffffe0",
            relief="solid",
            borderwidth=1,
            padding=2,
        )
        label.pack()

        self.tooltip_after_id = None

    def _hide_tooltip(self, event=None):
        """Скрыть всплывающую подсказку."""
        if self.tooltip_after_id:
            self.root.after_cancel(self.tooltip_after_id)
            self.tooltip_after_id = None

        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

    def _on_rotate_click(self, event):
        """Обработчик клика по кнопке поворота."""
        self._hide_tooltip()
        self.on_rotate(self.index)

    def destroy(self):
        """При уничтожении виджета удаляем тултип."""
        self._hide_tooltip()
        super().destroy()

    def _on_checkbox_click(self):
        """Обработчик клика по чекбоксу."""
        self.on_select(self.index, self.selected.get())

    def _on_enter(self, event):
        """При наведении мыши."""
        self.image_label.config(background=self.config.colors["thumbnail_hover"])
        self.on_hover(self.index, True)

    def _on_leave(self, event):
        """При уходе мыши."""
        self.image_label.config(background="")
        self.on_hover(self.index, False)
        self._hide_tooltip()

    def _on_click(self, event):
        """При клике на изображение."""
        self.selected.set(not self.selected.get())
        self._on_checkbox_click()

    def set_selected(self, selected):
        """Установить состояние выбора."""
        self.selected.set(selected)

    def is_selected(self):
        """Проверить, выбрана ли плитка."""
        return self.selected.get()

    def update_number(self, new_index):
        """Обновить номер фото."""
        self.index = new_index
        self.number_label.config(text=f"#{new_index + 1}")


class ThumbnailGrid(ttk.Frame):
    """Сетка с динамическим перестроением миниатюр."""

    def __init__(self, parent, image_processor, config, on_selection_change):
        super().__init__(parent)
        self.image_processor = image_processor
        self.config = config
        self.on_selection_change = on_selection_change
        self.tiles = []
        self.selected_indices = set()
        self.columns = 1
        self.tile_width = config.thumbnail_size[0] + 20

        self.create_grid_with_scroll()
        self.bind("<Configure>", self._on_parent_resize)

    def create_grid_with_scroll(self):
        """Создать прокручиваемую сетку."""
        self.canvas = tk.Canvas(self, highlightthickness=0, bg="white")
        self.scrollbar = ttk.Scrollbar(
            self, orient="vertical", command=self.canvas.yview
        )
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )

        self.canvas_window = self.canvas.create_window(
            (0, 0),
            window=self.scrollable_frame,
            anchor="nw",
            width=self.canvas.winfo_width(),
        )

        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_canvas_configure(self, event):
        """При изменении размера canvas обновляем ширину окна."""
        self.canvas.itemconfig(self.canvas_window, width=event.width)
        self._rebuild_grid()

    def _on_parent_resize(self, event):
        """При изменении размера родительского фрейма."""
        self._rebuild_grid()

    def _on_mousewheel(self, event):
        """Обработка колесика мыши."""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _calculate_columns(self):
        """Вычислить количество колонок на основе доступной ширины."""
        canvas_width = self.canvas.winfo_width()
        if canvas_width <= 1:
            return 1

        cols = max(1, canvas_width // (self.tile_width + 10))
        return cols

    def _rebuild_grid(self):
        """Перестроить сетку миниатюр."""
        count = self.image_processor.get_image_count()

        if count == 0:
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
            self.tiles.clear()
            return

        new_columns = self._calculate_columns()

        if new_columns != self.columns or len(self.tiles) != count:
            self.columns = new_columns
            self._rebuild_all_tiles()

    def _rebuild_all_tiles(self):
        """Полностью перестроить все плитки."""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.tiles.clear()

        count = self.image_processor.get_image_count()

        for i in range(count):
            thumbnail = self.image_processor.get_thumbnail(i)
            if thumbnail:
                tile = ThumbnailTile(
                    self.scrollable_frame,
                    i,
                    self.image_processor.get_image_paths()[i],
                    thumbnail,
                    self._on_tile_select,
                    self._on_tile_hover,
                    self._on_tile_rotate,
                    self.config,
                    self.image_processor,
                )
                self.tiles.append(tile)

                if i in self.selected_indices:
                    tile.set_selected(True)

        self._layout_tiles()

    def _layout_tiles(self):
        """Разместить плитки в сетке."""
        if not self.tiles:
            return

        for tile in self.tiles:
            tile.grid_forget()

        for i, tile in enumerate(self.tiles):
            row = i // self.columns
            col = i % self.columns
            tile.grid(row=row, column=col, padx=5, pady=5, sticky="nsew")

        for col in range(self.columns):
            self.scrollable_frame.columnconfigure(col, weight=1)

        self.scrollable_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_tile_select(self, index, selected):
        """Обработчик выбора плитки."""
        if selected:
            self.selected_indices.add(index)
        else:
            self.selected_indices.discard(index)

        self.on_selection_change(len(self.selected_indices))

    def _on_tile_hover(self, index, is_hover):
        """Обработчик наведения на плитку."""
        pass

    def _on_tile_rotate(self, index):
        """Обработчик поворота плитки."""
        if self.image_processor.rotate_image(index):
            self._rebuild_all_tiles()

    def add_images(self, image_paths):
        """Добавить новые изображения."""
        loaded = self.image_processor.load_images(image_paths)
        if loaded:
            self.columns = self._calculate_columns()
            self._rebuild_all_tiles()
        return loaded

    def remove_selected(self):
        """Удалить выбранные изображения."""
        if self.selected_indices:
            self.image_processor.remove_images(list(self.selected_indices))
            self.selected_indices.clear()
            self._rebuild_all_tiles()
            self.on_selection_change(0)

    def clear_all(self):
        """Очистить все изображения."""
        self.image_processor.clear_all()
        self.selected_indices.clear()
        self._rebuild_all_tiles()
        self.on_selection_change(0)

    def select_all(self):
        """Выделить все изображения."""
        self.selected_indices.clear()
        for i, tile in enumerate(self.tiles):
            tile.set_selected(True)
            self.selected_indices.add(i)
        self.on_selection_change(len(self.selected_indices))

    def get_selected_count(self):
        """Получить количество выбранных изображений."""
        return len(self.selected_indices)

    def get_selected_indices(self):
        """Получить индексы выбранных изображений."""
        return self.selected_indices


class PhotoTableApp:
    """Главное окно приложения."""

    def __init__(self, root):
        self.root = root
        self.config = AppConfig()
        self.image_processor = ImageProcessor(self.config)
        self.word_generator = WordGenerator(self.config, self.image_processor)

        self.setup_window()
        self.create_widgets()
        self.load_settings_to_ui()
        self.setup_bindings()

        self.generation_thread = None
        self.generation_cancelled = False

    def setup_window(self):
        """Настройка главного окна."""
        self.root.title(self.config.window_title)
        self.load_window_geometry()
        self.root.minsize(self.config.window_min_width, self.config.window_min_height)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        """Создание всех виджетов интерфейса."""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill="both", expand=True)

        self.create_top_panel(main_frame)
        self.create_grid_panel(main_frame)
        self.create_bottom_panel(main_frame)

    def create_top_panel(self, parent):
        """Создание верхней панели с кнопками."""
        top_frame = ttk.Frame(parent)
        top_frame.pack(fill="x", pady=(0, 10))

        self.btn_load = ttk.Button(
            top_frame, text="Загрузить (Ctrl+F)", command=self.load_images
        )
        self.btn_load.pack(side="left", padx=2)

        self.btn_clear = ttk.Button(
            top_frame, text="Очистить (Ctrl+Del)", command=self.clear_all
        )
        self.btn_clear.pack(side="left", padx=2)

        self.btn_delete = ttk.Button(
            top_frame, text="Удалить (0)", command=self.delete_selected
        )
        self.btn_delete.pack(side="left", padx=2)

        ttk.Separator(top_frame, orient="vertical").pack(side="left", fill="y", padx=10)

        self.total_label = ttk.Label(top_frame, text="Всего загружено: 0 фото")
        self.total_label.pack(side="left")

    def create_grid_panel(self, parent):
        """Создание панели с сеткой миниатюр."""
        grid_frame = ttk.LabelFrame(parent, text="Предпросмотр", padding="5")
        grid_frame.pack(fill="both", expand=True, pady=(0, 10))

        self.thumbnail_grid = ThumbnailGrid(
            grid_frame, self.image_processor, self.config, self.on_selection_change
        )
        self.thumbnail_grid.pack(fill="both", expand=True)

    def create_bottom_panel(self, parent):
        """Создание нижней панели с настройками и статусом."""
        bottom_frame = ttk.Frame(parent)
        bottom_frame.pack(fill="x")

        settings_frame = ttk.LabelFrame(
            bottom_frame, text="Настройки генерации", padding="5"
        )
        settings_frame.pack(side="left", fill="both", expand=True)

        # Ориентация
        ttk.Label(settings_frame, text="Ориентация:").grid(
            row=0, column=0, sticky="w", padx=5, pady=2
        )
        self.orientation = tk.StringVar(value="portrait")
        self.orientation.trace("w", self.on_orientation_change)

        portrait_text = "Книжная (2 колонки × 4 строки = 8 фото/стр)"
        landscape_text = "Альбомная (2 колонки × 2 строки = 4 фото/стр)"

        ttk.Radiobutton(
            settings_frame,
            text=portrait_text,
            variable=self.orientation,
            value="portrait",
        ).grid(row=0, column=1, columnspan=3, sticky="w", padx=5)
        ttk.Radiobutton(
            settings_frame,
            text=landscape_text,
            variable=self.orientation,
            value="landscape",
        ).grid(row=1, column=1, columnspan=3, sticky="w", padx=5)

        # Ширина таблицы
        ttk.Label(settings_frame, text="Ширина таблицы (см):").grid(
            row=2, column=0, sticky="w", padx=5, pady=(15, 5)
        )

        self.table_width = tk.DoubleVar(value=16.0)
        self.table_width_entry = ttk.Entry(
            settings_frame, textvariable=self.table_width, width=8
        )
        self.table_width_entry.grid(row=2, column=1, sticky="w", padx=5, pady=(15, 5))

        ttk.Label(settings_frame, text="см").grid(
            row=2, column=2, sticky="w", pady=(15, 5)
        )

        ttk.Label(
            settings_frame,
            text="(рекомендуется 16 для книжной, 25 для альбомной)",
            font=("Arial", 8, "italic"),
        ).grid(row=2, column=3, sticky="w", padx=5, pady=(15, 5))

        # Качество JPEG
        ttk.Label(settings_frame, text="Качество JPEG:").grid(
            row=3, column=0, sticky="w", padx=5, pady=(10, 5)
        )
        self.quality = tk.IntVar(value=85)

        quality_frame = ttk.Frame(settings_frame)
        quality_frame.grid(
            row=3, column=1, columnspan=3, sticky="ew", padx=5, pady=(10, 5)
        )

        quality_scale = ttk.Scale(
            quality_frame,
            from_=50,
            to=100,
            variable=self.quality,
            orient="horizontal",
            length=250,
        )
        quality_scale.pack(side="left", fill="x", expand=True)

        self.quality_label = ttk.Label(quality_frame, text="85%", width=5)
        self.quality_label.pack(side="right", padx=(10, 0))

        quality_scale.configure(
            command=lambda v: self.quality_label.config(text=f"{int(float(v))}%")
        )

        # Правая часть
        action_frame = ttk.Frame(bottom_frame)
        action_frame.pack(side="right", padx=(10, 0))

        self.progress = ttk.Progressbar(action_frame, mode="determinate", length=200)
        self.progress.pack(pady=(0, 5))

        self.btn_generate = ttk.Button(
            action_frame,
            text="Создать документ Word",
            command=self.generate_document,
            width=22,
        )
        self.btn_generate.pack()

        self.btn_cancel = ttk.Button(
            action_frame,
            text="Отмена",
            command=self.cancel_generation,
            width=22,
            state="disabled",
        )
        self.btn_cancel.pack(pady=(5, 0))

        # Статус
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill="x", pady=(5, 0))

        self.status_var = tk.StringVar(value=self.config.status_messages["ready"])
        self.status_label = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            foreground=self.config.colors["status_ready"],
        )
        self.status_label.pack(side="left")

    def load_settings_to_ui(self):
        """Загрузить сохраненные настройки в интерфейс."""
        self.orientation.set(self.config.get_orientation())
        self.table_width.set(self.config.get_table_width(self.orientation.get()))
        self.quality.set(self.config.get_jpeg_quality())
        self.quality_label.config(text=f"{self.quality.get()}%")

    def on_orientation_change(self, *args):
        """Обработчик изменения ориентации."""
        self.config.set_orientation(self.orientation.get())
        self.table_width.set(self.config.get_table_width(self.orientation.get()))

    def setup_bindings(self):
        """Настройка горячих клавиш."""
        self.root.bind("<Control-KeyPress>", self._on_ctrl_key)
        self.root.bind("<Control-Delete>", lambda e: self.clear_all())
        self.root.bind("<Delete>", lambda e: self.delete_selected())

    def _on_ctrl_key(self, event):
        """Обработка Ctrl+Key по кодам клавиш."""
        if event.keycode == 70:  # F
            self.load_images()
            return "break"
        if event.keycode == 65:  # A
            self.select_all()
            return "break"

    def load_images(self):
        """Загрузить изображения."""
        filetypes = [
            ("Изображения", "*.jpg *.jpeg *.png *.bmp *.gif *.tiff *.webp *.avif"),
            ("Все файлы", "*.*"),
        ]

        files = filedialog.askopenfilenames(
            title="Выберите фотографии",
            filetypes=filetypes,
            initialdir=self.config.get_last_path(),
        )

        if files:
            self.config.set_last_path(str(Path(files[0]).parent))
            loaded = self.thumbnail_grid.add_images(files)

            if loaded:
                self.update_total_count()
                self.set_status(self.config.status_messages["ready"], "ready")

    def delete_selected(self):
        """Удалить выбранные изображения."""
        count = self.thumbnail_grid.get_selected_count()
        if count > 0:
            if messagebox.askyesno("Подтверждение", f"Удалить {count} выбранных фото?"):
                self.thumbnail_grid.remove_selected()
                self.update_total_count()

    def clear_all(self):
        """Очистить все изображения."""
        if self.image_processor.get_image_count() > 0:
            if messagebox.askyesno("Подтверждение", "Очистить все загруженные фото?"):
                self.thumbnail_grid.clear_all()
                self.update_total_count()

    def select_all(self):
        """Выделить все изображения."""
        self.thumbnail_grid.select_all()

    def on_selection_change(self, selected_count):
        """Обработчик изменения выделения."""
        self.btn_delete.config(text=f"Удалить ({selected_count})")

    def update_total_count(self):
        """Обновить счетчик общего количества фото."""
        count = self.image_processor.get_image_count()
        self.total_label.config(text=f"Всего загружено: {count} фото")

    def generate_document(self):
        """Запустить генерацию документа."""
        if self.image_processor.get_image_count() == 0:
            messagebox.showwarning("Предупреждение", "Сначала загрузите фотографии!")
            return

        try:
            table_width = float(self.table_width.get())
            if table_width <= 0 or table_width > 50:
                messagebox.showerror(
                    "Ошибка", "Ширина таблицы должна быть от 1 до 50 см"
                )
                return
        except ValueError:
            messagebox.showerror(
                "Ошибка", "Введите корректное число для ширины таблицы"
            )
            return

        # Сохраняем настройки
        self.config.set_table_width(
            self.orientation.get(), float(self.table_width.get())
        )
        self.config.set_jpeg_quality(self.quality.get())

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"фототаблица_{timestamp}.docx"

        filepath = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
            initialdir=os.path.expanduser("~"),
            initialfile=default_filename,
        )

        if not filepath:
            return

        self.btn_generate.config(state="disabled")
        self.btn_cancel.config(state="normal")
        self.generation_cancelled = False

        self.current_filepath = filepath
        self.current_table_width = table_width

        self.generation_thread = threading.Thread(target=self._generate_thread)
        self.generation_thread.daemon = True
        self.generation_thread.start()

    def _generate_thread(self):
        """Поток генерации."""
        try:
            total_photos = self.image_processor.get_image_count()
            self.root.after(
                0,
                self.set_status,
                self.config.status_messages["processing"].format(
                    current=0, total=total_photos
                ),
                "processing",
            )
            self.root.after(0, self.progress.configure, {"maximum": total_photos})

            def progress_callback(current, total):
                self.root.after(0, self.progress.configure, {"value": current})
                self.root.after(
                    0,
                    self.set_status,
                    self.config.status_messages["processing"].format(
                        current=current, total=total
                    ),
                    "processing",
                )

            success = self.word_generator.save_to_file(
                self.current_filepath,
                orientation=self.orientation.get(),
                quality=self.quality.get(),
                table_width_cm=self.current_table_width,
                progress_callback=progress_callback,
            )

            if success and not self.generation_cancelled:
                self.root.after(
                    0,
                    self.set_status,
                    self.config.status_messages["success"],
                    "success",
                )
                self.root.after(0, lambda: self._ask_open_folder(self.current_filepath))

            elif self.generation_cancelled:
                self.root.after(0, self.set_status, "Генерация отменена", "ready")
                self.root.after(
                    0, lambda: messagebox.showinfo("Отмена", "Генерация отменена")
                )

        except Exception as e:
            logger.error(f"Ошибка генерации: {e}")
            self.root.after(0, self.set_status, f"Ошибка: {str(e)}", "error")
            self.root.after(0, lambda: messagebox.showerror("Ошибка", str(e)))

        finally:
            self.root.after(0, self.btn_generate.config, {"state": "normal"})
            self.root.after(0, self.btn_cancel.config, {"state": "disabled"})
            self.root.after(0, self.progress.configure, {"value": 0})

    def _ask_open_folder(self, filepath):
        """Спросить открыть ли папку с файлом."""
        if messagebox.askyesno("Готово", "Файл сохранен. Открыть папку с файлом?"):
            os.startfile(os.path.dirname(filepath))

    def cancel_generation(self):
        """Отменить генерацию."""
        self.generation_cancelled = True
        self.word_generator.cancel()

    def set_status(self, message, status_type):
        """Установить сообщение статуса."""
        self.status_var.set(message)

        color_map = {
            "ready": self.config.colors["status_ready"],
            "processing": self.config.colors["status_processing"],
            "success": self.config.colors["status_success"],
            "error": self.config.colors["status_error"],
        }

        self.status_label.config(foreground=color_map.get(status_type, "black"))

    def load_window_geometry(self):
        """Загрузить геометрию окна."""
        geometry = self.config.get_window_geometry()
        if geometry and all(k in geometry for k in ["width", "height", "x", "y"]):
            try:
                self.root.geometry(
                    f"{geometry['width']}x{geometry['height']}+"
                    f"{geometry['x']}+{geometry['y']}"
                )
                return
            except:
                pass

        self.root.state("zoomed")

    def save_window_geometry(self):
        """Сохранить геометрию окна."""
        if self.root.state() == "zoomed":
            return

        geometry = {
            "width": self.root.winfo_width(),
            "height": self.root.winfo_height(),
            "x": self.root.winfo_x(),
            "y": self.root.winfo_y(),
        }

        self.config.set_window_geometry(geometry)

    def on_closing(self):
        """При закрытии окна."""
        self.save_window_geometry()
        self.root.destroy()
