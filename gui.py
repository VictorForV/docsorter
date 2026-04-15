"""
GUI модуль DocSorter на customtkinter.
Иерархическая таблица с категориями, ручное управление.
"""

import asyncio
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

import customtkinter as ctk

from config import (
    load_config, save_config, is_config_valid,
    load_categories, save_categories,
    BASE_TEMPLATE_NAME,
    get_active_template, set_active_template, find_template,
    add_template, remove_template, rename_template, update_template_content,
)
from scanner import scan_folder
from analyzer import analyze_batch, analyze_file, check_suspicious
from grouper import group_documents, regroup_documents, generate_categories, with_page_count_suffix
from sorter import (
    build_folder_structure, build_numbering_structure,
    execute_sort, verify_sort, is_output_inside_source,
    filter_copyable,
)
from project import (
    save_project, load_project, get_default_project_path,
    build_project_state, normalize_document, find_by_hash, file_hash,
    PROJECT_FILENAME,
)
from slicer import (
    analyze_pdf_structure, verify_segments, slice_pdf, undo_slice,
    SLICE_SUBDIR,
)

import httpx


ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


OTHER_CATEGORY = "Прочее"


class DocSorterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("DocSorter — Сортировщик документов")
        self.geometry("1300x850")
        self.minsize(1000, 650)

        self.cfg = load_config()
        self.categories_library = load_categories()
        # Активный шаблон. Все вызовы grouper'а получают именно его.
        self.categories = get_active_template(self.categories_library)
        self.results = []              # Результаты анализа (плоский список)
        self.categories_order = []     # Порядок категорий в UI (список имён)
        self.source_dir = None
        self.source_count = 0          # Зафиксированное число файлов при сканировании
        self.source_pages = 0          # Зафиксированное число страниц при сканировании
        self.output_dir = None
        self._processing = False

        # Лог
        self._logs: list[str] = []

        # Проект
        self.project_path = None       # Path к текущему файлу проекта
        self._autosave_after_id = None  # ID отложенной задачи автосохранения
        self._suppress_autosave = False  # Флаг для подавления во время массовых обновлений

        # Глобальные биндинги для Ctrl+C/V/X/A (работают в любой раскладке)
        def _on_key_press(event):
            if event.state & 0x4:  # Ctrl
                if event.keycode == 86:   # V
                    event.widget.event_generate("<<Paste>>")
                    return "break"
                if event.keycode == 67:   # C
                    event.widget.event_generate("<<Copy>>")
                    return "break"
                if event.keycode == 88:   # X
                    event.widget.event_generate("<<Cut>>")
                    return "break"
                if event.keycode == 65:   # A
                    event.widget.event_generate("<<SelectAll>>")
                    return "break"

        self.bind_all("<Key>", _on_key_press)

        self._build_ui()
        self._build_menu()
        self._update_status()
        self._update_project_label()
        self._log("Приложение запущено")

    # ── UI ──────────────────────────────────────────────────────────

    def _build_ui(self):
        # Верхняя панель
        top = ctk.CTkFrame(self)
        top.pack(fill="x", padx=10, pady=(10, 5))

        self.panel_toggle_btn = ctk.CTkButton(
            top, text="☰", width=40, font=("", 16),
            command=self._toggle_left_panel,
        )
        self.panel_toggle_btn.pack(side="left", padx=(5, 5))

        self.status_label = ctk.CTkLabel(top, text="", font=("", 14))
        self.status_label.pack(side="left", padx=10)

        settings_btn = ctk.CTkButton(
            top, text="Настройки", width=100, command=self._open_settings,
        )
        settings_btn.pack(side="right", padx=5)

        cats_btn = ctk.CTkButton(
            top, text="Шаблоны категорий", width=150, command=self._open_categories,
        )
        cats_btn.pack(side="right", padx=5)

        # Выбор папки
        folder_frame = ctk.CTkFrame(self)
        folder_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(folder_frame, text="Папка:").pack(side="left", padx=(10, 5))

        self.folder_var = ctk.StringVar()
        self.folder_entry = ctk.CTkEntry(
            folder_frame, textvariable=self.folder_var, width=600,
        )
        self.folder_entry.pack(side="left", fill="x", expand=True, padx=5)

        browse_btn = ctk.CTkButton(
            folder_frame, text="Обзор...", width=80, command=self._browse_folder,
        )
        browse_btn.pack(side="left", padx=5)

        self.analyze_btn = ctk.CTkButton(
            folder_frame, text="Начать анализ", width=150,
            fg_color="green", hover_color="darkgreen",
            command=self._start_analysis,
        )
        self.analyze_btn.pack(side="right", padx=10)

        self.add_files_btn = ctk.CTkButton(
            folder_frame, text="+ Добавить файлы", width=150,
            command=self._add_files_to_project,
        )
        self.add_files_btn.pack(side="right", padx=5)

        # Прогресс
        self.progress_frame = ctk.CTkFrame(self)
        self.progress_frame.pack(fill="x", padx=10, pady=5)

        self.progress_label = ctk.CTkLabel(self.progress_frame, text="")
        self.progress_label.pack(side="left", padx=10)

        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=10)
        self.progress_bar.set(0)

        # Основная область: слева панель управления категориями, справа таблица
        main = ctk.CTkFrame(self)
        main.pack(fill="both", expand=True, padx=10, pady=5)

        # Левая панель управления (сворачиваемая)
        self.left_panel = ctk.CTkFrame(main, width=180)
        self.left_panel.pack(side="left", fill="y", padx=(0, 5))
        self.left_panel.pack_propagate(False)
        self._left_panel_visible = True

        ctk.CTkLabel(
            self.left_panel, text="Категории", font=("", 13, "bold"),
        ).pack(pady=(10, 5))

        ctk.CTkButton(
            self.left_panel, text="+ Добавить", width=160,
            command=self._add_category,
        ).pack(pady=3, padx=10)

        ctk.CTkButton(
            self.left_panel, text="Переименовать", width=160,
            command=self._rename_selected,
        ).pack(pady=3, padx=10)

        ctk.CTkButton(
            self.left_panel, text="Удалить", width=160,
            fg_color="#c0392b", hover_color="#922b21",
            command=self._delete_category,
        ).pack(pady=3, padx=10)

        ctk.CTkLabel(self.left_panel, text="").pack(pady=5)  # разделитель

        ctk.CTkButton(
            self.left_panel, text="▲ Выше", width=160,
            command=lambda: self._move_item(-1),
        ).pack(pady=3, padx=10)

        ctk.CTkButton(
            self.left_panel, text="▼ Ниже", width=160,
            command=lambda: self._move_item(1),
        ).pack(pady=3, padx=10)

        ctk.CTkLabel(self.left_panel, text="").pack(pady=5)

        ctk.CTkLabel(
            self.left_panel, text="Документы", font=("", 13, "bold"),
        ).pack(pady=(5, 5))

        ctk.CTkButton(
            self.left_panel, text="→ В категорию...", width=160,
            command=self._move_docs_to_category,
        ).pack(pady=3, padx=10)

        ctk.CTkButton(
            self.left_panel, text="✂ Нарезать PDF", width=160,
            fg_color="#d35400", hover_color="#a04000",
            command=self._slice_selected,
        ).pack(pady=3, padx=10)

        ctk.CTkButton(
            self.left_panel, text="Отменить нарезку", width=160,
            command=self._undo_slice_selected,
        ).pack(pady=3, padx=10)

        ctk.CTkButton(
            self.left_panel, text="Развернуть всё", width=160,
            command=self._expand_all,
        ).pack(pady=(15, 3), padx=10)

        ctk.CTkButton(
            self.left_panel, text="Свернуть всё", width=160,
            command=self._collapse_all,
        ).pack(pady=3, padx=10)

        # Таблица
        self.table_frame = ctk.CTkFrame(main)
        self.table_frame.pack(side="left", fill="both", expand=True)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Custom.Treeview",
            background="#2b2b2b",
            foreground="white",
            fieldbackground="#2b2b2b",
            rowheight=28,
            font=("", self.cfg.get("table_font_size", 11)),
        )
        style.configure(
            "Custom.Treeview.Heading",
            background="#1f538d",
            foreground="white",
            font=("", self.cfg.get("table_font_size", 11), "bold"),
        )
        style.map("Custom.Treeview", background=[("selected", "#1f538d")])

        self._all_columns = ("date", "title", "number", "party_1", "party_2", "amount", "goods_summary", "file", "type", "comment")
        self._collapsible_columns = ("file", "type", "party_2", "amount", "goods_summary")
        self._extra_visible = False
        self._displaycolumns_collapsed = tuple(c for c in self._all_columns if c not in self._collapsible_columns)

        self.tree = ttk.Treeview(
            self.table_frame, columns=self._all_columns, show="tree headings",
            style="Custom.Treeview", selectmode="extended",
            displaycolumns=self._displaycolumns_collapsed,
        )

        # Колонка дерева (#0) — категории с +/- для строк; в заголовке —
        # переключатель показа скрытых колонок "Файл" / "Тип".
        self._category_header = "Категория"
        self.tree.heading(
            "#0", text=f"{self._category_header}  [+]",
            command=self._toggle_extra_columns,
        )
        self.tree.column("#0", width=40, minwidth=30)

        self._headings = {
            "date": ("Дата", 80),
            "title": ("Основание", 220),
            "number": ("№", 60),
            "party_1": ("Сторона 1", 120),
            "party_2": ("Сторона 2", 120),
            "amount": ("Сумма", 100),
            "goods_summary": ("Предмет", 140),
            "file": ("Файл", 180),
            "type": ("Тип", 100),
            "comment": ("Комментарий", 160),
        }
        for col, (heading, width) in self._headings.items():
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=width, minwidth=40)

        # Теги для визуального различия
        self.tree.tag_configure("category", background="#1a3a5c", font=("", 11, "bold"))
        self.tree.tag_configure("document", background="#2b2b2b")
        self.tree.tag_configure("suspicious", background="#5c3a1a")
        self.tree.tag_configure("sliced", background="#3a3a3a", foreground="#888888")
        self.tree.tag_configure("missing", background="#2b2b2b", foreground="#666666")

        scrollbar_y = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(self.table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        # Сначала скроллбары, потом дерево — чтобы не было пустых зазоров
        scrollbar_x.pack(side="bottom", fill="x")
        scrollbar_y.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-3>", self._on_right_click)

        # Нижняя панель: промпт + действия
        bottom = ctk.CTkFrame(self)
        bottom.pack(fill="x", padx=10, pady=(5, 10))

        ctk.CTkLabel(bottom, text="Промпт:").pack(side="left", padx=(10, 5))

        self.prompt_var = ctk.StringVar()
        self.prompt_entry = ctk.CTkEntry(
            bottom, textvariable=self.prompt_var, width=350,
            placeholder_text="Перегруппируй: выдели судебные документы...",
        )
        self.prompt_entry.pack(side="left", fill="x", expand=True, padx=5)

        self.regroup_btn = ctk.CTkButton(
            bottom, text="Применить", width=100, command=self._regroup,
        )
        self.regroup_btn.pack(side="left", padx=5)

        # Режим сортировки — рядом с кнопкой копирования
        ctk.CTkLabel(bottom, text="Режим:").pack(side="left", padx=(20, 5))
        self.sort_mode_var = ctk.StringVar(value=self.cfg.get("sort_mode", "folders"))
        mode_dropdown = ctk.CTkOptionMenu(
            bottom,
            values=["По папкам", "По нумерации"],
            command=self._on_mode_change,
            width=140,
        )
        mode_dropdown.set(
            "По папкам" if self.sort_mode_var.get() == "folders" else "По нумерации"
        )
        mode_dropdown.pack(side="left", padx=5)

        self.sort_btn = ctk.CTkButton(
            bottom, text="Сортировать файлы", width=160,
            fg_color="#28a745", hover_color="#218838",
            command=self._execute_sort,
        )
        self.sort_btn.pack(side="right", padx=10)

        export_btn = ctk.CTkButton(
            bottom, text="Экспорт CSV", width=100, command=self._export_csv,
        )
        export_btn.pack(side="right", padx=5)

        # Статусбар (кликабельный — открывает окно логов)
        self.statusbar = ctk.CTkLabel(
            self, text="Готов к работе  📋", anchor="w", font=("", 11),
            cursor="hand2",
        )
        self.statusbar.pack(fill="x", padx=10, pady=(0, 5))
        self.statusbar.bind("<Button-1>", lambda e: self._open_log_window())

    def _on_mode_change(self, choice: str):
        self.sort_mode_var.set("folders" if choice == "По папкам" else "numbering")

    # ── Левая панель ───────────────────────────────────────────────

    def _toggle_left_panel(self):
        """Показать/скрыть левую панель управления."""
        if self._left_panel_visible:
            self.left_panel.pack_forget()
            self._left_panel_visible = False
        else:
            self.left_panel.pack(side="left", fill="y", padx=(0, 5), before=self.table_frame)
            self._left_panel_visible = True

    # ── Статус ─────────────────────────────────────────────────────

    def _update_status(self):
        if is_config_valid(self.cfg):
            self.status_label.configure(
                text="API настроен", text_color="lightgreen",
            )
        else:
            self.status_label.configure(
                text="API не настроен — откройте Настройки",
                text_color="orange",
            )

    def _apply_table_font(self):
        """Применяет размер шрифта таблицы из конфига."""
        import tkinter.ttk as _ttk
        size = self.cfg.get("table_font_size", 11)
        row_h = max(22, size + 16)
        style = _ttk.Style()
        style.configure("Custom.Treeview", font=("", size), rowheight=row_h)
        style.configure("Custom.Treeview.Heading", font=("", size, "bold"))
        # Перерисовка
        if hasattr(self, "tree") and self.tree.winfo_exists():
            self._populate_tree()

    def _set_statusbar(self, text: str):
        prefix = getattr(self, "_status_prefix", None)
        if self.project_path:
            name = Path(self.project_path).name
            prefix = f"Проект: {name}"
        elif prefix is None:
            prefix = "Проект: не сохранён"
        self.statusbar.configure(text=f"{prefix}  |  {text}  📋")

    # ── Логирование ────────────────────────────────────────────────

    def _log(self, message: str):
        """Добавляет запись в лог-буфер."""
        from datetime import datetime
        ts = datetime.now().strftime("%H:%M:%S")
        entry = f"[{ts}] {message}"
        self._logs.append(entry)
        # Обновляем окно логов если открыто
        log_win = getattr(self, "_log_window", None)
        if log_win and log_win.winfo_exists():
            textbox = getattr(self, "_log_textbox", None)
            if textbox:
                textbox.configure(state="normal")
                textbox.insert("end", entry + "\n")
                textbox.see("end")
                textbox.configure(state="disabled")

    def _open_log_window(self):
        """Открывает окно с логами."""
        # Если окно уже открыто — просто поднимаем
        log_win = getattr(self, "_log_window", None)
        if log_win and log_win.winfo_exists():
            log_win.lift()
            log_win.focus_force()
            return

        win = ctk.CTkToplevel(self)
        win.title("Лог")
        win.geometry("700x450")
        win.transient(self)
        self._log_window = win

        textbox = ctk.CTkTextbox(win, font=("Courier", 11))
        textbox.pack(fill="both", expand=True, padx=10, pady=10)
        textbox.configure(state="disabled")
        self._log_textbox = textbox

        # Заполняем существующими логами
        textbox.configure(state="normal")
        textbox.insert("1.0", "\n".join(self._logs) + "\n")
        textbox.see("end")
        textbox.configure(state="disabled")

        btn_frame = ctk.CTkFrame(win, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))

        def _copy():
            win.clipboard_clear()
            win.clipboard_append("\n".join(self._logs))

        ctk.CTkButton(btn_frame, text="Копировать всё", width=140, command=_copy).pack(side="left", padx=5)

        def _clear():
            self._logs.clear()
            textbox.configure(state="normal")
            textbox.delete("1.0", "end")
            textbox.configure(state="disabled")

        ctk.CTkButton(btn_frame, text="Очистить", width=100, command=_clear).pack(side="left", padx=5)

    # ── Настройки ──────────────────────────────────────────────────

    def _open_settings(self):
        win = ctk.CTkToplevel(self)
        win.title("Настройки")
        win.geometry("550x400")
        win.transient(self)
        win.grab_set()

        fields = {}
        row = 0

        labels = [
            ("api_key", "API ключ OpenRouter:", True),
            ("vision_model", "Vision модель:", False),
            ("text_model", "Текстовая модель:", False),
            ("max_concurrent", "Параллельных запросов:", False),
            ("max_pages_per_pdf", "Макс. страниц PDF:", False),
        ]

        # Поля ввода
        for key, label_text, is_secret in labels:
            ctk.CTkLabel(win, text=label_text).grid(
                row=row, column=0, padx=10, pady=8, sticky="e",
            )
            entry = ctk.CTkEntry(win, width=300)
            if is_secret:
                entry.configure(show="*")
            val = self.cfg.get(key, "")
            entry.insert(0, str(val))
            entry.grid(row=row, column=1, padx=10, pady=8, sticky="w")
            fields[key] = entry
            row += 1

        # Шаблон имени файла
        ctk.CTkLabel(win, text="Шаблон имени файла:").grid(
            row=row, column=0, padx=10, pady=8, sticky="e",
        )
        name_tpl_entry = ctk.CTkEntry(win, width=300)
        name_tpl_entry.insert(0, self.cfg.get("name_template", "{type} №{number} от {date} {party}"))
        name_tpl_entry.grid(row=row, column=1, padx=10, pady=8, sticky="w")
        fields["name_template"] = name_tpl_entry
        row += 1

        # Подсказка
        ctk.CTkLabel(
            win,
            text="Поля: {type} {number} {date} {party} {party_1} {party_2} {title} {amount}",
            font=("", 10), text_color="gray",
        ).grid(row=row, column=0, columnspan=2, pady=(0, 5))
        row += 1

        # Размер шрифта таблицы
        ctk.CTkLabel(win, text="Размер шрифта таблицы:").grid(
            row=row, column=0, padx=10, pady=8, sticky="e",
        )
        font_entry = ctk.CTkEntry(win, width=60)
        font_entry.insert(0, str(self.cfg.get("table_font_size", 11)))
        font_entry.grid(row=row, column=1, padx=10, pady=8, sticky="w")
        fields["table_font_size"] = font_entry
        row += 1

        def _save():
            for key, entry in fields.items():
                val = entry.get().strip()
                if key in ("max_concurrent", "max_pages_per_pdf", "table_font_size"):
                    try:
                        val = int(val)
                    except ValueError:
                        val = self.cfg.get(key, 5 if key != "table_font_size" else 11)
                self.cfg[key] = val
            save_config(self.cfg)
            self._apply_table_font()
            self._update_status()
            win.destroy()

        ctk.CTkButton(win, text="Сохранить", command=_save).grid(
            row=row, column=0, columnspan=2, pady=15,
        )

    # ── Шаблоны категорий ──────────────────────────────────────────

    def _open_categories(self):
        win = ctk.CTkToplevel(self)
        win.title("Шаблоны категорий")
        win.geometry("900x600")
        win.transient(self)
        win.grab_set()

        # Текущий выделенный в списке шаблон (изначально — активный)
        self._tpl_selected_name = self.categories_library.get(
            "active", BASE_TEMPLATE_NAME,
        )

        main = ctk.CTkFrame(win, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=10, pady=(10, 5))

        # ── Левая панель: список шаблонов ──
        left = ctk.CTkFrame(main, width=280)
        left.pack(side="left", fill="y", padx=(0, 10))
        left.pack_propagate(False)

        ctk.CTkLabel(
            left, text="Шаблоны", font=("", 13, "bold"),
        ).pack(anchor="w", padx=10, pady=(10, 5))

        self._tpl_list_frame = ctk.CTkScrollableFrame(left, fg_color="transparent")
        self._tpl_list_frame.pack(fill="both", expand=True, padx=5, pady=5)

        ctk.CTkLabel(
            left, text="✓ — активный   🔒 — базовый",
            font=("", 10), text_color="#888888",
        ).pack(anchor="w", padx=10, pady=(0, 8))

        # ── Правая панель: превью выбранного ──
        right = ctk.CTkFrame(main)
        right.pack(side="left", fill="both", expand=True)

        self._tpl_title_label = ctk.CTkLabel(
            right, text="", font=("", 14, "bold"), anchor="w", justify="left",
        )
        self._tpl_title_label.pack(fill="x", padx=10, pady=(10, 5))

        self._tpl_preview = ctk.CTkTextbox(right, font=("", 11), wrap="none")
        self._tpl_preview.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # ── Нижняя панель: кнопки действий ──
        actions = ctk.CTkFrame(win)
        actions.pack(fill="x", padx=10, pady=(0, 10))

        # Создание (всегда активно)
        row1 = ctk.CTkFrame(actions, fg_color="transparent")
        row1.pack(fill="x", pady=3, padx=5)
        ctk.CTkButton(
            row1, text="+ Создать через ИИ", width=170,
            command=lambda: self._tpl_create_via_ai(win),
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            row1, text="📋 Создать копию", width=150,
            command=lambda: self._tpl_copy(win),
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            row1, text="✓ Сделать активным", width=170,
            fg_color="#28a745", hover_color="#218838",
            command=lambda: self._tpl_set_active(win),
        ).pack(side="left", padx=2)

        # Действия над выделенным (часть disabled для базового / активного)
        row2 = ctk.CTkFrame(actions, fg_color="transparent")
        row2.pack(fill="x", pady=3, padx=5)
        self._tpl_btn_edit_ai = ctk.CTkButton(
            row2, text="✨ Изменить через ИИ", width=170,
            command=lambda: self._tpl_edit_via_ai(win),
        )
        self._tpl_btn_edit_ai.pack(side="left", padx=2)
        self._tpl_btn_edit_json = ctk.CTkButton(
            row2, text="✏ Редактировать JSON", width=170,
            command=lambda: self._tpl_edit_json(win),
        )
        self._tpl_btn_edit_json.pack(side="left", padx=2)
        self._tpl_btn_rename = ctk.CTkButton(
            row2, text="Переименовать", width=130,
            command=lambda: self._tpl_rename(win),
        )
        self._tpl_btn_rename.pack(side="left", padx=2)
        self._tpl_btn_delete = ctk.CTkButton(
            row2, text="🗑 Удалить", width=110,
            fg_color="#a73a3a", hover_color="#8a2a2a",
            command=lambda: self._tpl_delete(win),
        )
        self._tpl_btn_delete.pack(side="left", padx=2)

        ctk.CTkButton(win, text="Закрыть", command=win.destroy).pack(pady=(0, 10))

        self._tpl_refresh_list(win)

    # ── Вспомогательные методы окна шаблонов ─────────────────────

    def _tpl_refresh_list(self, win):
        for child in self._tpl_list_frame.winfo_children():
            child.destroy()

        active = self.categories_library.get("active")
        templates = self.categories_library.get("templates", [])

        # Если выбранного больше нет — выбираем активный
        names = [t.get("name") for t in templates]
        if self._tpl_selected_name not in names:
            self._tpl_selected_name = active if active in names else (
                names[0] if names else BASE_TEMPLATE_NAME
            )

        for t in templates:
            name = t.get("name", "?")
            is_base = t.get("is_base", False)
            is_active = name == active
            is_selected = name == self._tpl_selected_name

            prefix = "✓ " if is_active else "    "
            suffix = "  🔒" if is_base else ""
            text = f"{prefix}{name}{suffix}"

            fg = "#1f538d" if is_selected else "transparent"
            hover = "#264f7a" if is_selected else "#3a3a3a"

            btn = ctk.CTkButton(
                self._tpl_list_frame,
                text=text, anchor="w", height=32,
                fg_color=fg, hover_color=hover,
                command=lambda n=name: self._tpl_select(n, win),
            )
            btn.pack(fill="x", pady=1)

        self._tpl_refresh_preview()

    def _tpl_select(self, name, win):
        self._tpl_selected_name = name
        self._tpl_refresh_list(win)

    def _tpl_refresh_preview(self):
        name = self._tpl_selected_name
        t = find_template(self.categories_library, name)
        if not t:
            self._tpl_title_label.configure(text="—")
            self._tpl_preview.configure(state="normal")
            self._tpl_preview.delete("1.0", "end")
            self._tpl_preview.configure(state="disabled")
            return

        is_base = t.get("is_base", False)
        is_active = self.categories_library.get("active") == name

        title_parts = [name]
        marks = []
        if is_active:
            marks.append("активный")
        if is_base:
            marks.append("базовый, неизменяемый")
        if marks:
            title_parts.append("— " + ", ".join(marks))
        self._tpl_title_label.configure(text="  ".join(title_parts))

        # Рендерим категории как дерево
        lines = []
        for cat in t.get("categories", []):
            lines.append(f"📁 {cat.get('name', '')}")
            for sub in cat.get("subcategories", []):
                lines.append(f"      └ {sub}")
        text = "\n".join(lines) if lines else "(пусто)"

        self._tpl_preview.configure(state="normal")
        self._tpl_preview.delete("1.0", "end")
        self._tpl_preview.insert("1.0", text)
        self._tpl_preview.configure(state="disabled")

        # Доступность кнопок
        editable = "disabled" if is_base else "normal"
        self._tpl_btn_edit_ai.configure(state=editable)
        self._tpl_btn_edit_json.configure(state=editable)
        self._tpl_btn_rename.configure(state=editable)
        self._tpl_btn_delete.configure(
            state="disabled" if (is_base or is_active) else "normal",
        )

    def _tpl_persist(self):
        """Сохраняет библиотеку и обновляет self.categories (если активный изменился)."""
        save_categories(self.categories_library)
        self.categories = get_active_template(self.categories_library)

    # ── Действия над шаблонами ───────────────────────────────────

    def _tpl_set_active(self, win):
        name = self._tpl_selected_name
        if not find_template(self.categories_library, name):
            return
        if name == self.categories_library.get("active"):
            messagebox.showinfo("Активный шаблон", f'«{name}» уже активный.')
            return

        do_regroup = False
        if self.results:
            answer = messagebox.askyesnocancel(
                "Сменить активный шаблон",
                f'Сделать «{name}» активным?\n\n'
                f"В проекте {len(self.results)} документов.\n"
                "Перегруппировать их под новый шаблон?\n\n"
                "Да — сменить и перегруппировать (вызов API)\n"
                "Нет — только сменить, документы оставить как есть\n"
                "Отмена — ничего не делать",
            )
            if answer is None:
                return
            do_regroup = answer
        else:
            confirm = messagebox.askyesno(
                "Активный шаблон",
                f'Сделать «{name}» активным шаблоном?',
            )
            if not confirm:
                return

        set_active_template(self.categories_library, name)
        self._tpl_persist()
        self._tpl_refresh_list(win)

        if do_regroup:
            win.destroy()
            self._regroup_with_active_template()

    def _tpl_copy(self, win):
        src = self._tpl_selected_name
        src_t = find_template(self.categories_library, src)
        if not src_t:
            return

        new_name = self._ask_string(
            "Создать копию",
            f"Имя нового шаблона (на основе «{src}»):",
            initial=f"{src} (копия)",
        )
        if not new_name or not new_name.strip():
            return

        new_template = {
            "name": new_name.strip(),
            "categories": json.loads(json.dumps(src_t.get("categories", []))),
        }
        ok, err = add_template(self.categories_library, new_template)
        if not ok:
            messagebox.showerror("Ошибка", err)
            return
        self._tpl_persist()
        self._tpl_selected_name = new_name.strip()
        self._tpl_refresh_list(win)

    def _tpl_create_via_ai(self, win):
        if not is_config_valid(self.cfg):
            messagebox.showwarning("Внимание", "Настройте API ключ")
            return

        prompt = self._ask_string(
            "Создать шаблон через ИИ",
            "Опишите задачу (например: «арбитражные дела по налоговым спорам»):",
        )
        if not prompt or not prompt.strip():
            return
        prompt_text = prompt.strip()

        self._set_statusbar("Генерация шаблона через ИИ...")

        def _run():
            loop = asyncio.new_event_loop()
            try:
                new_cats = loop.run_until_complete(
                    generate_categories(
                        prompt_text, self.cfg["api_key"], self.cfg["text_model"],
                    )
                )
                self.after(0, lambda: self._on_tpl_generated(new_cats, win))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
                self.after(0, lambda: self._set_statusbar("Ошибка генерации"))
            finally:
                loop.close()

        threading.Thread(target=_run, daemon=True).start()

    def _on_tpl_generated(self, new_cats: dict, win):
        base_name = (new_cats.get("name") or "Новый шаблон").strip() or "Новый шаблон"
        name = base_name
        n = 2
        while find_template(self.categories_library, name):
            name = f"{base_name} ({n})"
            n += 1
        new_template = {
            "name": name,
            "categories": new_cats.get("categories", []),
        }
        ok, err = add_template(self.categories_library, new_template)
        if not ok:
            messagebox.showerror("Ошибка", err)
            return
        self._tpl_persist()
        self._tpl_selected_name = name
        self._tpl_refresh_list(win)
        self._set_statusbar(f"Шаблон создан: {name}")

    def _tpl_edit_via_ai(self, win):
        name = self._tpl_selected_name
        t = find_template(self.categories_library, name)
        if not t or t.get("is_base"):
            return
        if not is_config_valid(self.cfg):
            messagebox.showwarning("Внимание", "Настройте API ключ")
            return

        prompt = self._ask_string(
            "Изменить через ИИ",
            f"Что изменить в шаблоне «{name}»?",
        )
        if not prompt or not prompt.strip():
            return
        instruction = prompt.strip()

        current_json = json.dumps(
            {"name": t.get("name", ""), "categories": t.get("categories", [])},
            ensure_ascii=False, indent=2,
        )
        full_prompt = (
            f"Текущий шаблон категорий:\n{current_json}\n\n"
            f"Изменения, которые нужно внести: {instruction}\n\n"
            "Верни обновлённый шаблон в том же JSON-формате (поля name и categories). "
            "Сохрани категорию 'Прочее' в конце с пустым списком подкатегорий."
        )

        self._set_statusbar(f"Изменение шаблона «{name}»...")

        def _run():
            loop = asyncio.new_event_loop()
            try:
                new_cats = loop.run_until_complete(
                    generate_categories(
                        full_prompt, self.cfg["api_key"], self.cfg["text_model"],
                    )
                )
                self.after(0, lambda: self._on_tpl_edited(name, new_cats, win))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
                self.after(0, lambda: self._set_statusbar("Ошибка изменения"))
            finally:
                loop.close()

        threading.Thread(target=_run, daemon=True).start()

    def _on_tpl_edited(self, name: str, new_cats: dict, win):
        ok, err = update_template_content(self.categories_library, name, new_cats)
        if not ok:
            messagebox.showerror("Ошибка", err)
            return
        self._tpl_persist()
        self._tpl_refresh_list(win)
        self._set_statusbar(f"Шаблон обновлён: {name}")

    def _tpl_edit_json(self, win):
        name = self._tpl_selected_name
        t = find_template(self.categories_library, name)
        if not t or t.get("is_base"):
            return

        edit_win = ctk.CTkToplevel(win)
        edit_win.title(f"Редактировать JSON: {name}")
        edit_win.geometry("700x500")
        edit_win.transient(win)
        edit_win.grab_set()

        ctk.CTkLabel(
            edit_win,
            text="Редактируем категории шаблона. Поле 'name' изменяется отдельно (Переименовать).",
            font=("", 11), text_color="#888888",
        ).pack(anchor="w", padx=10, pady=(10, 5))

        text_widget = ctk.CTkTextbox(edit_win, font=("Courier", 11))
        text_widget.pack(fill="both", expand=True, padx=10, pady=5)
        text_widget.insert(
            "1.0",
            json.dumps(
                {"name": t.get("name", ""), "categories": t.get("categories", [])},
                ensure_ascii=False, indent=2,
            ),
        )

        def save():
            try:
                new_data = json.loads(text_widget.get("1.0", "end").strip())
            except json.JSONDecodeError as e:
                messagebox.showerror("Ошибка", f"Некорректный JSON: {e}")
                return
            ok, err = update_template_content(self.categories_library, name, new_data)
            if not ok:
                messagebox.showerror("Ошибка", err)
                return
            self._tpl_persist()
            self._tpl_refresh_list(win)
            edit_win.destroy()

        btns = ctk.CTkFrame(edit_win, fg_color="transparent")
        btns.pack(fill="x", padx=10, pady=10)
        ctk.CTkButton(btns, text="Сохранить", command=save).pack(side="right", padx=5)
        ctk.CTkButton(btns, text="Отмена", command=edit_win.destroy).pack(side="right", padx=5)

    def _tpl_rename(self, win):
        old = self._tpl_selected_name
        t = find_template(self.categories_library, old)
        if not t or t.get("is_base"):
            return
        new = self._ask_string("Переименовать шаблон", "Новое имя:", initial=old)
        if not new:
            return
        ok, err = rename_template(self.categories_library, old, new.strip())
        if not ok:
            messagebox.showerror("Ошибка", err)
            return
        self._tpl_persist()
        self._tpl_selected_name = new.strip()
        self._tpl_refresh_list(win)

    def _tpl_delete(self, win):
        name = self._tpl_selected_name
        t = find_template(self.categories_library, name)
        if not t or t.get("is_base"):
            return
        if self.categories_library.get("active") == name:
            messagebox.showwarning(
                "Внимание",
                "Нельзя удалить активный шаблон. Сначала переключитесь на другой.",
            )
            return
        confirm = messagebox.askyesno("Удалить шаблон", f"Удалить «{name}»?")
        if not confirm:
            return
        ok, err = remove_template(self.categories_library, name)
        if not ok:
            messagebox.showerror("Ошибка", err)
            return
        self._tpl_persist()
        self._tpl_selected_name = self.categories_library.get("active", BASE_TEMPLATE_NAME)
        self._tpl_refresh_list(win)

    def _regroup_with_active_template(self):
        """Перегруппировать всё под текущим активным шаблоном (без user-prompt)."""
        if not self.results:
            return
        if not is_config_valid(self.cfg):
            messagebox.showwarning("Внимание", "Настройте API ключ")
            return

        self._set_statusbar(f"Перегруппировка под «{self.categories.get('name', '')}»...")
        self.regroup_btn.configure(state="disabled")

        instruction = (
            "Перераспредели документы по новому набору категорий, сохраняя смысл "
            "групп связанных документов (договор + допсоглашения + первичка)."
        )

        def _run():
            loop = asyncio.new_event_loop()
            try:
                results = loop.run_until_complete(
                    regroup_documents(
                        self.results, self.categories, instruction,
                        self.cfg["api_key"], self.cfg["text_model"],
                    )
                )
                self.after(0, lambda: self._on_regroup_complete(results))
            except Exception as e:
                self.after(0, lambda: self._on_regroup_error(str(e)))
            finally:
                loop.close()

        threading.Thread(target=_run, daemon=True).start()

    # ── Выбор папки ────────────────────────────────────────────────

    def _browse_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку с документами")
        if folder:
            self.folder_var.set(folder)

    # ── Анализ ─────────────────────────────────────────────────────

    def _find_existing_project(self, source_dir: Path) -> Path | None:
        """Ищет файл проекта рядом с source: внутри самой папки и в соседних
        sorted-папках. Возвращает первый найденный путь или None.
        """
        candidates = [source_dir / PROJECT_FILENAME]
        parent = source_dir.parent
        if parent and parent != source_dir:
            # sorted, sorted_2, sorted_3, ... — проверяем без перебора всех чисел
            try:
                for sibling in parent.iterdir():
                    if sibling.is_dir() and sibling.name.startswith("sorted"):
                        candidates.append(sibling / PROJECT_FILENAME)
            except OSError:
                pass
        for c in candidates:
            if c.exists() and c.is_file():
                return c
        return None

    def _start_analysis(self):
        if self._processing:
            return

        folder = self.folder_var.get().strip()
        if not folder:
            messagebox.showwarning("Внимание", "Выберите папку")
            return

        if not is_config_valid(self.cfg):
            messagebox.showwarning("Внимание", "Настройте API ключ в Настройках")
            return

        new_source = Path(folder)
        if not new_source.exists():
            messagebox.showerror("Ошибка", "Папка не найдена")
            return

        # Авто-подхват существующего проекта в выбранной папке (или соседней sorted/).
        # Если уже открыт проект из ровно этого файла — не перезагружаем.
        existing_project = self._find_existing_project(new_source)
        if existing_project and (
            self.project_path is None or Path(self.project_path) != existing_project
        ):
            try:
                self._load_project_from_path(existing_project)
            except Exception as e:
                messagebox.showerror(
                    "Ошибка загрузки проекта",
                    f"Найден проект {existing_project.name}, но не загрузился:\n{e}",
                )
                return

        self.source_dir = new_source
        self._log(f"Исходная папка: {self.source_dir}")

        try:
            files = scan_folder(self.source_dir)
        except Exception as e:
            self._log(f"ОШИБКА сканирования: {e}")
            messagebox.showerror("Ошибка сканирования", str(e))
            return

        self._log(f"Сканирование: найдено {len(files)} файлов")

        if not files:
            messagebox.showinfo("Информация", "В папке нет поддерживаемых файлов")
            return

        # Если уже есть результаты (загруженный проект) — отфильтровываем уже обработанные
        is_incremental = bool(self.results)
        skipped = 0
        if is_incremental:
            new_files = []
            for f in files:
                h = file_hash(f["path"])
                if h and find_by_hash(self.results, h):
                    skipped += 1
                else:
                    new_files.append(f)
            files_to_analyze = new_files
        else:
            files_to_analyze = files

        if is_incremental and not files_to_analyze:
            self._log(f"Все {len(files)} файлов уже в проекте — новых нет")
            messagebox.showinfo(
                "Анализ",
                f"Проект уже содержит все {len(files)} файлов из папки. "
                "Новых файлов для анализа нет.",
            )
            return

        if is_incremental:
            self._set_statusbar(
                f"Найдено: {len(files)}, новых: {len(files_to_analyze)}, "
                f"уже в проекте: {skipped}",
            )
        else:
            # Фиксируем число файлов для финальной верификации (только при чистом запуске)
            self.source_count = len(files)
            self._set_statusbar(f"Найдено файлов: {len(files)}")
            self._clear_tree()

        self._processing = True
        self.analyze_btn.configure(state="disabled", text="Анализ...")
        self.progress_bar.set(0)
        self.progress_label.configure(text=f"0/{len(files_to_analyze)}")

        self._log(f"Начинаю анализ {len(files_to_analyze)} файлов")
        self._log(f"Vision: {self.cfg['vision_model']}, Text: {self.cfg['text_model']}")
        self._log(f"Параллельных запросов: {self.cfg.get('max_concurrent', 5)}")

        def _progress(current, total):
            self.after(0, lambda: self._update_progress(current, total))

        def _run():
            loop = asyncio.new_event_loop()
            try:
                self.after(0, lambda: self._log("Этап 1/2: анализ файлов через API..."))
                results = loop.run_until_complete(
                    analyze_batch(
                        files_to_analyze,
                        self.cfg["api_key"],
                        self.cfg["vision_model"],
                        self.cfg["text_model"],
                        self.cfg.get("max_pages_per_pdf", 3),
                        self.cfg.get("max_concurrent", 5),
                        progress_callback=_progress,
                        error_callback=lambda msg: self.after(0, lambda m=msg: self._log(m)),
                    )
                )
                self.after(0, lambda: self._log(f"Анализ завершён: {len(results)} результатов"))
                self.after(0, lambda: self._set_statusbar("Группировка документов..."))
                self.after(0, lambda: self._log("Этап 2/2: группировка по категориям..."))
                results = loop.run_until_complete(
                    group_documents(
                        results, self.categories,
                        self.cfg["api_key"], self.cfg["text_model"],
                        name_template=self.cfg.get("name_template", ""),
                    )
                )
                self.after(0, lambda: self._log("Группировка завершена"))
                self.after(
                    0,
                    lambda: self._on_analysis_complete(
                        results, is_incremental=is_incremental, skipped=skipped,
                    ),
                )
            except Exception as e:
                import traceback
                tb = traceback.format_exc()
                self.after(0, lambda: self._log(f"ОШИБКА: {e}\n{tb}"))
                self.after(0, lambda: self._on_analysis_error(str(e)))
            finally:
                loop.close()

        threading.Thread(target=_run, daemon=True).start()

    def _update_progress(self, current, total):
        self.progress_bar.set(current / total)
        self.progress_label.configure(text=f"{current}/{total}")

    def _on_analysis_complete(self, results, is_incremental: bool = False, skipped: int = 0):
        for doc in results:
            normalize_document(doc)

        pages = sum(d.get("_page_count", 0) for d in results)
        self._log(f"Анализ завершён. Файлов: {len(results)}, страниц: {pages}, пропущено: {skipped}")

        if is_incremental:
            self.results.extend(results)
            self.source_count += len(results)
            self.source_pages += pages
        else:
            self.results = results
            self.source_count = len(results)
            self.source_pages = pages

        self._processing = False
        self.analyze_btn.configure(state="normal", text="Начать анализ")
        self.progress_bar.set(1.0)

        # Инициализируем порядок категорий
        self._init_categories_order()
        self._populate_tree()

        if is_incremental:
            self._set_statusbar(
                f"Добавлено: {len(results)}, пропущено: {skipped}. "
                f"Всего документов: {len(self.results)}",
            )
        else:
            self._set_statusbar(f"Анализ завершён. Документов: {len(results)}")

        # Если проект ещё не привязан к файлу — биндим по умолчанию,
        # чтобы автосохранение начало писать в <source>/docsorter-project.json
        if self.project_path is None and self.source_dir:
            self.project_path = get_default_project_path(self.source_dir)
        self._schedule_autosave()

    def _on_analysis_error(self, error: str):
        self._log(f"ОШИБКА анализа: {error}")
        self._processing = False
        self.analyze_btn.configure(state="normal", text="Начать анализ")
        messagebox.showerror("Ошибка анализа", error)
        self._set_statusbar("Ошибка анализа")

    # ── Таблица (дерево) ───────────────────────────────────────────

    def _init_categories_order(self):
        """Инициализирует порядок категорий на основе результатов и шаблона."""
        # Сначала берём порядок из шаблона категорий
        order = [c["name"] for c in self.categories.get("categories", [])]

        # Добавляем категории из результатов, которых нет в шаблоне
        for doc in self.results:
            cat = doc.get("_category", OTHER_CATEGORY)
            if cat not in order:
                order.append(cat)

        # "Прочее" всегда в конце
        if OTHER_CATEGORY in order:
            order.remove(OTHER_CATEGORY)
        order.append(OTHER_CATEGORY)

        self.categories_order = order

    def _clear_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _docs_in_category(self, cat_name: str) -> list[tuple[int, dict]]:
        """Возвращает (index_in_results, doc) для всех документов категории."""
        docs = [(i, d) for i, d in enumerate(self.results)
                if d.get("_category", OTHER_CATEGORY) == cat_name]
        # Сортируем по sort_order
        docs.sort(key=lambda x: x[1].get("_sort_order", 99))
        return docs

    @staticmethod
    def _get_party_display(doc: dict, field: str) -> str:
        """Извлекает отображаемую строку из JSON party-поля."""
        import json as _json
        raw = doc.get(field, "")
        if not raw:
            if field == "party_1":
                return doc.get("counterparty", "")
            return ""
        try:
            data = _json.loads(raw)
            name = data.get("name", "")
            role = data.get("role", "")
            return f"{name} ({role})" if role else name
        except (_json.JSONDecodeError, TypeError):
            return raw

    def _populate_tree(self):
        """Заполняет дерево: категории → документы."""
        self._clear_tree()

        for cat_name in self.categories_order:
            docs = self._docs_in_category(cat_name)

            # Создаём узел категории, даже если пустая (чтобы можно было перемещать)
            cat_iid = f"cat:{cat_name}"
            cat_text = f"📁 {cat_name}  ({len(docs)})"
            self.tree.insert(
                "", "end", iid=cat_iid, text=cat_text,
                open=False, tags=("category",),
            )

            for idx, doc in docs:
                doc_iid = f"doc:{idx}"
                file_name = doc.get("_file_name", "?")

                # Гарантируем актуальный суффикс ' (N стр.)' в новом названии
                normalized = with_page_count_suffix(
                    doc.get("_new_name", ""), doc.get("_page_count", 0),
                )
                if normalized != doc.get("_new_name", ""):
                    doc["_new_name"] = normalized

                # Иконка и теги в зависимости от состояния
                tags = ["document"]
                icon = "📄"

                if doc.get("_slice_parts"):
                    icon = "📦"  # нарезан
                    tags = ["sliced"]
                elif doc.get("_suspicious"):
                    icon = "⚠️ 📄"
                    tags = ["suspicious"]

                # Проверка существования файла
                try:
                    if not Path(doc.get("_file_path", "")).exists():
                        tags = ["missing"]
                        icon = "❌"
                except Exception:
                    pass

                # В основание — ссылка на другой документ (reference)
                reason = doc.get("_suspicious_reason", "")
                if reason and doc.get("_suspicious"):
                    reference = f"[{reason}]"
                else:
                    reference = doc.get("reference", "")

                values = (
                    doc.get("date", ""),
                    reference,
                    doc.get("number", ""),
                    self._get_party_display(doc, "party_1"),
                    self._get_party_display(doc, "party_2"),
                    doc.get("amount", ""),
                    doc.get("goods_summary", ""),
                    file_name,
                    doc.get("doc_type", ""),
                    doc.get("_comment", ""),
                )
                self.tree.insert(
                    cat_iid, "end", iid=doc_iid,
                    text=f"  {icon}", values=values,
                    tags=tuple(tags),
                )

    def _expand_all(self):
        for iid in self.tree.get_children():
            self.tree.item(iid, open=True)

    def _collapse_all(self):
        for iid in self.tree.get_children():
            self.tree.item(iid, open=False)

    # ── Вспомогательные методы работы с выделением ─────────────────

    def _parse_iid(self, iid: str) -> tuple[str, str]:
        """Возвращает (type, value) где type=cat/doc."""
        if iid.startswith("cat:"):
            return ("cat", iid[4:])
        if iid.startswith("doc:"):
            return ("doc", iid[4:])
        return ("", iid)

    def _get_selected_category(self) -> str | None:
        """Возвращает имя выбранной категории (если выбрана категория)."""
        sel = self.tree.selection()
        if not sel:
            return None
        kind, val = self._parse_iid(sel[0])
        if kind == "cat":
            return val
        return None

    def _get_selected_docs(self) -> list[int]:
        """Возвращает индексы документов в self.results для выбранных строк."""
        sel = self.tree.selection()
        indices = []
        for iid in sel:
            kind, val = self._parse_iid(iid)
            if kind == "doc":
                try:
                    indices.append(int(val))
                except ValueError:
                    pass
        return indices

    # ── Скрытые колонки (Файл / Тип) ───────────────────────────────

    def _toggle_extra_columns(self):
        """Показать/скрыть колонки 'Файл' и 'Тип' по клику на заголовок 'Категория'."""
        self._extra_visible = not self._extra_visible
        if self._extra_visible:
            self.tree.configure(displaycolumns=self._all_columns)
            marker = "[−]"
        else:
            self.tree.configure(displaycolumns=self._displaycolumns_collapsed)
            marker = "[+]"
        self.tree.heading("#0", text=f"{self._category_header}  {marker}")

    def _open_file_in_system(self, file_path: str):
        """Открывает файл системным приложением по умолчанию
        (как если бы в Проводнике кликнули 2 раза)."""
        if not file_path:
            return
        p = Path(file_path)
        if not p.exists():
            messagebox.showwarning(
                "Файл не найден",
                f"Не удалось открыть файл:\n{file_path}",
            )
            return
        try:
            import os
            import sys
            import subprocess
            if sys.platform == "win32":
                os.startfile(str(p))  # noqa: E501 — только Windows
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(p)])
            else:
                subprocess.Popen(["xdg-open", str(p)])
        except Exception as e:
            messagebox.showerror("Ошибка открытия файла", str(e))

    # ── Редактирование по двойному клику ───────────────────────────

    def _on_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region not in ("cell", "tree"):
            return

        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return

        kind, val = self._parse_iid(row_id)

        # Двойной клик по категории = переименовать
        if kind == "cat":
            self._rename_category(val)
            return

        # Двойной клик по документу = редактировать ячейку
        if kind == "doc" and region == "cell":
            col = self.tree.identify_column(event.x)
            col_index = int(col.replace("#", "")) - 1
            if col_index < 0:
                return

            # identify_column возвращает индекс среди ОТОБРАЖАЕМЫХ колонок
            displaycols = self.tree.cget("displaycolumns")
            if not displaycols or displaycols == "#all" or (len(displaycols) == 1 and displaycols[0] == "#all"):
                used_cols = self._all_columns
            else:
                used_cols = tuple(displaycols)
            if col_index >= len(used_cols):
                return
            col_name = used_cols[col_index]

            # Двойной клик по "Файл" — открыть файл системным приложением
            if col_name == "file":
                try:
                    idx = int(val)
                except ValueError:
                    return
                if 0 <= idx < len(self.results):
                    self._open_file_in_system(
                        self.results[idx].get("_file_path", ""),
                    )
                return

            bbox = self.tree.bbox(row_id, col)
            if not bbox:
                return

            current_value = self.tree.set(row_id, col_name)

            entry = tk.Entry(self.tree, font=("", 11))
            entry.insert(0, current_value)
            entry.select_range(0, "end")
            entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
            entry.focus_set()

            def _save_edit(e=None):
                new_val = entry.get()
                entry.destroy()

                idx = int(val)
                field_map = {
                    "title": "reference",
                    "number": "number",
                    "date": "date",
                    "party_1": "party_1",
                    "party_2": "party_2",
                    "amount": "amount",
                    "goods_summary": "goods_summary",
                    "file": "_file_name",
                    "type": "doc_type",
                    "comment": "_comment",
                }
                data_key = field_map.get(col_name)
                if data_key and 0 <= idx < len(self.results):
                    # Для нового названия — всегда пересоздаём суффикс (N стр.)
                    if col_name == "new_name":
                        new_val = with_page_count_suffix(
                            new_val, self.results[idx].get("_page_count", 0),
                        )
                    self.results[idx][data_key] = new_val
                    self._schedule_autosave()
                self.tree.set(row_id, col_name, new_val)

            def _cancel(e=None):
                entry.destroy()

            entry.bind("<Return>", _save_edit)
            entry.bind("<Escape>", _cancel)
            entry.bind("<FocusOut>", _save_edit)

    # ── Контекстное меню ───────────────────────────────────────────

    def _on_right_click(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id and row_id not in self.tree.selection():
            self.tree.selection_set(row_id)

        sel = self.tree.selection()
        if not sel:
            return

        menu = tk.Menu(self, tearoff=0, bg="#2b2b2b", fg="white",
                       activebackground="#1f538d", activeforeground="white")

        # Определяем что выбрано
        kinds = {self._parse_iid(iid)[0] for iid in sel}

        if kinds == {"cat"} and len(sel) == 1:
            cat_name = self._parse_iid(sel[0])[1]
            menu.add_command(label="Переименовать", command=lambda: self._rename_category(cat_name))
            menu.add_command(label="Удалить", command=lambda: self._delete_category_by_name(cat_name))
            menu.add_separator()
            menu.add_command(label="▲ Выше", command=lambda: self._move_item(-1))
            menu.add_command(label="▼ Ниже", command=lambda: self._move_item(1))
            menu.add_separator()
            menu.add_command(label="+ Добавить категорию", command=self._add_category)

        elif kinds == {"doc"}:
            # Подменю "Переместить в категорию"
            move_menu = tk.Menu(menu, tearoff=0, bg="#2b2b2b", fg="white",
                                activebackground="#1f538d", activeforeground="white")
            for cat in self.categories_order:
                move_menu.add_command(
                    label=cat,
                    command=lambda c=cat: self._move_selected_to(c),
                )
            menu.add_cascade(label="→ Переместить в категорию", menu=move_menu)
            menu.add_separator()
            menu.add_command(label="▲ Выше (в группе)", command=lambda: self._move_item(-1))
            menu.add_command(label="▼ Ниже (в группе)", command=lambda: self._move_item(1))

        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    # ── Управление категориями ─────────────────────────────────────

    def _add_category(self):
        name = self._ask_string("Новая категория", "Введите название:")
        if not name:
            return
        name = name.strip()
        if not name:
            return
        if name in self.categories_order:
            messagebox.showwarning("Внимание", "Категория с таким именем уже существует")
            return

        # Вставляем перед "Прочее"
        if OTHER_CATEGORY in self.categories_order:
            idx = self.categories_order.index(OTHER_CATEGORY)
            self.categories_order.insert(idx, name)
        else:
            self.categories_order.append(name)

        self._populate_tree()
        # Выделяем новую категорию
        self.tree.selection_set(f"cat:{name}")
        self.tree.see(f"cat:{name}")

    def _rename_selected(self):
        sel = self.tree.selection()
        if not sel or len(sel) != 1:
            messagebox.showinfo("Информация", "Выберите одну категорию")
            return
        kind, val = self._parse_iid(sel[0])
        if kind != "cat":
            messagebox.showinfo("Информация", "Выберите категорию (не документ)")
            return
        self._rename_category(val)

    def _rename_category(self, old_name: str):
        new_name = self._ask_string("Переименовать категорию", "Новое название:", initial=old_name)
        if not new_name:
            return
        new_name = new_name.strip()
        if not new_name or new_name == old_name:
            return
        if new_name in self.categories_order:
            messagebox.showwarning("Внимание", "Категория с таким именем уже существует")
            return

        # Обновляем порядок
        idx = self.categories_order.index(old_name)
        self.categories_order[idx] = new_name

        # Обновляем документы
        for doc in self.results:
            if doc.get("_category") == old_name:
                doc["_category"] = new_name

        self._populate_tree()
        self.tree.selection_set(f"cat:{new_name}")

    def _delete_category(self):
        sel = self.tree.selection()
        if not sel or len(sel) != 1:
            messagebox.showinfo("Информация", "Выберите одну категорию")
            return
        kind, val = self._parse_iid(sel[0])
        if kind != "cat":
            messagebox.showinfo("Информация", "Выберите категорию")
            return
        self._delete_category_by_name(val)

    def _delete_category_by_name(self, cat_name: str):
        if cat_name == OTHER_CATEGORY:
            messagebox.showwarning("Внимание", f'Категорию "{OTHER_CATEGORY}" удалить нельзя')
            return

        docs = self._docs_in_category(cat_name)
        if docs:
            confirm = messagebox.askyesno(
                "Подтверждение",
                f'В категории "{cat_name}" {len(docs)} документ(ов).\n'
                f'Они будут перемещены в "{OTHER_CATEGORY}". Продолжить?',
            )
            if not confirm:
                return
            for idx, doc in docs:
                doc["_category"] = OTHER_CATEGORY

        if cat_name in self.categories_order:
            self.categories_order.remove(cat_name)

        # Гарантируем что "Прочее" есть
        if OTHER_CATEGORY not in self.categories_order:
            self.categories_order.append(OTHER_CATEGORY)

        self._populate_tree()

    def _move_item(self, direction: int):
        """direction: -1 = выше, +1 = ниже."""
        sel = self.tree.selection()
        if not sel:
            return

        # Определяем что выбрано
        first_kind = self._parse_iid(sel[0])[0]

        if first_kind == "cat" and len(sel) == 1:
            self._move_category(self._parse_iid(sel[0])[1], direction)
        elif first_kind == "doc":
            self._move_docs_within_category(direction)

    def _move_category(self, cat_name: str, direction: int):
        if cat_name == OTHER_CATEGORY:
            return  # "Прочее" всегда в конце
        order = self.categories_order
        if cat_name not in order:
            return
        idx = order.index(cat_name)
        new_idx = idx + direction
        # Не даём выйти за границы и не даём переставить ниже "Прочее"
        if new_idx < 0 or new_idx >= len(order):
            return
        if order[new_idx] == OTHER_CATEGORY and direction > 0:
            return
        order[idx], order[new_idx] = order[new_idx], order[idx]
        self._populate_tree()
        self.tree.selection_set(f"cat:{cat_name}")
        self.tree.see(f"cat:{cat_name}")

    def _move_docs_within_category(self, direction: int):
        """Меняет sort_order внутри категории."""
        indices = self._get_selected_docs()
        if not indices:
            return

        # Группируем по категории
        from collections import defaultdict
        by_cat = defaultdict(list)
        for idx in indices:
            cat = self.results[idx].get("_category", OTHER_CATEGORY)
            by_cat[cat].append(idx)

        for cat, idxs in by_cat.items():
            docs = self._docs_in_category(cat)  # [(idx, doc), ...]
            doc_idxs = [d[0] for d in docs]

            # Новые позиции
            positions = {idx: doc_idxs.index(idx) for idx in idxs if idx in doc_idxs}

            if direction < 0:
                for idx in sorted(positions, key=lambda x: positions[x]):
                    pos = doc_idxs.index(idx)
                    if pos > 0:
                        doc_idxs[pos], doc_idxs[pos - 1] = doc_idxs[pos - 1], doc_idxs[pos]
            else:
                for idx in sorted(positions, key=lambda x: positions[x], reverse=True):
                    pos = doc_idxs.index(idx)
                    if pos < len(doc_idxs) - 1:
                        doc_idxs[pos], doc_idxs[pos + 1] = doc_idxs[pos + 1], doc_idxs[pos]

            # Перезаписываем sort_order
            for new_pos, idx in enumerate(doc_idxs):
                self.results[idx]["_sort_order"] = new_pos + 1

        self._populate_tree()
        # Восстанавливаем выделение
        for idx in indices:
            self.tree.selection_add(f"doc:{idx}")

    # ── Перемещение документов между категориями ──────────────────

    def _move_docs_to_category(self):
        indices = self._get_selected_docs()
        if not indices:
            messagebox.showinfo("Информация", "Выберите один или несколько документов")
            return

        # Диалог выбора категории
        win = ctk.CTkToplevel(self)
        win.title("Переместить в категорию")
        win.geometry("350x400")
        win.transient(self)
        win.grab_set()

        ctk.CTkLabel(
            win, text=f"Документов: {len(indices)}", font=("", 13, "bold"),
        ).pack(pady=10)

        selected = {"cat": None}

        listbox = tk.Listbox(
            win, bg="#2b2b2b", fg="white", font=("", 11),
            selectbackground="#1f538d", height=15,
        )
        listbox.pack(fill="both", expand=True, padx=10, pady=5)

        for cat in self.categories_order:
            count = len(self._docs_in_category(cat))
            listbox.insert("end", f"{cat}  ({count})")

        def _apply():
            sel = listbox.curselection()
            if not sel:
                return
            cat = self.categories_order[sel[0]]
            self._move_selected_to(cat, indices)
            win.destroy()

        ctk.CTkButton(win, text="Переместить", command=_apply).pack(pady=10)

    def _move_selected_to(self, cat_name: str, indices: list[int] = None):
        if indices is None:
            indices = self._get_selected_docs()
        if not indices:
            return

        # Убедимся что категория существует
        if cat_name not in self.categories_order:
            self.categories_order.insert(
                max(0, len(self.categories_order) - 1), cat_name,
            )

        for idx in indices:
            self.results[idx]["_category"] = cat_name

        # Пересчитываем sort_order внутри целевой категории
        docs = self._docs_in_category(cat_name)
        for i, (idx, _) in enumerate(docs):
            self.results[idx]["_sort_order"] = i + 1

        self._populate_tree()
        for idx in indices:
            self.tree.selection_add(f"doc:{idx}")

    # ── Диалог ввода строки ────────────────────────────────────────

    def _ask_string(self, title: str, prompt: str, initial: str = "") -> str | None:
        win = ctk.CTkToplevel(self)
        win.title(title)
        win.geometry("400x150")
        win.transient(self)
        win.grab_set()

        ctk.CTkLabel(win, text=prompt).pack(pady=(15, 5))

        var = ctk.StringVar(value=initial)
        entry = ctk.CTkEntry(win, textvariable=var, width=350)
        entry.pack(pady=5)
        entry.focus_set()
        entry.select_range(0, "end")

        result = {"value": None}

        def _ok():
            result["value"] = var.get()
            win.destroy()

        def _cancel():
            win.destroy()

        entry.bind("<Return>", lambda e: _ok())
        entry.bind("<Escape>", lambda e: _cancel())

        btn_frame = ctk.CTkFrame(win, fg_color="transparent")
        btn_frame.pack(pady=10)
        ctk.CTkButton(btn_frame, text="OK", width=100, command=_ok).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Отмена", width=100, command=_cancel).pack(side="left", padx=5)

        self.wait_window(win)
        return result["value"]

    # ── Перегруппировка через LLM ──────────────────────────────────

    def _regroup(self):
        if not self.results:
            messagebox.showinfo("Информация", "Сначала проведите анализ")
            return

        prompt_text = self.prompt_var.get().strip()
        if not prompt_text:
            return

        self._set_statusbar("Перегруппировка...")
        self.regroup_btn.configure(state="disabled")

        def _run():
            loop = asyncio.new_event_loop()
            try:
                results = loop.run_until_complete(
                    regroup_documents(
                        self.results, self.categories, prompt_text,
                        self.cfg["api_key"], self.cfg["text_model"],
                    )
                )
                self.after(0, lambda: self._on_regroup_complete(results))
            except Exception as e:
                self.after(0, lambda: self._on_regroup_error(str(e)))
            finally:
                loop.close()

        threading.Thread(target=_run, daemon=True).start()

    def _on_regroup_complete(self, results):
        self.results = results
        self._init_categories_order()
        self._populate_tree()
        self.regroup_btn.configure(state="normal")
        self._set_statusbar("Перегруппировка завершена")

    def _on_regroup_error(self, error: str):
        self.regroup_btn.configure(state="normal")
        messagebox.showerror("Ошибка", error)
        self._set_statusbar("Ошибка перегруппировки")

    # ── Копирование ────────────────────────────────────────────────

    def _execute_sort(self):
        if not self.results:
            messagebox.showinfo("Информация", "Сначала проведите анализ")
            return

        # Дефолт: соседняя с source папка "sorted" (с числовым суффиксом, если занята)
        default_output = None
        if self.source_dir:
            base_parent = self.source_dir.parent
            candidate = base_parent / "sorted"
            n = 2
            while candidate.exists() and any(candidate.iterdir()):
                candidate = base_parent / f"sorted_{n}"
                n += 1
            default_output = candidate

        if default_output is not None:
            answer = messagebox.askyesnocancel(
                "Папка для сортировки",
                f"Скопировать файлы в:\n{default_output}\n\n"
                "Да — использовать эту папку\n"
                "Нет — выбрать другую\n"
                "Отмена — прервать",
            )
            if answer is None:
                return
            if answer:
                self.output_dir = default_output
            else:
                output = filedialog.askdirectory(
                    title="Выберите папку для отсортированных файлов",
                    initialdir=str(self.source_dir.parent),
                )
                if not output:
                    return
                self.output_dir = Path(output)
        else:
            output = filedialog.askdirectory(title="Выберите папку для отсортированных файлов")
            if not output:
                return
            self.output_dir = Path(output)

        # Проверка: output не должен быть внутри source
        if self.source_dir and is_output_inside_source(self.source_dir, self.output_dir):
            messagebox.showerror(
                "Ошибка",
                "Папка назначения находится внутри исходной папки.\n"
                "Это приведёт к зацикливанию. Выберите другую папку.",
            )
            return

        # Сортируем результаты согласно текущему порядку категорий
        sorted_results = self._get_sorted_results()

        # Строим пути
        mode = self.sort_mode_var.get()
        if mode == "folders":
            build_folder_structure(sorted_results, self.output_dir)
        else:
            build_numbering_structure(sorted_results, self.output_dir)

        self._set_statusbar("Копирование файлов...")
        result = execute_sort(sorted_results, self.output_dir)

        # Верификация с зафиксированным source_count/source_pages
        verification = verify_sort(
            self.source_count, self.output_dir,
            source_pages=self.source_pages, results=sorted_results,
        )

        msg = (
            f"Скопировано: {result['copied']} из {len(sorted_results)}\n"
            f"Ошибок: {len(result['errors'])}\n\n"
            f"Проверка:\n"
            f"  Файлов: {verification['source_count']} → {verification['dest_count']}"
            f"  {'Да ✓' if verification['match'] else 'НЕТ!'}\n"
        )
        if verification['source_pages'] > 0:
            msg += (
                f"  Страниц: {verification['source_pages']} → {verification['dest_pages']}"
                f"  {'Да ✓' if verification['pages_match'] else 'НЕТ!'}"
            )

        if result["errors"]:
            msg += "\n\nОшибки:\n" + "\n".join(result["errors"][:10])

        if verification["match"] and not result["errors"]:
            messagebox.showinfo("Готово", msg)
            self._set_statusbar("Сортировка завершена успешно")
        else:
            messagebox.showwarning("Внимание", msg)
            self._set_statusbar("Сортировка завершена с расхождениями")

    def _get_sorted_results(self) -> list[dict]:
        """Возвращает результаты в порядке категорий и sort_order."""
        ordered = []
        for cat_name in self.categories_order:
            docs = self._docs_in_category(cat_name)
            ordered.extend([d[1] for d in docs])

        # Добавим документы, категории которых почему-то нет в order
        seen_ids = {id(d) for d in ordered}
        for d in self.results:
            if id(d) not in seen_ids:
                ordered.append(d)

        return ordered

    # ── Экспорт ────────────────────────────────────────────────────

    def _export_csv(self):
        if not self.results:
            messagebox.showinfo("Информация", "Нет данных для экспорта")
            return

        path = filedialog.asksaveasfilename(
            title="Сохранить CSV",
            defaultextension=".csv",
            filetypes=[("CSV файлы", "*.csv")],
        )
        if not path:
            return

        import csv
        sorted_results = self._get_sorted_results()
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow([
                "№", "Исходный файл", "Тип", "Название", "Номер",
                "Дата", "Сторона 1", "Сторона 2", "Сумма", "Предмет",
                "Категория", "Группа", "Новое имя", "Комментарий", "Содержание",
            ])
            for i, doc in enumerate(sorted_results):
                writer.writerow([
                    i + 1,
                    doc.get("_file_name", ""),
                    doc.get("doc_type", ""),
                    doc.get("title", ""),
                    doc.get("number", ""),
                    doc.get("date", ""),
                    self._get_party_display(doc, "party_1"),
                    self._get_party_display(doc, "party_2"),
                    doc.get("amount", ""),
                    doc.get("goods_summary", ""),
                    doc.get("_category", ""),
                    doc.get("_group", ""),
                    doc.get("_new_name", ""),
                    doc.get("_comment", ""),
                    doc.get("summary", ""),
                ])

        self._set_statusbar(f"Экспорт: {path}")
        messagebox.showinfo("Экспорт", f"Данные сохранены:\n{path}")

    # ── Меню ───────────────────────────────────────────────────────

    def _build_menu(self):
        menubar = tk.Menu(self, bg="#2b2b2b", fg="white")

        file_menu = tk.Menu(menubar, tearoff=0, bg="#2b2b2b", fg="white",
                            activebackground="#1f538d", activeforeground="white")
        file_menu.add_command(label="Новый проект", command=self._on_new_project)
        file_menu.add_command(label="Открыть проект...", command=self._on_open_project)
        file_menu.add_separator()
        file_menu.add_command(label="Сохранить", accelerator="Ctrl+S", command=self._on_save_project)
        file_menu.add_command(label="Сохранить как...", command=self._on_save_as)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.destroy)
        menubar.add_cascade(label="Файл", menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=0, bg="#2b2b2b", fg="white",
                            activebackground="#1f538d", activeforeground="white")
        help_menu.add_command(label="О программе", command=self._show_about)
        menubar.add_cascade(label="Справка", menu=help_menu)

        self.configure(menu=menubar)
        self.bind("<Control-s>", lambda e: self._on_save_project())

    def _show_about(self):
        messagebox.showinfo(
            "О программе",
            "DocSorter — Сортировщик документов\n\n"
            "Автоматическая сортировка документов с помощью LLM (OpenRouter).\n"
            "Поддержка PDF, изображений, DOCX, XLSX.",
        )

    # ── Проект: новые / открыть / сохранить ────────────────────────

    def _update_project_label(self):
        """Обновляет индикатор проекта в статусбаре."""
        if self.project_path:
            name = Path(self.project_path).name
            prefix = f"Проект: {name}"
        else:
            prefix = "Проект: не сохранён"
        base = self.statusbar.cget("text")
        # Сохраняем текущую информацию, но добавляем префикс
        self._status_prefix = prefix
        self.statusbar.configure(text=f"{prefix}  |  {base.split('|', 1)[-1].strip() if '|' in base else base}")

    def _on_new_project(self):
        if self.results:
            confirm = messagebox.askyesno(
                "Новый проект",
                "Создать новый проект? Текущие данные будут сброшены "
                "(последнее автосохранение останется на диске).",
            )
            if not confirm:
                return

        self.results = []
        self.categories_order = []
        self.project_path = None
        self.source_dir = None
        self.source_count = 0
        self.source_pages = 0
        self.folder_var.set("")
        self._clear_tree()
        self.progress_bar.set(0)
        self.progress_label.configure(text="")
        self._set_statusbar("Новый проект")

    def _on_open_project(self):
        path = filedialog.askopenfilename(
            title="Открыть проект",
            filetypes=[("Проект DocSorter", "*.json"), ("Все файлы", "*.*")],
        )
        if not path:
            return
        self._load_project_from_path(Path(path))

    def _load_project_from_path(self, path: Path):
        try:
            data = load_project(path)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть проект:\n{e}")
            return

        self._suppress_autosave = True
        try:
            self.project_path = path
            self.source_dir = Path(data.get("source_dir", "")) if data.get("source_dir") else None
            self.folder_var.set(str(self.source_dir) if self.source_dir else "")
            self.output_dir = (
                Path(data.get("output_dir")) if data.get("output_dir") else None
            )
            self.sort_mode_var.set(data.get("sort_mode", "folders"))
            self.categories_order = list(data.get("categories_order", []))
            docs = [normalize_document(d) for d in data.get("documents", [])]
            self.results = docs
            self.source_count = len(docs)
            self.source_pages = sum(d.get("_page_count", 0) for d in docs)

            # Если categories_order пуст, пересчитываем
            if not self.categories_order:
                self._init_categories_order()

            self._populate_tree()
            self._set_statusbar(f"Проект загружен. Документов: {len(docs)}")
        finally:
            self._suppress_autosave = False

    def _on_save_project(self):
        if not self.results:
            messagebox.showinfo("Информация", "Нечего сохранять")
            return

        if self.project_path is None:
            # Если есть source_dir, предлагаем путь по умолчанию
            if self.source_dir:
                self.project_path = get_default_project_path(self.source_dir)
            else:
                self._on_save_as()
                return

        self._save_project_to_path(self.project_path)
        self._set_statusbar(f"Сохранено: {self.project_path.name}")

    def _on_save_as(self):
        if not self.results:
            messagebox.showinfo("Информация", "Нечего сохранять")
            return

        initial_dir = str(self.source_dir) if self.source_dir else ""
        path = filedialog.asksaveasfilename(
            title="Сохранить проект",
            defaultextension=".json",
            initialfile=PROJECT_FILENAME,
            initialdir=initial_dir,
            filetypes=[("Проект DocSorter", "*.json")],
        )
        if not path:
            return
        self.project_path = Path(path)
        self._save_project_to_path(self.project_path)
        self._set_statusbar(f"Сохранено: {self.project_path.name}")

    def _save_project_to_path(self, path: Path):
        try:
            state = build_project_state(
                source_dir=self.source_dir,
                output_dir=self.output_dir,
                sort_mode=self.sort_mode_var.get(),
                suspicious_page_threshold=self.cfg.get("suspicious_page_threshold", 5),
                categories_order=self.categories_order,
                documents=self.results,
            )
            save_project(state, path)
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e))

    def _schedule_autosave(self):
        """Запускает отложенное автосохранение с дебаунсом 2 сек."""
        if self._suppress_autosave:
            return
        if self.project_path is None:
            # Проект ещё не сохранён — не автосохраняемся
            return
        if self._autosave_after_id is not None:
            try:
                self.after_cancel(self._autosave_after_id)
            except Exception:
                pass
        self._autosave_after_id = self.after(2000, self._do_autosave)

    def _do_autosave(self):
        self._autosave_after_id = None
        if self.project_path is None or not self.results:
            return
        try:
            self._save_project_to_path(self.project_path)
            from datetime import datetime
            ts = datetime.now().strftime("%H:%M:%S")
            self._set_statusbar(f"Автосохранено в {ts}")
        except Exception as e:
            self._set_statusbar(f"Ошибка автосохранения: {e}")

    # ── Добавление файлов в проект ─────────────────────────────────

    def _add_files_to_project(self):
        if self._processing:
            return
        if not self.results:
            messagebox.showinfo(
                "Информация",
                "Сначала проведите анализ или откройте существующий проект",
            )
            return
        if not is_config_valid(self.cfg):
            messagebox.showwarning("Внимание", "Настройте API ключ в Настройках")
            return

        # Спрашиваем: папка или отдельные файлы
        choice = messagebox.askyesnocancel(
            "Добавить файлы",
            "Да — выбрать папку (рекурсивно).\n"
            "Нет — выбрать отдельные файлы.\n"
            "Отмена — ничего не делать.",
        )
        if choice is None:
            return

        files = []
        if choice:
            folder = filedialog.askdirectory(title="Выберите папку")
            if not folder:
                return
            try:
                files = scan_folder(Path(folder))
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))
                return
        else:
            paths = filedialog.askopenfilenames(title="Выберите файлы")
            if not paths:
                return
            from scanner import SUPPORTED_EXTENSIONS
            for p in paths:
                pp = Path(p)
                if pp.suffix.lower() not in SUPPORTED_EXTENSIONS:
                    continue
                files.append({
                    "path": pp,
                    "name": pp.name,
                    "ext": pp.suffix.lower(),
                    "size": pp.stat().st_size,
                    "rel_path": pp.name,
                })

        if not files:
            messagebox.showinfo("Информация", "Подходящих файлов не найдено")
            return

        # Дедуп по хэшу
        new_files = []
        skipped = 0
        for f in files:
            h = file_hash(f["path"])
            if find_by_hash(self.results, h):
                skipped += 1
            else:
                new_files.append(f)

        if not new_files:
            messagebox.showinfo(
                "Информация",
                f"Все {len(files)} файлов уже в проекте.",
            )
            return

        confirm = messagebox.askyesno(
            "Добавление файлов",
            f"Найдено: {len(files)}\n"
            f"Новых: {len(new_files)}\n"
            f"Уже в проекте: {skipped}\n\n"
            f"Проанализировать {len(new_files)} новых файлов?",
        )
        if not confirm:
            return

        self._processing = True
        self.add_files_btn.configure(state="disabled", text="Анализ...")
        self.progress_bar.set(0)
        self.progress_label.configure(text=f"0/{len(new_files)}")

        def _progress(current, total):
            self.after(0, lambda: self._update_progress(current, total))

        def _run():
            loop = asyncio.new_event_loop()
            try:
                results = loop.run_until_complete(
                    analyze_batch(
                        new_files,
                        self.cfg["api_key"],
                        self.cfg["vision_model"],
                        self.cfg["text_model"],
                        self.cfg.get("max_pages_per_pdf", 3),
                        self.cfg.get("max_concurrent", 5),
                        self.cfg.get("suspicious_page_threshold", 5),
                        progress_callback=_progress,
                    )
                )
                self.after(0, lambda: self._set_statusbar("Группировка новых файлов..."))
                results = loop.run_until_complete(
                    group_documents(
                        results, self.categories,
                        self.cfg["api_key"], self.cfg["text_model"],
                        name_template=self.cfg.get("name_template", ""),
                    )
                )
                self.after(0, lambda: self._on_add_files_complete(results, skipped))
            except Exception as e:
                self.after(0, lambda: self._on_add_files_error(str(e)))
            finally:
                loop.close()

        threading.Thread(target=_run, daemon=True).start()

    def _on_add_files_complete(self, new_results: list[dict], skipped: int):
        # Дополняем дефолтными полями
        for doc in new_results:
            normalize_document(doc)

        self.results.extend(new_results)
        self.source_count += len(new_results)
        self._init_categories_order()
        self._populate_tree()
        self._processing = False
        self.add_files_btn.configure(state="normal", text="+ Добавить файлы")
        self.progress_bar.set(1.0)
        self._set_statusbar(
            f"Добавлено: {len(new_results)}, пропущено: {skipped}"
        )
        self._schedule_autosave()

    def _on_add_files_error(self, error: str):
        self._processing = False
        self.add_files_btn.configure(state="normal", text="+ Добавить файлы")
        messagebox.showerror("Ошибка", error)

    # ── Нарезка PDF ────────────────────────────────────────────────

    def _slice_selected(self):
        if self._processing:
            return
        indices = self._get_selected_docs()
        if len(indices) != 1:
            messagebox.showinfo("Информация", "Выберите один PDF-документ для нарезки")
            return

        idx = indices[0]
        doc = self.results[idx]
        if doc.get("_ext") != ".pdf":
            messagebox.showinfo("Информация", "Нарезка поддерживается только для PDF")
            return
        if doc.get("_slice_parts"):
            messagebox.showinfo("Информация", "Этот документ уже нарезан")
            return
        # Повторная нарезка уже нарезанной части допустима — старая часть будет
        # заменена новыми (см. _execute_slicing / _on_slicing_complete).

        pdf_path = Path(doc["_file_path"])
        if not pdf_path.exists():
            messagebox.showerror("Ошибка", "Файл не найден")
            return

        page_count = doc.get("_page_count", 0)
        if page_count < 2:
            messagebox.showinfo("Информация", "PDF содержит менее 2 страниц — нарезка не требуется")
            return

        # Предупреждение для больших PDF
        if page_count > 50:
            confirm = messagebox.askyesno(
                "Большой PDF",
                f"PDF содержит {page_count} страниц. Нарезка может занять время и стоить денег на API. Продолжить?",
            )
            if not confirm:
                return

        self._processing = True
        self._set_statusbar(f"Анализ структуры PDF ({page_count} стр.)...")
        self.progress_bar.set(0)

        def _progress(current, total):
            self.after(0, lambda: self._update_progress(current, total))

        def _run():
            loop = asyncio.new_event_loop()
            try:
                segments = loop.run_until_complete(
                    analyze_pdf_structure(
                        pdf_path,
                        self.cfg["api_key"],
                        self.cfg["vision_model"],
                        self.cfg.get("slice_batch_size", 10),
                        progress_callback=_progress,
                    )
                )
                self.after(0, lambda: self._on_structure_ready(idx, segments, page_count))
            except Exception as e:
                self.after(0, lambda: self._on_slice_error(str(e)))
            finally:
                loop.close()

        threading.Thread(target=_run, daemon=True).start()

    def _on_structure_ready(self, idx: int, segments: list[dict], total_pages: int):
        self._processing = False
        self.progress_bar.set(1.0)

        ok, err = verify_segments(segments, total_pages)

        if not ok:
            # Открываем диалог ручной правки
            self._open_segments_editor(idx, segments, total_pages, err)
        else:
            # Подтверждение и нарезка
            summary = "\n".join(
                f"  {i+1}. стр. {s['page_from']}-{s['page_to']}: "
                f"{s.get('doc_type', '?')} — {s.get('title', '')}"
                for i, s in enumerate(segments)
            )
            confirm = messagebox.askyesno(
                "Нарезать PDF",
                f"Найдено {len(segments)} документов:\n\n{summary}\n\nВыполнить нарезку?",
            )
            if confirm:
                self._execute_slicing(idx, segments)

    def _open_segments_editor(
        self, idx: int, segments: list[dict], total_pages: int, error_msg: str,
    ):
        """Диалог ручной правки сегментов."""
        win = ctk.CTkToplevel(self)
        win.title("Правка сегментов нарезки")
        win.geometry("700x500")
        win.transient(self)
        win.grab_set()

        ctk.CTkLabel(
            win,
            text=f"Проблема: {error_msg}\nВсего страниц: {total_pages}",
            font=("", 12, "bold"),
            text_color="orange",
            justify="left",
        ).pack(padx=10, pady=10, anchor="w")

        # Таблица сегментов
        frame = ctk.CTkFrame(win)
        frame.pack(fill="both", expand=True, padx=10, pady=5)

        tree = ttk.Treeview(
            frame, columns=("from", "to", "type", "title"),
            show="headings", style="Custom.Treeview",
        )
        tree.heading("from", text="От стр.")
        tree.heading("to", text="До стр.")
        tree.heading("type", text="Тип")
        tree.heading("title", text="Название")
        tree.column("from", width=70)
        tree.column("to", width=70)
        tree.column("type", width=150)
        tree.column("title", width=350)
        tree.pack(fill="both", expand=True)

        def refresh():
            for iid in tree.get_children():
                tree.delete(iid)
            for i, s in enumerate(segments):
                tree.insert(
                    "", "end", iid=str(i),
                    values=(s.get("page_from", ""), s.get("page_to", ""),
                            s.get("doc_type", ""), s.get("title", "")),
                )

        refresh()

        status = ctk.CTkLabel(win, text="", text_color="lightgreen")
        status.pack(pady=5)

        def validate():
            ok, err = verify_segments(segments, total_pages)
            if ok:
                status.configure(text="Сегменты корректны ✓", text_color="lightgreen")
                apply_btn.configure(state="normal")
            else:
                status.configure(text=err, text_color="orange")
                apply_btn.configure(state="disabled")
            return ok

        def edit_cell(event):
            row = tree.identify_row(event.y)
            col = tree.identify_column(event.x)
            if not row:
                return
            col_idx = int(col.replace("#", "")) - 1
            col_names = ("page_from", "page_to", "doc_type", "title")
            col_name = col_names[col_idx]
            bbox = tree.bbox(row, col)
            if not bbox:
                return

            entry = tk.Entry(tree, font=("", 11))
            entry.insert(0, str(segments[int(row)].get(col_name, "")))
            entry.select_range(0, "end")
            entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
            entry.focus_set()

            def save(e=None):
                val = entry.get()
                if col_name in ("page_from", "page_to"):
                    try:
                        val = int(val)
                    except ValueError:
                        entry.destroy()
                        return
                segments[int(row)][col_name] = val
                entry.destroy()
                refresh()
                validate()

            entry.bind("<Return>", save)
            entry.bind("<FocusOut>", save)
            entry.bind("<Escape>", lambda e: entry.destroy())

        tree.bind("<Double-1>", edit_cell)

        # Кнопки управления сегментами
        btns = ctk.CTkFrame(win, fg_color="transparent")
        btns.pack(fill="x", padx=10, pady=5)

        def add_seg():
            segments.append({
                "doc_type": "Документ", "title": "",
                "page_from": 1, "page_to": 1,
            })
            refresh()
            validate()

        def remove_seg():
            sel = tree.selection()
            if not sel:
                return
            idx_ = int(sel[0])
            if 0 <= idx_ < len(segments):
                segments.pop(idx_)
            refresh()
            validate()

        ctk.CTkButton(btns, text="+ Добавить сегмент", command=add_seg).pack(side="left", padx=5)
        ctk.CTkButton(btns, text="− Удалить", command=remove_seg).pack(side="left", padx=5)

        bottom = ctk.CTkFrame(win, fg_color="transparent")
        bottom.pack(fill="x", padx=10, pady=10)

        def apply():
            if not validate():
                return
            win.destroy()
            self._execute_slicing(idx, segments)

        apply_btn = ctk.CTkButton(
            bottom, text="Продолжить нарезку", command=apply,
            fg_color="#28a745", hover_color="#218838", state="disabled",
        )
        apply_btn.pack(side="right", padx=5)

        ctk.CTkButton(
            bottom, text="Отмена", command=win.destroy,
        ).pack(side="right", padx=5)

        validate()

    def _execute_slicing(self, idx: int, segments: list[dict]):
        doc = self.results[idx]
        pdf_path = Path(doc["_file_path"])

        # Рабочая папка для нарезок — соседняя с исходной, чтобы НЕ трогать source.
        # Если source_dir неизвестен (режим выбора отдельных файлов) — кладём рядом с PDF.
        if self.source_dir:
            work_dir = self.source_dir.parent / f"{self.source_dir.name}_sliced"
        else:
            work_dir = pdf_path.parent / "_sliced"

        # Повторная нарезка ранее нарезанной части?
        is_resplit = bool(doc.get("_sliced_from"))
        root_path = doc.get("_sliced_from") or doc["_file_path"]

        self._set_statusbar("Нарезка PDF...")

        try:
            out_paths = slice_pdf(pdf_path, segments, work_dir)
        except Exception as e:
            messagebox.showerror("Ошибка нарезки", str(e))
            return

        new_paths_str = [str(p) for p in out_paths]

        if is_resplit:
            # Заменяем старую часть в _slice_parts корневого оригинала на новые
            root_doc = next(
                (d for d in self.results if d.get("_file_path") == root_path),
                None,
            )
            if root_doc is not None:
                old_part = doc["_file_path"]
                existing = root_doc.get("_slice_parts") or []
                root_doc["_slice_parts"] = [
                    p for p in existing if p != old_part
                ] + new_paths_str
            # Помечаем старую часть на удаление из results после анализа
            doc["_to_remove_after_slice"] = True
            # Удаляем файл старой части с диска
            try:
                Path(doc["_file_path"]).unlink(missing_ok=True)
            except Exception:
                pass
        else:
            # Первая нарезка: помечаем сам документ как разрезанный
            doc["_slice_parts"] = new_paths_str

        # Создаём записи для новых файлов
        self._set_statusbar("Анализ нарезанных частей...")

        new_file_infos = []
        for i, (p, seg) in enumerate(zip(out_paths, segments), start=1):
            try:
                rel = str(p.relative_to(work_dir))
            except ValueError:
                rel = p.name
            new_file_infos.append({
                "path": p,
                "name": p.name,
                "ext": ".pdf",
                "size": p.stat().st_size if p.exists() else 0,
                "rel_path": rel,
                "_preset_segment": seg,
            })

        def _progress(current, total):
            self.after(0, lambda: self._update_progress(current, total))

        def _run():
            loop = asyncio.new_event_loop()
            try:
                results = loop.run_until_complete(
                    analyze_batch(
                        new_file_infos,
                        self.cfg["api_key"],
                        self.cfg["vision_model"],
                        self.cfg["text_model"],
                        self.cfg.get("max_pages_per_pdf", 3),
                        self.cfg.get("max_concurrent", 5),
                        self.cfg.get("suspicious_page_threshold", 5),
                        progress_callback=_progress,
                    )
                )
                results = loop.run_until_complete(
                    group_documents(
                        results, self.categories,
                        self.cfg["api_key"], self.cfg["text_model"],
                        name_template=self.cfg.get("name_template", ""),
                    )
                )
                # Помечаем что они нарезаны из корневого оригинала
                for r in results:
                    r["_sliced_from"] = root_path
                    normalize_document(r)

                self.after(0, lambda: self._on_slicing_complete(results))
            except Exception as e:
                self.after(0, lambda: self._on_slice_error(str(e)))
            finally:
                loop.close()

        self._processing = True
        threading.Thread(target=_run, daemon=True).start()

    def _on_slicing_complete(self, new_results: list[dict]):
        # Удаляем старые части, помеченные при повторной нарезке
        self.results = [
            r for r in self.results if not r.get("_to_remove_after_slice")
        ]
        self.results.extend(new_results)
        self._init_categories_order()
        self._populate_tree()
        self._processing = False
        self._set_statusbar(f"Нарезка завершена. Создано частей: {len(new_results)}")
        self._schedule_autosave()

    def _on_slice_error(self, error: str):
        self._processing = False
        messagebox.showerror("Ошибка нарезки", error)
        self._set_statusbar("Ошибка нарезки")

    def _undo_slice_selected(self):
        indices = self._get_selected_docs()
        if len(indices) != 1:
            messagebox.showinfo("Информация", "Выберите нарезанный оригинал")
            return

        idx = indices[0]
        doc = self.results[idx]
        slice_parts = doc.get("_slice_parts")
        if not slice_parts:
            messagebox.showinfo("Информация", "Этот документ не нарезан")
            return

        confirm = messagebox.askyesno(
            "Отменить нарезку",
            f"Удалить {len(slice_parts)} нарезанных файлов и восстановить оригинал в копирование?",
        )
        if not confirm:
            return

        # Удаляем записи нарезок
        original_path = doc["_file_path"]
        self.results = [
            r for r in self.results if r.get("_sliced_from") != original_path
        ]

        # Удаляем файлы
        undo_slice(original_path, slice_parts)

        # Очищаем оригинал
        doc["_slice_parts"] = None

        self._init_categories_order()
        self._populate_tree()
        self._set_statusbar("Нарезка отменена")
        self._schedule_autosave()
