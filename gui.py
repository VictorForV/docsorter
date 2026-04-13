"""
GUI модуль DocSorter на customtkinter.
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
)
from scanner import scan_folder
from analyzer import analyze_batch
from grouper import group_documents, regroup_documents, generate_categories
from sorter import (
    build_folder_structure, build_numbering_structure,
    execute_sort, verify_sort, sanitize_filename,
)


ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class DocSorterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("DocSorter — Сортировщик документов")
        self.geometry("1200x800")
        self.minsize(900, 600)

        self.cfg = load_config()
        self.categories = load_categories()
        self.results = []  # Результаты анализа
        self.source_dir = None
        self.output_dir = None
        self._processing = False

        self._build_ui()
        self._update_status()

    # ── UI ──────────────────────────────────────────────────────────

    def _build_ui(self):
        # Верхняя панель
        top = ctk.CTkFrame(self)
        top.pack(fill="x", padx=10, pady=(10, 5))

        self.status_label = ctk.CTkLabel(top, text="", font=("", 14))
        self.status_label.pack(side="left", padx=10)

        settings_btn = ctk.CTkButton(
            top, text="Настройки", width=100, command=self._open_settings,
        )
        settings_btn.pack(side="right", padx=5)

        cats_btn = ctk.CTkButton(
            top, text="Категории", width=100, command=self._open_categories,
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

        # Режим сортировки + кнопка анализа
        controls = ctk.CTkFrame(self)
        controls.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(controls, text="Режим:").pack(side="left", padx=(10, 5))

        self.sort_mode_var = ctk.StringVar(value=self.cfg.get("sort_mode", "folders"))
        mode_folders = ctk.CTkRadioButton(
            controls, text="По папкам", variable=self.sort_mode_var, value="folders",
        )
        mode_folders.pack(side="left", padx=5)

        mode_numbers = ctk.CTkRadioButton(
            controls, text="По нумерации", variable=self.sort_mode_var, value="numbering",
        )
        mode_numbers.pack(side="left", padx=5)

        self.analyze_btn = ctk.CTkButton(
            controls, text="Начать анализ", width=150,
            fg_color="green", hover_color="darkgreen",
            command=self._start_analysis,
        )
        self.analyze_btn.pack(side="right", padx=10)

        # Прогресс
        self.progress_frame = ctk.CTkFrame(self)
        self.progress_frame.pack(fill="x", padx=10, pady=5)

        self.progress_label = ctk.CTkLabel(self.progress_frame, text="")
        self.progress_label.pack(side="left", padx=10)

        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=10)
        self.progress_bar.set(0)

        # Таблица результатов
        table_frame = ctk.CTkFrame(self)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Используем ttk.Treeview для таблицы (customtkinter не имеет своей)
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Custom.Treeview",
            background="#2b2b2b",
            foreground="white",
            fieldbackground="#2b2b2b",
            rowheight=28,
            font=("", 11),
        )
        style.configure(
            "Custom.Treeview.Heading",
            background="#1f538d",
            foreground="white",
            font=("", 11, "bold"),
        )
        style.map("Custom.Treeview", background=[("selected", "#1f538d")])

        columns = ("num", "file", "type", "title", "number", "date", "counterparty", "category", "group", "new_name")
        self.tree = ttk.Treeview(
            table_frame, columns=columns, show="headings",
            style="Custom.Treeview", selectmode="extended",
        )

        headings = {
            "num": ("#", 40),
            "file": ("Исходный файл", 150),
            "type": ("Тип", 100),
            "title": ("Название", 200),
            "number": ("Номер", 70),
            "date": ("Дата", 90),
            "counterparty": ("Контрагент", 130),
            "category": ("Категория", 120),
            "group": ("Группа", 150),
            "new_name": ("Новое имя", 200),
        }
        for col, (heading, width) in headings.items():
            self.tree.heading(col, text=heading)
            self.tree.column(col, width=width, minwidth=40)

        scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")

        # Двойной клик для редактирования
        self.tree.bind("<Double-1>", self._on_double_click)

        # Нижняя панель: промпт + действия
        bottom = ctk.CTkFrame(self)
        bottom.pack(fill="x", padx=10, pady=(5, 10))

        ctk.CTkLabel(bottom, text="Промпт:").pack(side="left", padx=(10, 5))

        self.prompt_var = ctk.StringVar()
        self.prompt_entry = ctk.CTkEntry(
            bottom, textvariable=self.prompt_var, width=400,
            placeholder_text="Перегруппируй: выдели судебные документы отдельно...",
        )
        self.prompt_entry.pack(side="left", fill="x", expand=True, padx=5)

        self.regroup_btn = ctk.CTkButton(
            bottom, text="Применить", width=100, command=self._regroup,
        )
        self.regroup_btn.pack(side="left", padx=5)

        self.sort_btn = ctk.CTkButton(
            bottom, text="Сортировать файлы", width=150,
            fg_color="#28a745", hover_color="#218838",
            command=self._execute_sort,
        )
        self.sort_btn.pack(side="right", padx=10)

        export_btn = ctk.CTkButton(
            bottom, text="Экспорт CSV", width=100, command=self._export_csv,
        )
        export_btn.pack(side="right", padx=5)

        # Статусбар
        self.statusbar = ctk.CTkLabel(
            self, text="Готов к работе", anchor="w", font=("", 11),
        )
        self.statusbar.pack(fill="x", padx=10, pady=(0, 5))

    # ── Статус ─────────────────────────────────────────────────────

    def _update_status(self):
        if is_config_valid(self.cfg):
            self.status_label.configure(
                text="API настроен",
                text_color="lightgreen",
            )
        else:
            self.status_label.configure(
                text="API не настроен — откройте Настройки",
                text_color="orange",
            )

    def _set_statusbar(self, text: str):
        self.statusbar.configure(text=text)

    # ── Настройки ──────────────────────────────────────────────────

    def _open_settings(self):
        win = ctk.CTkToplevel(self)
        win.title("Настройки")
        win.geometry("500x350")
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

        def _save():
            for key, entry in fields.items():
                val = entry.get().strip()
                if key in ("max_concurrent", "max_pages_per_pdf"):
                    try:
                        val = int(val)
                    except ValueError:
                        val = self.cfg.get(key, 5)
                self.cfg[key] = val
            save_config(self.cfg)
            self._update_status()
            win.destroy()

        ctk.CTkButton(win, text="Сохранить", command=_save).grid(
            row=row, column=0, columnspan=2, pady=15,
        )

    # ── Категории ──────────────────────────────────────────────────

    def _open_categories(self):
        win = ctk.CTkToplevel(self)
        win.title("Категории документов")
        win.geometry("600x500")
        win.transient(self)
        win.grab_set()

        ctk.CTkLabel(
            win, text=f"Набор: {self.categories.get('name', '—')}",
            font=("", 14, "bold"),
        ).pack(padx=10, pady=10)

        text_widget = ctk.CTkTextbox(win, width=560, height=300)
        text_widget.pack(padx=10, pady=5)
        text_widget.insert("1.0", json.dumps(self.categories, ensure_ascii=False, indent=2))

        # Промпт для генерации
        prompt_frame = ctk.CTkFrame(win)
        prompt_frame.pack(fill="x", padx=10, pady=5)

        gen_entry = ctk.CTkEntry(
            prompt_frame, width=400,
            placeholder_text="Сгенерировать категории: арбитражные дела, налоговые...",
        )
        gen_entry.pack(side="left", fill="x", expand=True, padx=5)

        def _generate():
            prompt_text = gen_entry.get().strip()
            if not prompt_text:
                return
            if not is_config_valid(self.cfg):
                messagebox.showwarning("Ошибка", "Настройте API ключ")
                return

            def _run():
                loop = asyncio.new_event_loop()
                try:
                    new_cats = loop.run_until_complete(
                        generate_categories(prompt_text, self.cfg["api_key"], self.cfg["text_model"])
                    )
                    self.after(0, lambda: _update_text(new_cats))
                except Exception as e:
                    self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
                finally:
                    loop.close()

            threading.Thread(target=_run, daemon=True).start()

        def _update_text(new_cats):
            text_widget.delete("1.0", "end")
            text_widget.insert("1.0", json.dumps(new_cats, ensure_ascii=False, indent=2))

        ctk.CTkButton(prompt_frame, text="Сгенерировать", width=120, command=_generate).pack(
            side="left", padx=5,
        )

        def _save_cats():
            try:
                raw = text_widget.get("1.0", "end").strip()
                new_cats = json.loads(raw)
                self.categories = new_cats
                save_categories(new_cats)
                win.destroy()
            except json.JSONDecodeError:
                messagebox.showerror("Ошибка", "Некорректный JSON")

        ctk.CTkButton(win, text="Сохранить", command=_save_cats).pack(pady=10)

    # ── Выбор папки ────────────────────────────────────────────────

    def _browse_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку с документами")
        if folder:
            self.folder_var.set(folder)

    # ── Анализ ─────────────────────────────────────────────────────

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

        self.source_dir = Path(folder)
        if not self.source_dir.exists():
            messagebox.showerror("Ошибка", "Папка не найдена")
            return

        # Сканируем
        try:
            files = scan_folder(self.source_dir)
        except Exception as e:
            messagebox.showerror("Ошибка сканирования", str(e))
            return

        if not files:
            messagebox.showinfo("Информация", "В папке нет поддерживаемых файлов")
            return

        self._set_statusbar(f"Найдено файлов: {len(files)}")
        self._processing = True
        self.analyze_btn.configure(state="disabled", text="Анализ...")
        self.progress_bar.set(0)
        self.progress_label.configure(text=f"0/{len(files)}")

        # Очищаем таблицу
        for item in self.tree.get_children():
            self.tree.delete(item)

        def _progress(current, total):
            self.after(0, lambda: self._update_progress(current, total))

        def _run():
            loop = asyncio.new_event_loop()
            try:
                # Шаг 1: Анализ файлов
                results = loop.run_until_complete(
                    analyze_batch(
                        files,
                        self.cfg["api_key"],
                        self.cfg["vision_model"],
                        self.cfg["text_model"],
                        self.cfg.get("max_pages_per_pdf", 3),
                        self.cfg.get("max_concurrent", 5),
                        progress_callback=_progress,
                    )
                )

                self.after(0, lambda: self._set_statusbar("Группировка документов..."))

                # Шаг 2: Группировка
                results = loop.run_until_complete(
                    group_documents(
                        results, self.categories,
                        self.cfg["api_key"], self.cfg["text_model"],
                    )
                )

                self.after(0, lambda: self._on_analysis_complete(results))

            except Exception as e:
                self.after(0, lambda: self._on_analysis_error(str(e)))
            finally:
                loop.close()

        threading.Thread(target=_run, daemon=True).start()

    def _update_progress(self, current, total):
        self.progress_bar.set(current / total)
        self.progress_label.configure(text=f"{current}/{total}")

    def _on_analysis_complete(self, results):
        self.results = results
        self._processing = False
        self.analyze_btn.configure(state="normal", text="Начать анализ")
        self.progress_bar.set(1.0)
        self._populate_table()
        self._set_statusbar(f"Анализ завершён. Документов: {len(results)}")

    def _on_analysis_error(self, error: str):
        self._processing = False
        self.analyze_btn.configure(state="normal", text="Начать анализ")
        messagebox.showerror("Ошибка анализа", error)
        self._set_statusbar("Ошибка анализа")

    # ── Таблица ────────────────────────────────────────────────────

    def _populate_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Сортируем по категории и sort_order
        sorted_results = sorted(
            self.results,
            key=lambda d: (d.get("_category", ""), d.get("_sort_order", 99)),
        )
        self.results = sorted_results

        for i, doc in enumerate(self.results):
            values = (
                i + 1,
                doc.get("_file_name", ""),
                doc.get("doc_type", ""),
                doc.get("title", ""),
                doc.get("number", ""),
                doc.get("date", ""),
                doc.get("counterparty", ""),
                doc.get("_category", ""),
                doc.get("_group", ""),
                doc.get("_new_name", ""),
            )
            self.tree.insert("", "end", iid=str(i), values=values)

    def _on_double_click(self, event):
        """Редактирование ячейки по двойному клику."""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        col = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return

        col_index = int(col.replace("#", "")) - 1
        columns = ("num", "file", "type", "title", "number", "date", "counterparty", "category", "group", "new_name")

        if col_index == 0 or col_index == 1:  # num и file не редактируются
            return

        col_name = columns[col_index]

        # Получаем bbox ячейки
        bbox = self.tree.bbox(row_id, col)
        if not bbox:
            return

        current_value = self.tree.set(row_id, col_name)

        # Создаём Entry поверх ячейки
        entry = tk.Entry(self.tree, font=("", 11))
        entry.insert(0, current_value)
        entry.select_range(0, "end")
        entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        entry.focus_set()

        def _save_edit(e=None):
            new_val = entry.get()
            self.tree.set(row_id, col_name, new_val)
            entry.destroy()

            # Обновляем данные
            idx = int(row_id)
            field_map = {
                "type": "doc_type",
                "title": "title",
                "number": "number",
                "date": "date",
                "counterparty": "counterparty",
                "category": "_category",
                "group": "_group",
                "new_name": "_new_name",
            }
            data_key = field_map.get(col_name)
            if data_key and 0 <= idx < len(self.results):
                self.results[idx][data_key] = new_val

        def _cancel_edit(e=None):
            entry.destroy()

        entry.bind("<Return>", _save_edit)
        entry.bind("<Escape>", _cancel_edit)
        entry.bind("<FocusOut>", _save_edit)

    # ── Перегруппировка ────────────────────────────────────────────

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
        self._populate_table()
        self.regroup_btn.configure(state="normal")
        self._set_statusbar("Перегруппировка завершена")

    def _on_regroup_error(self, error: str):
        self.regroup_btn.configure(state="normal")
        messagebox.showerror("Ошибка", error)
        self._set_statusbar("Ошибка перегруппировки")

    # ── Сортировка (копирование) ───────────────────────────────────

    def _execute_sort(self):
        if not self.results:
            messagebox.showinfo("Информация", "Сначала проведите анализ")
            return

        # Выбираем папку назначения
        output = filedialog.askdirectory(title="Выберите папку для отсортированных файлов")
        if not output:
            return

        self.output_dir = Path(output)

        # Строим структуру
        mode = self.sort_mode_var.get()
        if mode == "folders":
            build_folder_structure(self.results, self.output_dir)
        else:
            build_numbering_structure(self.results, self.output_dir)

        # Подтверждение
        confirm = messagebox.askyesno(
            "Подтверждение",
            f"Скопировать {len(self.results)} файлов в:\n{self.output_dir}\n\nПродолжить?",
        )
        if not confirm:
            return

        # Копируем
        self._set_statusbar("Копирование файлов...")
        result = execute_sort(self.results, self.output_dir)

        # Верификация
        verification = verify_sort(self.source_dir, self.output_dir)

        # Отчёт
        msg = (
            f"Скопировано: {result['copied']} файлов\n"
            f"Ошибок: {len(result['errors'])}\n\n"
            f"Проверка:\n"
            f"  Исходных файлов: {verification['source_count']}\n"
            f"  Скопированных: {verification['dest_count']}\n"
            f"  Совпадение: {'Да' if verification['match'] else 'НЕТ!'}"
        )

        if result["errors"]:
            msg += "\n\nОшибки:\n" + "\n".join(result["errors"][:10])

        if verification["match"]:
            messagebox.showinfo("Готово", msg)
            self._set_statusbar("Сортировка завершена успешно")
        else:
            messagebox.showwarning("Внимание", msg)
            self._set_statusbar("Сортировка завершена с расхождениями")

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
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow([
                "№", "Исходный файл", "Тип", "Название", "Номер",
                "Дата", "Контрагент", "Категория", "Группа", "Новое имя", "Содержание",
            ])
            for i, doc in enumerate(self.results):
                writer.writerow([
                    i + 1,
                    doc.get("_file_name", ""),
                    doc.get("doc_type", ""),
                    doc.get("title", ""),
                    doc.get("number", ""),
                    doc.get("date", ""),
                    doc.get("counterparty", ""),
                    doc.get("_category", ""),
                    doc.get("_group", ""),
                    doc.get("_new_name", ""),
                    doc.get("summary", ""),
                ])

        self._set_statusbar(f"Экспорт: {path}")
        messagebox.showinfo("Экспорт", f"Данные сохранены:\n{path}")
