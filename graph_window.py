"""
Модуль визуализации графа связей документов.
Показывает явные ссылки (reference_number → number) между документами.
Несуществующие документы отображаются серыми «призрачными» нодами.
"""

import tkinter as tk
from collections import defaultdict

import customtkinter as ctk
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import networkx as nx

from linker import build_indexes, _normalize_number
from doctypes import get_type_to_category


# ── Цвета по категории документа ────────────────────────────────────

_DOC_TYPE_TO_CATEGORY = get_type_to_category()

CATEGORY_COLORS = {
    "Договоры": "#4a90d9",
    "Первичная документация": "#f0a030",
    "Корпоративные документы": "#a060d0",
    "Судебные документы": "#d04040",
    "Переписка": "#50b858",
    "Финансовые документы": "#d0a040",
    "Документы надзорных органов": "#808080",
    "Прочее": "#707070",
}
GHOST_COLOR = "#555555"
BG_COLOR = "#1c1c2a"
TEXT_COLOR = "#d8d8e8"
EDGE_COLOR = "#555577"


class GraphWindow:
    """Окно графа связей документов."""

    def __init__(self, parent, results: list[dict],
                 link_overrides: dict, on_save_overrides: callable):
        self.parent = parent
        self.results = results
        self.link_overrides = dict(link_overrides)
        self.on_save_overrides = on_save_overrides

        self.graph = nx.DiGraph()
        self.pos = {}          # node_id → (x, y)
        self.ghost_nodes = {}  # ghost_key → info dict
        self._dragging = None
        self._pick_radius = 0.08

        self._build_ui()
        self._build_graph()
        self._draw()

    # ── UI ────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.win = ctk.CTkToplevel(self.parent)
        self.win.title("Граф связей документов")
        self.win.geometry("1100x750")
        self.win.transient(self.parent)

        # Тулбар
        toolbar = ctk.CTkFrame(self.win, fg_color="transparent")
        toolbar.pack(fill="x", padx=10, pady=(10, 5))

        ctk.CTkLabel(
            toolbar, text="Граф явных ссылок между документами",
            font=("", 13, "bold"), text_color=TEXT_COLOR,
        ).pack(side="left")

        ctk.CTkButton(
            toolbar, text="Закрыть", width=80,
            command=self.win.destroy,
        ).pack(side="right", padx=5)

        ctk.CTkButton(
            toolbar, text="Обновить", width=80,
            command=self._rebuild_and_draw,
        ).pack(side="right", padx=5)

        # Легенда
        legend = ctk.CTkFrame(self.win, fg_color="transparent")
        legend.pack(fill="x", padx=10, pady=(0, 3))
        for cat, color in CATEGORY_COLORS.items():
            ctk.CTkLabel(
                legend, text=f"  ● {cat}  ", font=("", 10),
                text_color=color,
            ).pack(side="left")
        ctk.CTkLabel(
            legend, text="  ● Не найден  ", font=("", 10),
            text_color=GHOST_COLOR,
        ).pack(side="left")

        # Matplotlib
        self.fig = Figure(figsize=(11, 6.5), dpi=100, facecolor=BG_COLOR)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_facecolor(BG_COLOR)
        self.ax.axis("off")
        self.fig.subplots_adjust(left=0.02, right=0.98, top=0.98, bottom=0.02)

        canvas_frame = ctk.CTkFrame(self.win, fg_color=BG_COLOR)
        canvas_frame.pack(fill="both", expand=True, padx=10, pady=(0, 5))

        self.canvas = FigureCanvasTkAgg(self.fig, master=canvas_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        nav_frame = ctk.CTkFrame(canvas_frame, fg_color="transparent")
        nav_frame.pack(fill="x")
        self.nav_toolbar = NavigationToolbar2Tk(self.canvas, nav_frame)
        self.nav_toolbar.update()

        # События
        self.canvas.mpl_connect("button_press_event", self._on_press)
        self.canvas.mpl_connect("motion_notify_event", self._on_motion)
        self.canvas.mpl_connect("button_release_event", self._on_release)
        self.canvas.mpl_connect("scroll_event", self._on_scroll)

    # ── Построение графа ──────────────────────────────────────────────

    def _build_graph(self):
        G = nx.DiGraph()
        indexes = build_indexes(self.results)
        self.ghost_nodes.clear()

        # Собираем все документы, участвующие в ссылках
        linked_docs = set()  # индексы документов, у которых есть связи

        # Шаг 1: для каждого документа с reference_number ищем цель
        for i, doc in enumerate(self.results):
            ref_num_raw = doc.get("reference_number", "")
            if not ref_num_raw:
                continue

            ref_num = _normalize_number(ref_num_raw)
            ref_date = doc.get("reference_date", "").strip()
            ghost_key = f"ghost:{ref_num}:{ref_date}"

            # Проверяем override
            override = self.link_overrides.get(ghost_key)
            if override and override.get("action") == "link":
                j = override["linked_doc_index"]
                if 0 <= j < len(self.results):
                    linked_docs.add(i)
                    linked_docs.add(j)
                    G.add_node(f"doc:{i}", **self._doc_node_attr(i))
                    G.add_node(f"doc:{j}", **self._doc_node_attr(j))
                    G.add_edge(f"doc:{i}", f"doc:{j}")
                    continue

            # Ищем в индексах
            key = (ref_num, ref_date)
            matched = indexes["by_number_date"].get(key, [])
            if not matched and not ref_date:
                for (n, _d), idxs in indexes["by_number_date"].items():
                    if n == ref_num:
                        matched = idxs
                        break

            if matched:
                for j in matched:
                    if j != i:
                        linked_docs.add(i)
                        linked_docs.add(j)
                        G.add_node(f"doc:{i}", **self._doc_node_attr(i))
                        G.add_node(f"doc:{j}", **self._doc_node_attr(j))
                        G.add_edge(f"doc:{i}", f"doc:{j}")
            else:
                # Призрачная нода
                label = f"№{ref_num_raw.strip()}"
                if ref_date:
                    label += f" от {ref_date}"
                ghost_data = {
                    "label": label,
                    "ref_number": ref_num_raw.strip(),
                    "ref_date": ref_date,
                    "normalized_number": ref_num,
                    "referenced_by": [],
                }
                if ghost_key not in self.ghost_nodes:
                    self.ghost_nodes[ghost_key] = ghost_data
                    G.add_node(ghost_key, ghost=True, label=label)
                self.ghost_nodes[ghost_key]["referenced_by"].append(i)
                linked_docs.add(i)
                G.add_node(f"doc:{i}", **self._doc_node_attr(i))
                G.add_edge(f"doc:{i}", ghost_key)

        self.graph = G

        # Layout
        if G.number_of_nodes() == 0:
            self.pos = {}
            return

        try:
            self.pos = nx.spring_layout(
                G, seed=42,
                k=2.5 / (G.number_of_nodes() ** 0.5),
                iterations=100,
            )
        except Exception:
            self.pos = nx.spring_layout(G, seed=42)

    def _doc_node_attr(self, idx: int) -> dict:
        doc = self.results[idx]
        doc_type = doc.get("doc_type", "")
        cat, _ = _DOC_TYPE_TO_CATEGORY.get(doc_type, ("Прочее", ""))
        color = CATEGORY_COLORS.get(cat, "#707070")
        # Формируем информативную подпись: Тип\n№Номер от Дата
        parts = [doc_type or doc.get("_file_name", "?")]
        num = doc.get("number", "").strip()
        dt = doc.get("date", "").strip()
        second_line = ""
        if num:
            second_line += f"№{num}"
        if dt:
            second_line += f" от {dt}" if second_line else dt
        if second_line:
            parts.append(second_line)
        label = "\n".join(parts)
        return {"ghost": False, "color": color, "label": label, "doc_index": idx}

    # ── Отрисовка ─────────────────────────────────────────────────────

    def _draw(self):
        self.ax.clear()
        self.ax.set_facecolor(BG_COLOR)
        self.ax.axis("off")

        G = self.graph
        if G.number_of_nodes() == 0:
            self.ax.text(
                0.5, 0.5, "Нет связей между документами",
                ha="center", va="center", fontsize=14, color=TEXT_COLOR,
                transform=self.ax.transAxes,
            )
            self.canvas.draw_idle()
            return

        # Цвета и размеры
        node_colors = []
        node_sizes = []
        node_labels = {}
        for node_id in G.nodes():
            data = G.nodes[node_id]
            if data.get("ghost"):
                node_colors.append(GHOST_COLOR)
                node_sizes.append(1200)
                node_labels[node_id] = data.get("label", "?")
            else:
                node_colors.append(data.get("color", "#707070"))
                deg = G.degree(node_id)
                node_sizes.append(max(1200, 600 + deg * 300))
                node_labels[node_id] = data.get("label", "?")

        # Рёбра
        nx.draw_networkx_edges(
            G, self.pos, ax=self.ax,
            edge_color=EDGE_COLOR, arrows=True,
            arrowsize=18, arrowstyle="-|>",
            connectionstyle="arc3,rad=0.1",
            width=2.0, alpha=0.7,
        )

        # Ноды
        nx.draw_networkx_nodes(
            G, self.pos, ax=self.ax,
            node_color=node_colors, node_size=node_sizes,
            edgecolors="#333344",
            linewidths=1.5, alpha=0.9,
        )

        # Призрачные ноды — пунктирная обводка
        ghost_ids = [n for n in G.nodes if G.nodes[n].get("ghost")]
        if ghost_ids:
            ghost_pos = {n: self.pos[n] for n in ghost_ids}
            ghost_sizes = [1200] * len(ghost_ids)
            nx.draw_networkx_nodes(
                G, ghost_pos, ax=self.ax, nodelist=ghost_ids,
                node_color="none", node_size=ghost_sizes,
                edgecolors="#888888", linewidths=2.0, linestyle="dashed",
            )

        # Подписи с фоном для читаемости
        bbox_props = dict(
            boxstyle="round,pad=0.3",
            facecolor=BG_COLOR,
            edgecolor="none",
            alpha=0.85,
        )
        for node_id, label in node_labels.items():
            x, y = self.pos[node_id]
            data = G.nodes[node_id]
            font_color = "#999999" if data.get("ghost") else TEXT_COLOR
            self.ax.text(
                x, y, label,
                ha="center", va="center",
                fontsize=8, fontfamily="sans-serif",
                fontweight="bold",
                color=font_color,
                bbox=bbox_props,
            )

        self.canvas.draw_idle()

    def _rebuild_and_draw(self):
        self._build_graph()
        self._draw()

    # ── Интерактивность ───────────────────────────────────────────────

    def _on_press(self, event):
        if event.inaxes != self.ax or event.xdata is None:
            return

        if event.button == 3:
            # Правый клик — контекстное меню для призрачных нод
            ghost_id = self._find_ghost_at(event.xdata, event.ydata)
            if ghost_id:
                self._show_ghost_menu(ghost_id, event)
            return

        if event.button == 1:
            # Левый клик — начало перетаскивания
            node = self._find_node_at(event.xdata, event.ydata)
            if node:
                self._dragging = node

    def _on_motion(self, event):
        if self._dragging is None or event.inaxes != self.ax or event.xdata is None:
            return
        self.pos[self._dragging] = (event.xdata, event.ydata)
        self._draw()

    def _on_release(self, event):
        self._dragging = None

    def _on_scroll(self, event):
        if event.xdata is None or event.ydata is None:
            return
        scale = 1.2
        if event.button == "up":
            scale = 1 / scale
        xlim = self.ax.get_xlim()
        ylim = self.ax.get_ylim()
        x, y = event.xdata, event.ydata
        self.ax.set_xlim(x - (x - xlim[0]) * scale, x + (xlim[1] - x) * scale)
        self.ax.set_ylim(y - (y - ylim[0]) * scale, y + (ylim[1] - y) * scale)
        self.canvas.draw_idle()

    def _find_node_at(self, x, y) -> str | None:
        if not self.pos:
            return None
        best_dist = self._pick_radius
        best_node = None
        for node_id, (nx_, ny) in self.pos.items():
            d = ((x - nx_) ** 2 + (y - ny) ** 2) ** 0.5
            if d < best_dist:
                best_dist = d
                best_node = node_id
        return best_node

    def _find_ghost_at(self, x, y) -> str | None:
        node = self._find_node_at(x, y)
        if node and self.graph.nodes[node].get("ghost"):
            return node
        return None

    # ── Контекстное меню для призрачных нод ───────────────────────────

    def _show_ghost_menu(self, ghost_id: str, event):
        menu = tk.Menu(
            self.win, tearoff=0,
            bg="#2b2b2b", fg="white",
            activebackground="#1f538d", activeforeground="white",
        )
        menu.add_command(
            label="Привязать к документу...",
            command=lambda: self._link_ghost_to_real(ghost_id),
        )
        menu.add_command(
            label="Изменить ссылку...",
            command=lambda: self._edit_ghost_reference(ghost_id),
        )
        try:
            menu.tk_popup(event.guiEvent.x_root, event.guiEvent.y_root)
        finally:
            menu.grab_release()

    def _link_ghost_to_real(self, ghost_id: str):
        """Диалог выбора реального документа для призрачной ноды."""
        sel_win = ctk.CTkToplevel(self.win)
        sel_win.title("Выберите документ")
        sel_win.geometry("650x450")
        sel_win.transient(self.win)
        sel_win.grab_set()

        ghost = self.ghost_nodes.get(ghost_id, {})
        ctk.CTkLabel(
            sel_win,
            text=f"Привязать «{ghost.get('label', '?')}» к документу:",
            font=("", 12, "bold"), text_color=TEXT_COLOR,
        ).pack(padx=10, pady=(10, 5), anchor="w")

        # Фильтр
        filter_var = ctk.StringVar()
        filter_entry = ctk.CTkEntry(sel_win, textvariable=filter_var,
                                    placeholder_text="Поиск по типу / номеру / имени...")
        filter_entry.pack(fill="x", padx=10, pady=(0, 5))

        scroll = ctk.CTkScrollableFrame(sel_win)
        scroll.pack(fill="both", expand=True, padx=10, pady=5)

        buttons = []

        def _populate_filter(*_):
            for b in buttons:
                b.destroy()
            buttons.clear()
            q = filter_var.get().lower()
            for i, doc in enumerate(self.results):
                text = (
                    f"{doc.get('doc_type', '?')}  "
                    f"№{doc.get('number', '')}  "
                    f"от {doc.get('date', '')}  "
                    f"— {doc.get('_file_name', '')}"
                )
                if q and q not in text.lower():
                    continue
                btn = ctk.CTkButton(
                    scroll, text=text, anchor="w", height=28,
                    font=("", 11),
                    command=lambda idx=i: self._apply_link_override(
                        ghost_id, idx, sel_win,
                    ),
                )
                btn.pack(fill="x", pady=1)
                buttons.append(btn)

        filter_var.trace_add("write", _populate_filter)
        _populate_filter()

        ctk.CTkButton(
            sel_win, text="Отмена", command=sel_win.destroy, width=100,
        ).pack(pady=10)

    def _apply_link_override(self, ghost_id: str, doc_index: int, dialog):
        self.link_overrides[ghost_id] = {
            "action": "link",
            "linked_doc_index": doc_index,
        }
        dialog.destroy()
        self.on_save_overrides(self.link_overrides)
        self._rebuild_and_draw()

    def _edit_ghost_reference(self, ghost_id: str):
        """Диалог редактирования номера/даты ссылки."""
        ghost = self.ghost_nodes.get(ghost_id, {})

        edit_win = ctk.CTkToplevel(self.win)
        edit_win.title("Изменить ссылку")
        edit_win.geometry("400x220")
        edit_win.transient(self.win)
        edit_win.grab_set()

        ctk.CTkLabel(
            edit_win, text="Номер документа-ссылки:",
            font=("", 11), text_color=TEXT_COLOR,
        ).pack(padx=15, pady=(15, 2), anchor="w")

        num_entry = ctk.CTkEntry(edit_win, width=350)
        num_entry.insert(0, ghost.get("ref_number", ""))
        num_entry.pack(padx=15, pady=(0, 8))

        ctk.CTkLabel(
            edit_win, text="Дата документа-ссылки:",
            font=("", 11), text_color=TEXT_COLOR,
        ).pack(padx=15, pady=(0, 2), anchor="w")

        date_entry = ctk.CTkEntry(edit_win, width=350)
        date_entry.insert(0, ghost.get("ref_date", ""))
        date_entry.pack(padx=15, pady=(0, 8))

        def _apply():
            new_num = num_entry.get().strip()
            new_date = date_entry.get().strip()
            # Обновляем reference в исходных документах
            for src_idx in ghost.get("referenced_by", []):
                if 0 <= src_idx < len(self.results):
                    self.results[src_idx]["reference_number"] = new_num
                    self.results[src_idx]["reference_date"] = new_date
            # Удаляем старый override (ссылка изменилась)
            self.link_overrides.pop(ghost_id, None)
            edit_win.destroy()
            self.on_save_overrides(self.link_overrides)
            self._rebuild_and_draw()

        btn_frame = ctk.CTkFrame(edit_win, fg_color="transparent")
        btn_frame.pack(pady=10)
        ctk.CTkButton(btn_frame, text="Применить", command=_apply, width=100).pack(
            side="left", padx=5,
        )
        ctk.CTkButton(
            btn_frame, text="Отмена", command=edit_win.destroy, width=100,
            fg_color="#555555",
        ).pack(side="left", padx=5)
