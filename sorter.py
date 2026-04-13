"""
Модуль сортировки (копирования) файлов.
Создаёт новую структуру папок и копирует файлы с переименованием.
"""

import re
import shutil
from pathlib import Path
from collections import Counter


def sanitize_filename(name: str) -> str:
    """Убирает недопустимые символы из имени файла."""
    # Заменяем запрещённые символы
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    # Убираем множественные пробелы/подчёркивания
    name = re.sub(r'[_\s]+', ' ', name)
    name = name.strip('. ')
    # Ограничиваем длину (Windows — 255 символов для имени)
    if len(name) > 200:
        name = name[:200]
    return name


def build_folder_structure(results: list[dict], output_dir: Path) -> list[dict]:
    """
    Режим 'Папки': строит структуру папок по категориям и группам.
    Возвращает список с заполненным полем _dest_path.
    """
    # Собираем уникальные категории в порядке появления
    categories_order = []
    seen_cats = set()
    for doc in results:
        cat = doc.get("_category", "Прочее")
        if cat not in seen_cats:
            categories_order.append(cat)
            seen_cats.add(cat)

    # Нумеруем категории
    cat_nums = {cat: f"{i+1:02d}" for i, cat in enumerate(categories_order)}

    # Отслеживаем дубли имён
    name_counter = Counter()

    for doc in results:
        cat = doc.get("_category", "Прочее")
        cat_num = cat_nums.get(cat, "99")
        folder_name = f"{cat_num}_{sanitize_filename(cat)}"

        # Группа (подпапка)
        group = doc.get("_group", "")
        if group:
            group_folder = sanitize_filename(group)
        else:
            group_folder = ""

        # Имя файла
        new_name = doc.get("_new_name", "документ")
        new_name = sanitize_filename(new_name)
        ext = doc.get("_ext", "")

        full_name = f"{new_name}{ext}"

        # Проверяем дубли
        if group_folder:
            dest_dir = output_dir / folder_name / group_folder
        else:
            dest_dir = output_dir / folder_name

        dest_path = dest_dir / full_name
        key = str(dest_path).lower()
        name_counter[key] += 1
        if name_counter[key] > 1:
            full_name = f"{new_name} (дубль {name_counter[key] - 1}){ext}"
            dest_path = dest_dir / full_name

        doc["_dest_path"] = str(dest_path)

    return results


def build_numbering_structure(results: list[dict], output_dir: Path) -> list[dict]:
    """
    Режим 'Нумерация': плоская структура с нумерацией вида 1.1, 1.2, 2.1, ...
    """
    # Группируем по категориям
    categories_order = []
    seen_cats = set()
    for doc in results:
        cat = doc.get("_category", "Прочее")
        if cat not in seen_cats:
            categories_order.append(cat)
            seen_cats.add(cat)

    # Нумеруем
    cat_nums = {cat: i + 1 for i, cat in enumerate(categories_order)}
    cat_counters = {cat: 0 for cat in categories_order}

    # Сортируем внутри каждой категории по sort_order
    sorted_results = sorted(
        results,
        key=lambda d: (
            list(categories_order).index(d.get("_category", "Прочее"))
            if d.get("_category", "Прочее") in categories_order else 999,
            d.get("_sort_order", 99),
        ),
    )

    name_counter = Counter()

    for doc in sorted_results:
        cat = doc.get("_category", "Прочее")
        cat_num = cat_nums.get(cat, 99)
        cat_counters[cat] = cat_counters.get(cat, 0) + 1
        sub_num = cat_counters[cat]

        new_name = doc.get("_new_name", "документ")
        new_name = sanitize_filename(new_name)
        ext = doc.get("_ext", "")

        full_name = f"{cat_num}.{sub_num} — {new_name}{ext}"

        dest_path = output_dir / full_name
        key = str(dest_path).lower()
        name_counter[key] += 1
        if name_counter[key] > 1:
            full_name = f"{cat_num}.{sub_num} — {new_name} (дубль {name_counter[key] - 1}){ext}"
            dest_path = output_dir / full_name

        doc["_dest_path"] = str(dest_path)

    return results


def execute_sort(results: list[dict], output_dir: Path) -> dict:
    """
    Копирует файлы в новую структуру.
    Возвращает статистику: {"copied": int, "errors": list[str]}
    """
    output_dir = Path(output_dir)
    copied = 0
    errors = []

    for doc in results:
        src = Path(doc["_file_path"])
        dst = Path(doc["_dest_path"])

        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(str(src), str(dst))
            copied += 1
        except Exception as e:
            errors.append(f"{src.name}: {str(e)}")

    return {"copied": copied, "errors": errors}


def verify_sort(source_dir: Path, output_dir: Path) -> dict:
    """
    Верифицирует что все файлы скопированы.
    Возвращает: {"source_count": int, "dest_count": int, "match": bool}
    """
    from scanner import scan_folder

    source_files = scan_folder(source_dir)
    dest_count = sum(1 for f in Path(output_dir).rglob("*") if f.is_file())

    return {
        "source_count": len(source_files),
        "dest_count": dest_count,
        "match": len(source_files) == dest_count,
    }
