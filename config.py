"""
Модуль конфигурации DocSorter.
Хранение и загрузка настроек приложения и категорий документов.
"""

import json
import os
import sys
from pathlib import Path

APP_NAME = "DocSorter"
CONFIG_FILENAME = "config.json"
CATEGORIES_FILENAME = "categories.json"


def get_app_dir() -> Path:
    """Директория приложения (рядом с exe или рядом со скриптом)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def get_config_path() -> Path:
    return get_app_dir() / CONFIG_FILENAME


def get_categories_path() -> Path:
    return get_app_dir() / CATEGORIES_FILENAME


# ── Конфиг приложения ──────────────────────────────────────────────

DEFAULT_CONFIG = {
    "api_key": "",
    "vision_model": "qwen/qwen3-vl-8b-instruct",
    "text_model": "google/gemini-2.5-flash",
    "max_concurrent": 5,
    "max_pages_per_pdf": 3,
    "sort_mode": "folders",  # "folders" | "numbering"
    "suspicious_page_threshold": 10,
    "slice_batch_size": 10,
    "fallback_enabled": True,
    "name_template": "{type} №{number} от {date} {party}",
    "table_font_size": 11,
}


def load_config() -> dict:
    path = get_config_path()
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            saved = json.load(f)
        # Дополняем значениями по умолчанию
        merged = {**DEFAULT_CONFIG, **saved}
        return merged
    return dict(DEFAULT_CONFIG)


def save_config(cfg: dict) -> None:
    path = get_config_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def is_config_valid(cfg: dict) -> bool:
    return bool(cfg.get("api_key", "").strip())


# ── Конфиг категорий ───────────────────────────────────────────────

DEFAULT_CATEGORIES = {
    "name": "Стандартный юридический",
    "categories": [
        {
            "id": "contracts",
            "name": "Договоры",
            "subcategories": [
                "Основной договор",
                "Дополнительное соглашение",
                "Приложение к договору",
            ],
        },
        {
            "id": "primary_docs",
            "name": "Первичная документация",
            "subcategories": [
                "УПД",
                "Счёт",
                "Счёт-фактура",
                "Акт выполненных работ",
                "Товарная накладная",
            ],
        },
        {
            "id": "legal",
            "name": "Правовые документы",
            "subcategories": [
                "Доверенность",
                "Протокол",
                "Решение",
                "Устав",
                "Выписка из ЕГРЮЛ",
            ],
        },
        {
            "id": "correspondence",
            "name": "Переписка",
            "subcategories": [
                "Письмо",
                "Уведомление",
                "Претензия",
                "Ответ на претензию",
            ],
        },
        {
            "id": "financial",
            "name": "Финансовые документы",
            "subcategories": [
                "Платёжное поручение",
                "Акт сверки",
                "Банковская выписка",
            ],
        },
        {
            "id": "other",
            "name": "Прочее",
            "subcategories": [],
        },
    ],
}


BASE_TEMPLATE_NAME = DEFAULT_CATEGORIES["name"]


def _make_base_template() -> dict:
    """Создаёт неизменяемую копию базового шаблона."""
    base = json.loads(json.dumps(DEFAULT_CATEGORIES))  # deep copy
    base["is_base"] = True
    return base


def save_categories(library: dict) -> None:
    """Сохраняет всю библиотеку шаблонов на диск."""
    path = get_categories_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(library, f, ensure_ascii=False, indent=2)


def load_categories() -> dict:
    """Загружает библиотеку шаблонов категорий.

    Формат:
        {"active": "имя", "templates": [{"name", "categories", "is_base"?}, ...]}

    Базовый шаблон всегда присутствует и неизменяем (восстанавливается из
    DEFAULT_CATEGORIES). Старый формат (один шаблон в корне) автоматически
    мигрируется как пользовательский шаблон рядом с базовым.
    """
    path = get_categories_path()
    data: dict = {}
    if path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError):
            data = {}

    # Миграция: старый формат — один шаблон в корне (есть "categories", нет "templates")
    if "templates" not in data:
        if "categories" in data:
            old_name = data.get("name", "Пользовательский").strip() or "Пользовательский"
            if old_name == BASE_TEMPLATE_NAME:
                # Совпадает с базовым — отбрасываем, базовый и так появится
                data = {"active": BASE_TEMPLATE_NAME, "templates": []}
            else:
                migrated = {
                    "name": old_name,
                    "categories": data.get("categories", []),
                }
                data = {"active": old_name, "templates": [migrated]}
        else:
            data = {"active": BASE_TEMPLATE_NAME, "templates": []}

    templates = data.get("templates", [])
    if not isinstance(templates, list):
        templates = []

    # Гарантируем наличие базового шаблона (всегда первый, всегда актуальный)
    base = _make_base_template()
    base_idx = next(
        (i for i, t in enumerate(templates)
         if t.get("is_base") or t.get("name") == BASE_TEMPLATE_NAME),
        -1,
    )
    if base_idx == -1:
        templates.insert(0, base)
    else:
        # Перезаписываем содержимое — базовый immutable
        templates[base_idx] = base
        if base_idx != 0:
            templates.insert(0, templates.pop(base_idx))

    data["templates"] = templates

    # Валидация активного
    names = {t["name"] for t in templates if t.get("name")}
    if data.get("active") not in names:
        data["active"] = BASE_TEMPLATE_NAME

    save_categories(data)
    return data


def find_template(library: dict, name: str) -> dict | None:
    for t in library.get("templates", []):
        if t.get("name") == name:
            return t
    return None


def get_active_template(library: dict) -> dict:
    """Возвращает активный шаблон или базовый, если активный не найден."""
    active_name = library.get("active", BASE_TEMPLATE_NAME)
    t = find_template(library, active_name)
    if t:
        return t
    base = find_template(library, BASE_TEMPLATE_NAME)
    return base or _make_base_template()


def set_active_template(library: dict, name: str) -> bool:
    if find_template(library, name):
        library["active"] = name
        return True
    return False


def add_template(library: dict, template: dict) -> tuple[bool, str]:
    """Добавляет шаблон. Возвращает (ok, error)."""
    name = (template.get("name") or "").strip()
    if not name:
        return False, "Имя шаблона не может быть пустым"
    if find_template(library, name):
        return False, "Шаблон с таким именем уже существует"
    new_t = {
        "name": name,
        "categories": template.get("categories", []),
    }
    library.setdefault("templates", []).append(new_t)
    return True, ""


def remove_template(library: dict, name: str) -> tuple[bool, str]:
    t = find_template(library, name)
    if not t:
        return False, "Шаблон не найден"
    if t.get("is_base"):
        return False, "Базовый шаблон удалить нельзя"
    if library.get("active") == name:
        return False, "Нельзя удалить активный шаблон — сначала переключитесь на другой"
    library["templates"] = [x for x in library["templates"] if x.get("name") != name]
    return True, ""


def rename_template(library: dict, old: str, new: str) -> tuple[bool, str]:
    t = find_template(library, old)
    if not t:
        return False, "Шаблон не найден"
    if t.get("is_base"):
        return False, "Базовый шаблон переименовать нельзя"
    new = (new or "").strip()
    if not new:
        return False, "Имя не может быть пустым"
    if new == old:
        return True, ""
    if find_template(library, new):
        return False, "Шаблон с таким именем уже существует"
    t["name"] = new
    if library.get("active") == old:
        library["active"] = new
    return True, ""


def update_template_content(library: dict, name: str, new_data: dict) -> tuple[bool, str]:
    """Обновляет категории в шаблоне (имя не трогаем)."""
    t = find_template(library, name)
    if not t:
        return False, "Шаблон не найден"
    if t.get("is_base"):
        return False, "Базовый шаблон редактировать нельзя"
    if "categories" not in new_data or not isinstance(new_data["categories"], list):
        return False, "В данных нет поля 'categories' (массив)"
    t["categories"] = new_data["categories"]
    return True, ""


def get_category_names(cats: dict) -> list[str]:
    """Возвращает плоский список всех категорий и подкатегорий."""
    result = []
    for cat in cats.get("categories", []):
        result.append(cat["name"])
        for sub in cat.get("subcategories", []):
            result.append(f"  {cat['name']} > {sub}")
    return result


def find_category_by_name(cats: dict, name: str) -> dict | None:
    for cat in cats.get("categories", []):
        if cat["name"] == name:
            return cat
    return None
