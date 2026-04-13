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
    "vision_model": "google/gemini-2.0-flash-001",
    "text_model": "google/gemini-2.0-flash-001",
    "max_concurrent": 5,
    "max_pages_per_pdf": 3,
    "sort_mode": "folders",  # "folders" | "numbering"
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


def load_categories() -> dict:
    path = get_categories_path()
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    # Создаём файл по умолчанию
    save_categories(DEFAULT_CATEGORIES)
    return dict(DEFAULT_CATEGORIES)


def save_categories(cats: dict) -> None:
    path = get_categories_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cats, f, ensure_ascii=False, indent=2)


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
