"""
Модуль работы с файлом проекта DocSorter.
Проект = сохранённое состояние таблицы (JSON), которое можно открыть заново,
пополнять файлами, редактировать, нарезать и т.д.
"""

import hashlib
import json
import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path

PROJECT_FILENAME = "docsorter-project.json"
PROJECT_VERSION = 1


def file_hash(path: Path, chunk_size: int = 64 * 1024) -> str:
    """
    Быстрый идентификатор файла: md5 первых chunk_size байт + размер.
    Достаточно для дедупа в пределах проекта (вероятность коллизии ничтожна).
    """
    path = Path(path)
    if not path.exists() or not path.is_file():
        return ""
    size = path.stat().st_size
    h = hashlib.md5()
    with open(path, "rb") as f:
        h.update(f.read(chunk_size))
    return f"{h.hexdigest()}_{size}"


def get_default_project_path(source_dir: Path) -> Path:
    return Path(source_dir) / PROJECT_FILENAME


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def build_project_state(
    source_dir: Path,
    output_dir: Path | None,
    sort_mode: str,
    suspicious_page_threshold: int,
    categories_order: list[str],
    documents: list[dict],
    created: str | None = None,
) -> dict:
    return {
        "version": PROJECT_VERSION,
        "created": created or now_iso(),
        "updated": now_iso(),
        "source_dir": str(source_dir) if source_dir else "",
        "output_dir": str(output_dir) if output_dir else None,
        "sort_mode": sort_mode,
        "suspicious_page_threshold": suspicious_page_threshold,
        "categories_order": list(categories_order),
        "documents": documents,
    }


def save_project(state: dict, path: Path) -> None:
    """
    Атомарное сохранение: пишем в .tmp и переименовываем.
    Обновляем поле 'updated'.
    """
    path = Path(path)
    state["updated"] = now_iso()
    path.parent.mkdir(parents=True, exist_ok=True)

    fd, tmp_name = tempfile.mkstemp(
        prefix=".docsorter-", suffix=".tmp", dir=str(path.parent),
    )
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
        # На Windows os.replace атомарен
        os.replace(tmp_name, str(path))
    except Exception:
        try:
            os.unlink(tmp_name)
        except OSError:
            pass
        raise


def load_project(path: Path) -> dict:
    path = Path(path)
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return migrate_if_needed(data)


def migrate_if_needed(data: dict) -> dict:
    """Миграция между версиями формата проекта."""
    version = data.get("version", 0)

    if version == PROJECT_VERSION:
        return data

    # Будущие миграции тут
    # if version == 0: ... version = 1

    data["version"] = PROJECT_VERSION
    return data


def default_document_fields() -> dict:
    """Поля документа, которые должны присутствовать после загрузки."""
    return {
        "_file_path": "",
        "_rel_path": "",
        "_file_name": "",
        "_ext": "",
        "_file_hash": "",
        "_page_count": 0,
        "doc_type": "",
        "title": "",
        "number": "",
        "date": "",
        "counterparty": "",
        "reference": "",
        "summary": "",
        "_category": "Прочее",
        "_group": "",
        "_sort_order": 99,
        "_new_name": "",
        "_comment": "",
        "_suspicious": False,
        "_suspicious_reason": "",
        "_sliced_from": None,
        "_slice_parts": None,
    }


def normalize_document(doc: dict) -> dict:
    """Дополняет документ дефолтными полями (для совместимости со старыми проектами)."""
    defaults = default_document_fields()
    for key, val in defaults.items():
        if key not in doc:
            doc[key] = val
    return doc


def find_by_hash(documents: list[dict], hash_: str) -> dict | None:
    if not hash_:
        return None
    for doc in documents:
        if doc.get("_file_hash") == hash_:
            return doc
    return None
