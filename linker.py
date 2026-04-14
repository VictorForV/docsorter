"""
Модуль детерминистической связи документов.
Строит граф связей по извлечённым метаданным,
выделяет связные компоненты (группы), назначает категории и порядок.
"""

import json
from collections import defaultdict


# ── Иерархия типов для sort_order внутри группы ────────────────────

DOC_TYPE_HIERARCHY = {
    "Договор": 1,
    "Дополнительное соглашение": 2,
    "Приложение к договору": 2,
    "Спецификация": 3,
    "Счёт": 4,
    "УПД": 5,
    "Товарная накладная": 5,
    "Счёт-фактура": 6,
    "Акт": 7,
    "Акт выполненных работ": 7,
    "Акт оказанных услуг": 7,
    "Платёжное поручение": 8,
    "Акт сверки": 9,
    "Доверенность": 10,
    "Письмо": 10,
    "Протокол": 10,
    "Решение": 10,
    "Устав": 10,
    "Выписка из ЕГРЮЛ": 10,
    "Выписка": 10,
    "Уведомление": 10,
    "Претензия": 10,
    "Ответ на претензию": 10,
    "Банковская выписка": 10,
    "Прочее": 99,
}

# ── Маппинг doc_type → (категория, подкатегория) ──────────────────

DOC_TYPE_TO_CATEGORY = {
    "Договор": ("Договоры", "Основной договор"),
    "Дополнительное соглашение": ("Договоры", "Дополнительное соглашение"),
    "Приложение к договору": ("Договоры", "Приложение к договору"),
    "Спецификация": ("Договоры", "Приложение к договору"),
    "УПД": ("Первичная документация", "УПД"),
    "Счёт": ("Первичная документация", "Счёт"),
    "Счёт-фактура": ("Первичная документация", "Счёт-фактура"),
    "Акт": ("Первичная документация", "Акт выполненных работ"),
    "Акт выполненных работ": ("Первичная документация", "Акт выполненных работ"),
    "Акт оказанных услуг": ("Первичная документация", "Акт выполненных работ"),
    "Товарная накладная": ("Первичная документация", "Товарная накладная"),
    "Доверенность": ("Правовые документы", "Доверенность"),
    "Протокол": ("Правовые документы", "Протокол"),
    "Решение": ("Правовые документы", "Решение"),
    "Устав": ("Правовые документы", "Устав"),
    "Выписка из ЕГРЮЛ": ("Правовые документы", "Выписка из ЕГРЮЛ"),
    "Письмо": ("Переписка", "Письмо"),
    "Уведомление": ("Переписка", "Уведомление"),
    "Претензия": ("Переписка", "Претензия"),
    "Ответ на претензию": ("Переписка", "Ответ на претензию"),
    "Платёжное поручение": ("Финансовые документы", "Платёжное поручение"),
    "Акт сверки": ("Финансовые документы", "Акт сверки"),
    "Банковская выписка": ("Финансовые документы", "Банковская выписка"),
    "Выписка": ("Финансовые документы", "Банковская выписка"),
}

# Типы, совместимые для implicit linking (ключ — тип якоря, значение — типы подчинённых)
_COMPATIBLE_TYPES = {
    "Договор": {
        "Дополнительное соглашение", "Приложение к договору", "Спецификация",
        "Счёт", "УПД", "Товарная накладная", "Счёт-фактура",
        "Акт", "Акт выполненных работ", "Акт оказанных услуг",
        "Платёжное поручение",
    },
    "Счёт": {"УПД", "Счёт-фактура", "Товарная накладная", "Платёжное поручение"},
    "УПД": {"Счёт-фактура", "Акт", "Акт выполненных работ"},
}


# ── Утилиты ────────────────────────────────────────────────────────

def _normalize_name(name: str) -> str:
    """Нормализация названия организации для нечёткого сравнения."""
    s = name.lower().strip()
    # Убираем кавычки
    for ch in ('"', "'", "«", "»", "\u201e", "\u201c"):
        s = s.replace(ch, "")
    # Убираем распространённые префиксы
    for prefix in ("ооо ", "ао ", "зао ", "пао ", "ип ", "оао ", "нао "):
        if s.startswith(prefix):
            s = s[len(prefix):]
    return s.strip()


def _parse_party(raw: str) -> dict:
    """Парсит JSON-строку party-поля → {"name": ..., "role": ...}."""
    if not raw:
        return {}
    try:
        data = json.loads(raw)
        if isinstance(data, dict) and data.get("name"):
            return data
    except (json.JSONDecodeError, TypeError):
        pass
    return {}


def _normalize_number(num: str) -> str:
    """Нормализация номера документа для сравнения."""
    s = num.strip()
    # Убираем общие префиксы
    for prefix in ("№", "N", "№ ", "N "):
        if s.upper().startswith(prefix):
            s = s[len(prefix):]
    return s.strip()


# ── Индексы ────────────────────────────────────────────────────────

def build_indexes(results: list[dict]) -> dict:
    """Строит индексы для быстрого поиска связанных документов."""
    indexes = {
        "by_number_date": defaultdict(list),
        "by_party": defaultdict(list),
        "by_doc_type": defaultdict(list),
    }
    for i, doc in enumerate(results):
        num = _normalize_number(doc.get("number", ""))
        date = doc.get("date", "").strip()
        if num and date:
            indexes["by_number_date"][(num, date)].append(i)

        for pfield in ("party_1", "party_2"):
            party = _parse_party(doc.get(pfield, ""))
            if party.get("name"):
                norm = _normalize_name(party["name"])
                if norm:
                    indexes["by_party"][norm].append(i)

        dt = doc.get("doc_type", "").strip()
        if dt:
            indexes["by_doc_type"][dt].append(i)

    return indexes


# ── Поиск связей ───────────────────────────────────────────────────

def find_explicit_links(results: list[dict], indexes: dict) -> list[tuple[int, int]]:
    """Находит явные ссылки: документ ссылается на другой по номеру+дате."""
    links = []
    seen = set()
    for i, doc in enumerate(results):
        ref_num = _normalize_number(doc.get("reference_number", ""))
        ref_date = doc.get("reference_date", "").strip()
        if not ref_num:
            continue

        # Точное совпадение номер + дата
        key = (ref_num, ref_date)
        matched = indexes["by_number_date"].get(key, [])

        # Если дата пуста — ищем только по номеру
        if not matched and not ref_date:
            for (n, _d), idxs in indexes["by_number_date"].items():
                if n == ref_num:
                    matched = idxs
                    break

        for j in matched:
            if j != i:
                pair = (min(i, j), max(i, j))
                if pair not in seen:
                    seen.add(pair)
                    links.append((i, j))

    return links


def _are_compatible_types(type_a: str, type_b: str) -> bool:
    """Проверяет, могут ли два типа документов быть связаны."""
    compatible = _COMPATIBLE_TYPES.get(type_a, set())
    if type_b in compatible:
        return True
    compatible = _COMPATIBLE_TYPES.get(type_b, set())
    if type_a in compatible:
        return True
    return False


def find_implicit_links(results: list[dict], indexes: dict) -> list[tuple[int, int]]:
    """Находит неявные связи по общим контрагентам и совместимым типам."""
    links = []
    seen = set()
    for _norm_name, idxs in indexes["by_party"].items():
        if len(idxs) < 2:
            continue
        for a in range(len(idxs)):
            for b in range(a + 1, len(idxs)):
                i, j = idxs[a], idxs[b]
                pair = (min(i, j), max(i, j))
                if pair in seen:
                    continue
                if _are_compatible_types(
                    results[i].get("doc_type", ""),
                    results[j].get("doc_type", ""),
                ):
                    seen.add(pair)
                    links.append((i, j))
    return links


# ── Union-Find для связных компонент ──────────────────────────────

def _connected_components(n: int, links: list[tuple[int, int]]) -> list[list[int]]:
    """Union-Find → список связных компонент."""
    parent = list(range(n))

    def find(x: int) -> int:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a: int, b: int):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    for i, j in links:
        union(i, j)

    groups = defaultdict(list)
    for i in range(n):
        groups[find(i)].append(i)

    return list(groups.values())


# ── Назначение метаданных групп ───────────────────────────────────

def _make_group_name(anchor: dict) -> str:
    """Формирует имя группы из якорного документа."""
    parts = []
    dt = anchor.get("doc_type", "").strip()
    if dt and dt != "Прочее":
        parts.append(dt)

    num = anchor.get("number", "").strip()
    if num:
        parts.append(f"№{num}")

    date = anchor.get("date", "").strip()
    if date:
        parts.append(f"от {date}")

    # Контрагент — берём из party_2 (обычно заказчик/покупатель)
    p2 = _parse_party(anchor.get("party_2", ""))
    p1 = _parse_party(anchor.get("party_1", ""))
    party_name = p2.get("name", "") or p1.get("name", "")
    if party_name:
        parts.append(party_name)

    return " ".join(parts) if parts else anchor.get("_file_name", "документ")


def _make_new_name(doc: dict) -> str:
    """Формирует новое имя файла из полей документа."""
    parts = []
    dt = doc.get("doc_type", "").strip()
    if dt and dt != "Прочее":
        parts.append(dt)

    num = doc.get("number", "").strip()
    if num:
        parts.append(f"№{num}")

    date = doc.get("date", "").strip()
    if date:
        parts.append(f"от {date}")

    p1 = _parse_party(doc.get("party_1", ""))
    p2 = _parse_party(doc.get("party_2", ""))
    party_name = p2.get("name", "") or p1.get("name", "")
    if party_name:
        parts.append(party_name)

    return " ".join(parts) if parts else doc.get("_file_name", "документ").rsplit(".", 1)[0]


def _find_category_in_template(
    doc_type: str, categories_template: dict,
) -> tuple[str, str]:
    """Ищет категорию в шаблоне по маппингу doc_type → (category, subcategory)."""
    mapped = DOC_TYPE_TO_CATEGORY.get(doc_type)
    if not mapped:
        return ("Прочее", "")

    target_cat, target_sub = mapped

    # Проверяем, что категория есть в активном шаблоне
    for cat in categories_template.get("categories", []):
        if cat.get("name") == target_cat:
            # Проверяем подкатегорию
            subs = cat.get("subcategories", [])
            if target_sub in subs:
                return (target_cat, target_sub)
            # Категория есть, подкатегории нет — возвращаем категорию без подкатегории
            return (target_cat, "")

    # Категории нет в шаблоне — fallback на Прочее
    return ("Прочее", "")


def assign_group_metadata(
    results: list[dict],
    groups: list[list[int]],
    categories_template: dict,
) -> list[int]:
    """
    Назначает _category, _subcategory, _group, _sort_order, _new_name.
    Возвращает индексы документов-сирот (группы из 1 документа без явных связей).
    """
    orphans = []

    for group_indices in groups:
        if len(group_indices) == 1:
            i = group_indices[0]
            doc = results[i]
            # Одиночный документ — назначаем категорию и имя, но помечаем как сироту
            cat, sub = _find_category_in_template(
                doc.get("doc_type", ""), categories_template,
            )
            doc["_category"] = cat
            doc["_subcategory"] = sub
            doc["_group"] = _make_group_name(doc)
            doc["_sort_order"] = DOC_TYPE_HIERARCHY.get(
                doc.get("doc_type", ""), 99,
            )
            doc["_new_name"] = _make_new_name(doc)
            orphans.append(i)
            continue

        # Группа из нескольких документов — находим якорь (минимальный hierarchy)
        anchor_idx = min(
            group_indices,
            key=lambda idx: DOC_TYPE_HIERARCHY.get(
                results[idx].get("doc_type", ""), 99,
            ),
        )
        anchor = results[anchor_idx]
        group_name = _make_group_name(anchor)

        for idx in group_indices:
            doc = results[idx]
            doc["_group"] = group_name
            doc["_sort_order"] = DOC_TYPE_HIERARCHY.get(
                doc.get("doc_type", ""), 99,
            )
            cat, sub = _find_category_in_template(
                doc.get("doc_type", ""), categories_template,
            )
            doc["_category"] = cat
            doc["_subcategory"] = sub
            doc["_new_name"] = _make_new_name(doc)

    return orphans


# ── Основная функция ──────────────────────────────────────────────

def link_documents(
    results: list[dict],
    categories_template: dict,
) -> tuple[list[dict], list[int]]:
    """
    Детерминистическая группировка документов.

    Args:
        results: список документов с полями анализа
        categories_template: активный шаблон категорий

    Returns:
        (results с назначенными _category/_group/_sort_order/_new_name,
         список индексов документов-сирот для LLM-обработки)
    """
    if not results:
        return results, []

    indexes = build_indexes(results)
    explicit = find_explicit_links(results, indexes)
    implicit = find_implicit_links(results, indexes)
    all_links = explicit + implicit

    groups = _connected_components(len(results), all_links)
    orphans = assign_group_metadata(results, groups, categories_template)

    return results, orphans
