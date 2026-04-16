"""
Модуль детерминистической связи документов.
Строит граф связей по извлечённым метаданным,
выделяет связные компоненты (группы), назначает категории и порядок.
"""

import json
from collections import Counter, defaultdict
from difflib import SequenceMatcher

from doctypes import get_hierarchy, get_type_to_category, COMPATIBLE_TYPES, resolve_base_type


# ── Иерархия типов для sort_order внутри группы ────────────────────
# Генерируется из doctypes.py

DOC_TYPE_HIERARCHY = get_hierarchy()

# ── Маппинг doc_type → (категория, подкатегория) ──────────────────
# Генерируется из doctypes.py

DOC_TYPE_TO_CATEGORY = get_type_to_category()

# Типы, совместимые для implicit linking — из doctypes.py
_COMPATIBLE_TYPES = COMPATIBLE_TYPES


# ── Утилиты ────────────────────────────────────────────────────────

# Таблица гомоглифов: кириллица → латиница (визуально одинаковые символы)
_HOMOGLYPH_MAP = str.maketrans({
    "\u0410": "A",  # А → A
    "\u0412": "B",  # В → B
    "\u0415": "E",  # Е → E
    "\u041a": "K",  # К → K
    "\u041c": "M",  # М → M
    "\u041d": "H",  # Н → H
    "\u041e": "O",  # О → O
    "\u0420": "P",  # Р → P
    "\u0421": "C",  # С → C
    "\u0422": "T",  # Т → T
    "\u0425": "X",  # Х → X
    "\u0430": "a",  # а → a
    "\u0435": "e",  # е → e
    "\u043e": "o",  # о → o
    "\u0440": "p",  # р → p
    "\u0441": "c",  # с → c
    "\u0443": "y",  # у → y
    "\u0445": "x",  # х → x
})


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
    # Гомоглифы: кириллица → латиница
    s = s.translate(_HOMOGLYPH_MAP)
    # Убираем общие префиксы
    for prefix in ("№", "N", "№ ", "N "):
        if s.upper().startswith(prefix):
            s = s[len(prefix):]
    return s.strip()


# ── Нормализация контрагентов ───────────────────────────────────────

# Пороги нечёткого сравнения:
#   ratio >= HIGH → всегда объединяем
#   LOW <= ratio < HIGH → объединяем только при общем суффиксе >= MIN_SUFFIX
_FUZZY_HIGH = 0.80
_FUZZY_LOW = 0.60
_MIN_SUFFIX = 4


def _common_suffix_len(a: str, b: str) -> int:
    """Длина общего суффикса двух строк."""
    min_len = min(len(a), len(b))
    for i in range(1, min_len + 1):
        if a[-i] != b[-i]:
            return i - 1
    return min_len


def normalize_party_names(results: list[dict]) -> list[dict]:
    """Кластеризует названия контрагентов по нечёткому совпадению
    и заменяет все варианты на наиболее частый (каноничный).

    Вызывается после анализа всех документов, до построения индексов.
    """
    # Шаг 1: собираем все party-записи
    parties: list[tuple[int, str, str, str]] = []
    # (index в results, поле "party_1"/"party_2", оригинальное имя, нормализованное)

    for i, r in enumerate(results):
        for pfield in ("party_1", "party_2"):
            party = _parse_party(r.get(pfield, ""))
            name = party.get("name", "")
            if name:
                norm = _normalize_name(name)
                if norm:
                    parties.append((i, pfield, name, norm))

    if not parties:
        return results

    # Шаг 2: уникальные нормализованные имена → Union-Find
    unique_norms = list({p[3] for p in parties})

    parent = list(range(len(unique_norms)))

    def _find(x: int) -> int:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def _union(a: int, b: int) -> None:
        ra, rb = _find(a), _find(b)
        if ra != rb:
            parent[ra] = rb

    # Попарное нечёткое сравнение с двухуровневым порогом
    for a in range(len(unique_norms)):
        for b in range(a + 1, len(unique_norms)):
            sa, sb = unique_norms[a], unique_norms[b]
            ratio = SequenceMatcher(None, sa, sb).ratio()
            if ratio >= _FUZZY_HIGH:
                _union(a, b)
            elif ratio >= _FUZZY_LOW:
                # Borderline — объединяем только при значимом общем суффиксе
                if _common_suffix_len(sa, sb) >= _MIN_SUFFIX:
                    _union(a, b)

    # Шаг 3: строим кластеры
    clusters: dict[int, list[str]] = defaultdict(list)
    for k in range(len(unique_norms)):
        clusters[_find(k)].append(unique_norms[k])

    # Шаг 4: для каждого кластера >1 имени — находим каноничное и заменяем
    for _root, norm_names in clusters.items():
        if len(norm_names) <= 1:
            continue

        norm_set = set(norm_names)

        # Собираем все оригинальные написания в кластере
        originals = []
        for _, _, orig, norm in parties:
            if norm in norm_set:
                originals.append(orig)

        # Наиболее частый вариант = каноничный
        canonical = Counter(originals).most_common(1)[0][0]

        # Заменяем все варианты на каноничный
        for idx, pfield, orig, norm in parties:
            if norm in norm_set and orig != canonical:
                party = _parse_party(results[idx].get(pfield, ""))
                party["name"] = canonical
                results[idx][pfield] = json.dumps(party, ensure_ascii=False)

    # Обновляем legacy-поле counterparty
    for r in results:
        p1 = _parse_party(r.get("party_1", ""))
        p2 = _parse_party(r.get("party_2", ""))
        name = p1.get("name", "") or p2.get("name", "")
        if name:
            r["counterparty"] = name

    return results


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
    # Разрешаем подвиды договоров (напр. «Договор поставки» → базовый «Договор»)
    base_a = resolve_base_type(type_a)
    base_b = resolve_base_type(type_b)

    compatible = _COMPATIBLE_TYPES.get(base_a, set())
    if base_b in compatible or type_b in compatible:
        return True
    compatible = _COMPATIBLE_TYPES.get(base_b, set())
    if base_a in compatible or type_a in compatible:
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

    num = _normalize_number(anchor.get("number", ""))
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


def _make_new_name(doc: dict, template: str = "") -> str:
    """Формирует новое имя файла из полей документа по шаблону."""
    p1 = _parse_party(doc.get("party_1", ""))
    p2 = _parse_party(doc.get("party_2", ""))
    party_name = p2.get("name", "") or p1.get("name", "")

    replacements = {
        "{type}": doc.get("doc_type", "").strip(),
        "{number}": _normalize_number(doc.get("number", "")),
        "{date}": doc.get("date", "").strip(),
        "{party}": party_name,
        "{party_1}": p1.get("name", ""),
        "{party_2}": p2.get("name", ""),
        "{title}": doc.get("title", "").strip(),
        "{amount}": doc.get("amount", "").strip(),
    }

    if template:
        result = template
        for key, val in replacements.items():
            result = result.replace(key, val)
        # Убираем лишние пробелы вокруг пустых подстановок
        import re
        result = re.sub(r"\s{2,}", " ", result).strip()
        # Убираем висячие "№" и "от" если номер/дата пустые
        result = result.replace("№ ", "").replace(" от ", " ")
        return result if result else doc.get("_file_name", "документ").rsplit(".", 1)[0]

    # Fallback: дефолтный шаблон
    return _make_new_name(doc, "{type} №{number} от {date} {party}")


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
    name_template: str = "",
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
            doc["_new_name"] = _make_new_name(doc, name_template)
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
            doc["_new_name"] = _make_new_name(doc, name_template)

    return orphans


# ── Основная функция ──────────────────────────────────────────────

def link_documents(
    results: list[dict],
    categories_template: dict,
    name_template: str = "",
) -> tuple[list[dict], list[int]]:
    """
    Детерминистическая группировка документов.

    Args:
        results: список документов с полями анализа
        categories_template: активный шаблон категорий
        name_template: шаблон имени файла (напр. '{type} №{number} от {date} {party}')

    Returns:
        (results с назначенными _category/_group/_sort_order/_new_name,
         список индексов документов-сирот для LLM-обработки)
    """
    if not results:
        return results, []

    # Нормализуем названия контрагентов (нечёткое сравнение)
    results = normalize_party_names(results)

    indexes = build_indexes(results)
    explicit = find_explicit_links(results, indexes)
    implicit = find_implicit_links(results, indexes)
    all_links = explicit + implicit

    groups = _connected_components(len(results), all_links)
    orphans = assign_group_metadata(results, groups, categories_template, name_template)

    return results, orphans
