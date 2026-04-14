"""
Модуль группировки документов.
Использует текстовую LLM для распределения документов по категориям
и установления связей между документами.
"""

import asyncio
import json
import re

import httpx

from analyzer import OPENROUTER_URL

_PAGES_SUFFIX_RE = re.compile(r"\s*\(\s*\d+\s*стр\.?\s*\)\s*$")


def with_page_count_suffix(name: str, page_count: int) -> str:
    """Возвращает имя с суффиксом ' (N стр.)' в конце.

    - Если в имени уже есть похожий суффикс, он заменяется на актуальный
      (на случай пересчёта/нарезки страниц).
    - Если page_count <= 0, существующий суффикс просто срезается.
    Идемпотентно: повторный вызов не плодит дубли.
    """
    if not name:
        return name
    base = _PAGES_SUFFIX_RE.sub("", name).rstrip()
    if not page_count or page_count <= 0:
        return base
    return f"{base} ({page_count} стр.)"

GROUPING_PROMPT_TEMPLATE = """Ты — помощник юриста. Тебе дан список документов и набор категорий.
Твоя задача — распределить документы по категориям и установить связи между ними.

КАТЕГОРИИ:
{categories}

ДОКУМЕНТЫ:
{documents}

Верни СТРОГО JSON (без markdown, без ```json) — массив объектов, по одному на каждый документ:
[
  {{
    "index": 0,
    "category": "название категории из списка выше",
    "subcategory": "подкатегория (если применимо) или пустая строка",
    "group": "идентификатор группы связанных документов (напр. 'Договор №123 ООО Ромашка')",
    "sort_order": 1,
    "new_name": "предложенное имя файла без расширения (кратко и информативно)"
  }},
  ...
]

Правила:
- index — порядковый номер документа из списка (начиная с 0)
- Связанные документы (договор + доп.соглашения + УПД к нему) должны иметь одинаковый group
- sort_order — порядок внутри группы (1 = основной документ, 2, 3... = связанные)
- new_name — краткое, но информативное имя (напр. "Договор №123 от 15.03.2024 ООО Ромашка")
- Если документ не подходит ни под одну категорию, используй "Прочее"
- Верни ТОЛЬКО JSON-массив, без пояснений
"""

REGROUP_PROMPT_TEMPLATE = """Ты — помощник юриста. Пользователь хочет перегруппировать документы.

Текущее распределение документов:
{current_state}

Запрос пользователя:
{user_request}

Доступные категории:
{categories}

Перегруппируй документы согласно запросу. Верни СТРОГО JSON (без markdown) — массив объектов:
[
  {{
    "index": 0,
    "category": "категория",
    "subcategory": "подкатегория или пустая строка",
    "group": "группа связанных документов",
    "sort_order": 1,
    "new_name": "имя файла без расширения"
  }},
  ...
]
"""

GENERATE_CATEGORIES_PROMPT = """Ты — помощник юриста. Пользователь хочет создать свой набор категорий для сортировки документов.

Запрос пользователя:
{user_request}

Сгенерируй JSON-конфиг категорий. Верни СТРОГО JSON (без markdown, без ```json):
{{
  "name": "Название набора категорий",
  "categories": [
    {{
      "id": "уникальный_id",
      "name": "Название категории",
      "subcategories": ["подкатегория1", "подкатегория2"]
    }},
    ...
  ]
}}

Обязательно добавь категорию "Прочее" с пустым списком подкатегорий в конце.
"""


def _format_categories(cats: dict) -> str:
    lines = []
    for cat in cats.get("categories", []):
        subs = ", ".join(cat.get("subcategories", []))
        line = f"- {cat['name']}"
        if subs:
            line += f" ({subs})"
        lines.append(line)
    return "\n".join(lines)


def _format_documents(results: list[dict]) -> str:
    lines = []
    for i, doc in enumerate(results):
        # Парсим party-поля для отображения
        p1_name = _parse_party_display(doc.get("party_1", ""))
        p2_name = _parse_party_display(doc.get("party_2", ""))
        counterparty = doc.get("counterparty", "")
        side1 = p1_name or counterparty
        side2 = p2_name

        line = (
            f"[{i}] Файл: {doc.get('_file_name', '?')} | "
            f"Тип: {doc.get('doc_type', '?')} | "
            f"Название: {doc.get('title', '?')} | "
            f"Номер: {doc.get('number', '')} | "
            f"Дата: {doc.get('date', '')} | "
            f"Сторона 1: {side1} | "
            f"Сторона 2: {side2} | "
            f"Ссылка на: №{doc.get('reference_number', '')} от {doc.get('reference_date', '')} | "
            f"Сумма: {doc.get('amount', '')} | "
            f"Предмет: {doc.get('goods_summary', '')} | "
            f"Содержание: {doc.get('summary', '')}"
        )
        lines.append(line)
    return "\n".join(lines)


def _parse_party_display(raw: str) -> str:
    """Извлекает имя из JSON-строки party-поля."""
    if not raw:
        return ""
    try:
        data = json.loads(raw)
        return data.get("name", "")
    except (json.JSONDecodeError, TypeError):
        return raw


async def _call_llm(api_key: str, model: str, prompt: str) -> str:
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://docsorter.app",
        "X-Title": "DocSorter",
    }
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.1,
        "max_tokens": 4000,
    }
    async with httpx.AsyncClient() as client:
        resp = await client.post(
            OPENROUTER_URL, json=payload, headers=headers, timeout=120.0,
        )
        resp.raise_for_status()
        data = resp.json()
        raw = data["choices"][0]["message"]["content"].strip()
        # Убираем markdown
        if raw.startswith("```"):
            lines = raw.split("\n")
            lines = [l for l in lines if not l.strip().startswith("```")]
            raw = "\n".join(lines)
        return raw


async def group_documents(
    results: list[dict],
    categories: dict,
    api_key: str,
    text_model: str,
    name_template: str = "",
) -> list[dict]:
    """
    Фаза 1: детерминистическая группировка через linker.
    Фаза 2: LLM только для документов-сирот.
    """
    from linker import link_documents

    results, orphan_indices = link_documents(results, categories, name_template)

    if orphan_indices:
        orphan_docs = [results[i] for i in orphan_indices]
        prompt = GROUPING_PROMPT_TEMPLATE.format(
            categories=_format_categories(categories),
            documents=_format_documents(orphan_docs),
        )
        try:
            raw = await _call_llm(api_key, text_model, prompt)
            grouping = json.loads(raw)

            for item in grouping:
                idx = item.get("index", -1)
                if 0 <= idx < len(orphan_docs):
                    real_idx = orphan_indices[idx]
                    results[real_idx]["_category"] = item.get("category", "Прочее")
                    results[real_idx]["_subcategory"] = item.get("subcategory", "")
                    results[real_idx]["_group"] = item.get("group", results[real_idx].get("_group", ""))
                    results[real_idx]["_sort_order"] = item.get("sort_order", results[real_idx].get("_sort_order", 99))
                    results[real_idx]["_new_name"] = item.get(
                        "new_name", results[real_idx].get("_new_name", results[real_idx].get("_file_name", "документ"))
                    )
                    results[real_idx]["_new_name"] = with_page_count_suffix(
                        results[real_idx]["_new_name"],
                        results[real_idx].get("_page_count", 0),
                    )
        except Exception:
            # LLM не справилась с сиротами — оставляем детерминистические значения
            pass

    # Safety net: документы без _category (не должно быть, но на всякий случай)
    for doc in results:
        if "_category" not in doc:
            doc["_category"] = "Прочее"
            doc["_subcategory"] = ""
            doc["_group"] = ""
            doc["_sort_order"] = 99
            doc["_new_name"] = doc.get("_file_name", "документ").rsplit(".", 1)[0]
            doc["_new_name"] = with_page_count_suffix(doc["_new_name"], doc.get("_page_count", 0))

    return results


async def regroup_documents(
    results: list[dict],
    categories: dict,
    user_request: str,
    api_key: str,
    text_model: str,
) -> list[dict]:
    """Перегруппирует документы по запросу пользователя."""
    current_state = _format_documents_with_groups(results)
    prompt = REGROUP_PROMPT_TEMPLATE.format(
        current_state=current_state,
        user_request=user_request,
        categories=_format_categories(categories),
    )
    raw = await _call_llm(api_key, text_model, prompt)
    grouping = json.loads(raw)

    for item in grouping:
        idx = item.get("index", -1)
        if 0 <= idx < len(results):
            results[idx]["_category"] = item.get("category", results[idx].get("_category", "Прочее"))
            results[idx]["_subcategory"] = item.get("subcategory", "")
            results[idx]["_group"] = item.get("group", "")
            results[idx]["_sort_order"] = item.get("sort_order", 99)
            results[idx]["_new_name"] = item.get("new_name", results[idx].get("_new_name", ""))
            results[idx]["_new_name"] = with_page_count_suffix(
                results[idx]["_new_name"], results[idx].get("_page_count", 0),
            )

    return results


async def generate_categories(
    user_request: str,
    api_key: str,
    text_model: str,
) -> dict:
    """Генерирует новый набор категорий по запросу пользователя."""
    prompt = GENERATE_CATEGORIES_PROMPT.format(user_request=user_request)
    raw = await _call_llm(api_key, text_model, prompt)
    return json.loads(raw)


def _format_documents_with_groups(results: list[dict]) -> str:
    lines = []
    for i, doc in enumerate(results):
        line = (
            f"[{i}] Файл: {doc.get('_file_name', '?')} | "
            f"Тип: {doc.get('doc_type', '?')} | "
            f"Категория: {doc.get('_category', '?')} | "
            f"Группа: {doc.get('_group', '')} | "
            f"Порядок: {doc.get('_sort_order', '?')} | "
            f"Новое имя: {doc.get('_new_name', '?')}"
        )
        lines.append(line)
    return "\n".join(lines)
