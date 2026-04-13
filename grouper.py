"""
Модуль группировки документов.
Использует текстовую LLM для распределения документов по категориям
и установления связей между документами.
"""

import asyncio
import json

import httpx

from analyzer import OPENROUTER_URL

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
        line = (
            f"[{i}] Файл: {doc.get('_file_name', '?')} | "
            f"Тип: {doc.get('doc_type', '?')} | "
            f"Название: {doc.get('title', '?')} | "
            f"Номер: {doc.get('number', '')} | "
            f"Дата: {doc.get('date', '')} | "
            f"Контрагент: {doc.get('counterparty', '')} | "
            f"Ссылка: {doc.get('reference', '')} | "
            f"Содержание: {doc.get('summary', '')}"
        )
        lines.append(line)
    return "\n".join(lines)


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
) -> list[dict]:
    """
    Группирует документы по категориям с помощью LLM.
    Возвращает список с полями группировки для каждого документа.
    """
    prompt = GROUPING_PROMPT_TEMPLATE.format(
        categories=_format_categories(categories),
        documents=_format_documents(results),
    )
    raw = await _call_llm(api_key, text_model, prompt)
    grouping = json.loads(raw)

    # Мержим результаты
    for item in grouping:
        idx = item.get("index", -1)
        if 0 <= idx < len(results):
            results[idx]["_category"] = item.get("category", "Прочее")
            results[idx]["_subcategory"] = item.get("subcategory", "")
            results[idx]["_group"] = item.get("group", "")
            results[idx]["_sort_order"] = item.get("sort_order", 99)
            results[idx]["_new_name"] = item.get("new_name", results[idx].get("_file_name", "документ"))

    # Для документов, которые LLM пропустила
    for doc in results:
        if "_category" not in doc:
            doc["_category"] = "Прочее"
            doc["_subcategory"] = ""
            doc["_group"] = ""
            doc["_sort_order"] = 99
            doc["_new_name"] = doc.get("_file_name", "документ").rsplit(".", 1)[0]

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
