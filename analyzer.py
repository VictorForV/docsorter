"""
Модуль анализа документов через OpenRouter API.
Vision-модель для сканов/изображений, текстовая модель для текстовых документов.
Батчинг через asyncio.
"""

import asyncio
import base64
import json
import io
from pathlib import Path

import httpx
import fitz  # PyMuPDF
from docx import Document as DocxDocument
from openpyxl import load_workbook

from scanner import get_file_type

OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"

ANALYSIS_PROMPT = """Ты — помощник юриста, анализирующий документы. Проанализируй предоставленный документ и верни СТРОГО JSON (без markdown, без ```json):

{
  "doc_type": "тип документа (Договор / Дополнительное соглашение / УПД / Счёт / Счёт-фактура / Акт / Письмо / Доверенность / Платёжное поручение / Акт сверки / Выписка / Прочее)",
  "title": "полное название документа",
  "number": "номер документа или пустая строка",
  "date": "дата документа в формате ДД.ММ.ГГГГ или пустая строка",
  "counterparty": "контрагент (название организации/ИП) или пустая строка",
  "reference": "ссылка на связанный документ (напр. 'к Договору №123 от 01.01.2024') или пустая строка",
  "summary": "краткое содержание (1-2 предложения)"
}

Важно:
- Если какое-то поле не удаётся определить, оставь пустую строку
- Тип документа выбирай из списка выше
- Дату приводи к формату ДД.ММ.ГГГГ
- Верни ТОЛЬКО JSON, без пояснений
"""


def _pdf_to_images(pdf_path: Path, max_pages: int = 3) -> list[bytes]:
    """Рендерит страницы PDF в PNG-изображения."""
    doc = fitz.open(str(pdf_path))
    images = []
    for i, page in enumerate(doc):
        if i >= max_pages:
            break
        # Рендерим с разрешением 200 DPI (достаточно для OCR)
        pix = page.get_pixmap(dpi=200)
        images.append(pix.tobytes("png"))
    doc.close()
    return images


def _image_to_base64(image_bytes: bytes) -> str:
    return base64.b64encode(image_bytes).decode("utf-8")


def _read_image_file(path: Path) -> bytes:
    with open(path, "rb") as f:
        return f.read()


def _extract_docx_text(path: Path) -> str:
    doc = DocxDocument(str(path))
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    return "\n".join(paragraphs[:100])  # Первые 100 абзацев


def _extract_xlsx_text(path: Path) -> str:
    wb = load_workbook(str(path), read_only=True, data_only=True)
    lines = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        lines.append(f"=== Лист: {sheet_name} ===")
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i >= 50:  # Первые 50 строк
                break
            cells = [str(c) if c is not None else "" for c in row]
            lines.append(" | ".join(cells))
    wb.close()
    return "\n".join(lines)


def _extract_text(path: Path) -> str:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read(10000)
    except UnicodeDecodeError:
        with open(path, "r", encoding="cp1251") as f:
            return f.read(10000)


def _get_mime_type(ext: str) -> str:
    mapping = {
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png",
        ".tiff": "image/tiff",
        ".tif": "image/tiff",
        ".bmp": "image/bmp",
        ".webp": "image/webp",
    }
    return mapping.get(ext.lower(), "image/png")


def _build_vision_messages(images: list[bytes], ext: str = ".png") -> list[dict]:
    """Формирует сообщения для vision-модели."""
    content = [{"type": "text", "text": ANALYSIS_PROMPT}]
    mime = _get_mime_type(ext)
    for img in images:
        b64 = _image_to_base64(img)
        content.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:{mime};base64,{b64}",
            },
        })
    return [{"role": "user", "content": content}]


def _build_text_messages(text: str) -> list[dict]:
    """Формирует сообщения для текстовой модели."""
    return [
        {
            "role": "user",
            "content": f"{ANALYSIS_PROMPT}\n\nТекст документа:\n\n{text}",
        }
    ]


async def _call_openrouter(
    client: httpx.AsyncClient,
    api_key: str,
    model: str,
    messages: list[dict],
) -> dict:
    """Вызов OpenRouter API."""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://docsorter.app",
        "X-Title": "DocSorter",
    }
    payload = {
        "model": model,
        "messages": messages,
        "temperature": 0.1,
        "max_tokens": 1000,
    }
    resp = await client.post(
        OPENROUTER_URL,
        json=payload,
        headers=headers,
        timeout=120.0,
    )
    resp.raise_for_status()
    data = resp.json()

    raw_text = data["choices"][0]["message"]["content"].strip()
    # Убираем markdown-обёртку если есть
    if raw_text.startswith("```"):
        lines = raw_text.split("\n")
        # Убираем первую и последнюю строки с ```
        lines = [l for l in lines if not l.strip().startswith("```")]
        raw_text = "\n".join(lines)

    return json.loads(raw_text)


async def analyze_file(
    file_info: dict,
    api_key: str,
    vision_model: str,
    text_model: str,
    max_pages: int,
    semaphore: asyncio.Semaphore,
    client: httpx.AsyncClient,
) -> dict:
    """Анализирует один файл и возвращает метаданные."""
    async with semaphore:
        path = file_info["path"]
        ext = file_info["ext"]
        file_type = get_file_type(ext)

        try:
            if file_type == "image":
                img_data = _read_image_file(path)
                messages = _build_vision_messages([img_data], ext)
                result = await _call_openrouter(client, api_key, vision_model, messages)

            elif file_type == "pdf":
                images = _pdf_to_images(path, max_pages)
                if images:
                    messages = _build_vision_messages(images, ".png")
                    result = await _call_openrouter(client, api_key, vision_model, messages)
                else:
                    result = _empty_result("Пустой PDF")

            elif file_type == "docx":
                text = _extract_docx_text(path)
                if text.strip():
                    messages = _build_text_messages(text)
                    result = await _call_openrouter(client, api_key, text_model, messages)
                else:
                    result = _empty_result("Пустой документ")

            elif file_type == "xlsx":
                text = _extract_xlsx_text(path)
                if text.strip():
                    messages = _build_text_messages(text)
                    result = await _call_openrouter(client, api_key, text_model, messages)
                else:
                    result = _empty_result("Пустая таблица")

            elif file_type == "text":
                text = _extract_text(path)
                if text.strip():
                    messages = _build_text_messages(text)
                    result = await _call_openrouter(client, api_key, text_model, messages)
                else:
                    result = _empty_result("Пустой файл")

            else:
                result = _empty_result("Неподдерживаемый формат")

        except json.JSONDecodeError:
            result = _empty_result("Ошибка разбора ответа LLM")
        except httpx.HTTPStatusError as e:
            result = _empty_result(f"Ошибка API: {e.response.status_code}")
        except Exception as e:
            result = _empty_result(f"Ошибка: {str(e)[:100]}")

        # Дополняем информацией о файле
        result["_file_name"] = file_info["name"]
        result["_file_path"] = str(file_info["path"])
        result["_rel_path"] = file_info["rel_path"]
        result["_ext"] = ext
        return result


def _empty_result(reason: str) -> dict:
    return {
        "doc_type": "Прочее",
        "title": reason,
        "number": "",
        "date": "",
        "counterparty": "",
        "reference": "",
        "summary": reason,
    }


async def analyze_batch(
    files: list[dict],
    api_key: str,
    vision_model: str,
    text_model: str,
    max_pages: int = 3,
    max_concurrent: int = 5,
    progress_callback=None,
) -> list[dict]:
    """
    Анализирует все файлы с батчингом.
    progress_callback(current, total) вызывается после каждого файла.
    """
    semaphore = asyncio.Semaphore(max_concurrent)
    results = []

    async with httpx.AsyncClient() as client:
        tasks = []
        for file_info in files:
            task = analyze_file(
                file_info, api_key, vision_model, text_model,
                max_pages, semaphore, client,
            )
            tasks.append(task)

        # Обрабатываем по мере завершения
        completed = 0
        for coro in asyncio.as_completed(tasks):
            result = await coro
            results.append(result)
            completed += 1
            if progress_callback:
                progress_callback(completed, len(files))

    # Сортируем по оригинальному порядку файлов
    file_order = {str(f["path"]): i for i, f in enumerate(files)}
    results.sort(key=lambda r: file_order.get(r.get("_file_path", ""), 0))

    return results
