"""
Модуль нарезки многодокументных PDF.
LLM (vision) определяет границы документов внутри PDF по батчам страниц,
программа режет PDF через PyMuPDF.
"""

import asyncio
import base64
import json
import shutil
from pathlib import Path

import httpx
import fitz  # PyMuPDF

from analyzer import OPENROUTER_URL, _image_to_base64, _get_mime_type

SLICE_SUBDIR = "_sliced"

SLICING_PROMPT = """Ты — помощник юриста. Тебе даны страницы с {batch_start} по {batch_end} из общего PDF-файла ({total} страниц).

{prev_context}

Определи границы отдельных документов в этом батче страниц. Документ = цельный юридический/деловой документ (договор, УПД, счёт, акт, письмо и т.д.).

Верни СТРОГО JSON-массив (без markdown, без ```json):

[
  {{
    "doc_type": "Договор / УПД / Счёт / Акт / Письмо / ...",
    "title": "краткое название документа",
    "page_from": N,
    "page_to": M,
    "continues_previous": true/false
  }},
  ...
]

Правила:
- page_from и page_to — АБСОЛЮТНЫЕ номера страниц в PDF (начиная с 1), не относительные к батчу.
- continues_previous=true только для ПЕРВОГО документа в батче, если он является продолжением документа из предыдущего батча.
- Документы в массиве должны идти последовательно без пропусков страниц.
- Сумма всех диапазонов (page_to - page_from + 1) должна равняться количеству страниц в батче ({batch_size}).
- Верни ТОЛЬКО JSON-массив.
"""


def _render_page(doc, page_num: int, dpi: int = 200) -> bytes:
    """Рендерит одну страницу PDF в PNG."""
    page = doc[page_num]
    pix = page.get_pixmap(dpi=dpi)
    return pix.tobytes("png")


async def _call_slicing(
    client: httpx.AsyncClient,
    api_key: str,
    model: str,
    images: list[bytes],
    prompt: str,
) -> list[dict]:
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://docsorter.app",
        "X-Title": "DocSorter",
    }
    content = [{"type": "text", "text": prompt}]
    for img in images:
        b64 = _image_to_base64(img)
        content.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{b64}"},
        })

    payload = {
        "model": model,
        "messages": [{"role": "user", "content": content}],
        "temperature": 0.1,
        "max_tokens": 2000,
    }
    resp = await client.post(
        OPENROUTER_URL, json=payload, headers=headers, timeout=300.0,
    )
    resp.raise_for_status()
    data = resp.json()
    raw = data["choices"][0]["message"]["content"].strip()
    if raw.startswith("```"):
        lines = raw.split("\n")
        lines = [l for l in lines if not l.strip().startswith("```")]
        raw = "\n".join(lines)
    return json.loads(raw)


async def analyze_pdf_structure(
    pdf_path: Path,
    api_key: str,
    vision_model: str,
    batch_size: int = 10,
    progress_callback=None,
) -> list[dict]:
    """
    Определяет структуру многодокументного PDF через LLM.
    Возвращает список сегментов: [{doc_type, title, page_from, page_to}, ...]
    """
    doc = fitz.open(str(pdf_path))
    total_pages = doc.page_count

    try:
        segments = []

        async with httpx.AsyncClient() as client:
            for batch_start in range(0, total_pages, batch_size):
                batch_end = min(batch_start + batch_size, total_pages)
                batch_len = batch_end - batch_start

                # Рендерим страницы батча
                images = [_render_page(doc, i) for i in range(batch_start, batch_end)]

                # Контекст предыдущих сегментов
                if segments:
                    last = segments[-1]
                    prev_context = (
                        f"Предыдущий батч закончился документом '{last.get('title', '')}' "
                        f"(тип: {last.get('doc_type', '')}), "
                        f"начатым со страницы {last.get('page_from')} и "
                        f"окончившимся на странице {last.get('page_to')}. "
                        f"Если первый документ текущего батча — его продолжение, "
                        f"поставь continues_previous=true."
                    )
                else:
                    prev_context = "Это первый батч."

                prompt = SLICING_PROMPT.format(
                    batch_start=batch_start + 1,
                    batch_end=batch_end,
                    total=total_pages,
                    batch_size=batch_len,
                    prev_context=prev_context,
                )

                batch_segments = await _call_slicing(
                    client, api_key, vision_model, images, prompt,
                )

                # Склейка: если первый сегмент — продолжение предыдущего
                if batch_segments and batch_segments[0].get("continues_previous") and segments:
                    segments[-1]["page_to"] = batch_segments[0].get(
                        "page_to", segments[-1]["page_to"]
                    )
                    batch_segments = batch_segments[1:]

                # Удаляем служебное поле continues_previous из остальных
                for seg in batch_segments:
                    seg.pop("continues_previous", None)

                segments.extend(batch_segments)

                if progress_callback:
                    progress_callback(batch_end, total_pages)

        # Сортируем и проверяем последовательность
        segments.sort(key=lambda s: s.get("page_from", 0))
        return segments

    finally:
        doc.close()


def verify_segments(segments: list[dict], total_pages: int) -> tuple[bool, str]:
    """
    Проверяет что сегменты покрывают все страницы без пропусков и наложений.
    Возвращает (ok, error_msg).
    """
    if not segments:
        return False, "Нет сегментов"

    # Проверяем границы
    for seg in segments:
        pf, pt = seg.get("page_from"), seg.get("page_to")
        if pf is None or pt is None:
            return False, "В сегменте отсутствует page_from или page_to"
        if not (1 <= pf <= pt <= total_pages):
            return False, f"Некорректные границы: {pf}-{pt} (всего {total_pages} стр.)"

    # Сортируем
    sorted_segs = sorted(segments, key=lambda s: s["page_from"])

    # Проверяем покрытие без пропусков/наложений
    if sorted_segs[0]["page_from"] != 1:
        return False, f"Первый сегмент начинается со стр. {sorted_segs[0]['page_from']}, а не с 1"

    for i in range(1, len(sorted_segs)):
        if sorted_segs[i]["page_from"] != sorted_segs[i - 1]["page_to"] + 1:
            return (
                False,
                f"Разрыв/наложение между сегментами: "
                f"{sorted_segs[i - 1]['page_from']}-{sorted_segs[i - 1]['page_to']} и "
                f"{sorted_segs[i]['page_from']}-{sorted_segs[i]['page_to']}",
            )

    if sorted_segs[-1]["page_to"] != total_pages:
        return (
            False,
            f"Последний сегмент заканчивается на стр. {sorted_segs[-1]['page_to']}, "
            f"а не на {total_pages}",
        )

    total_covered = sum(s["page_to"] - s["page_from"] + 1 for s in sorted_segs)
    if total_covered != total_pages:
        return False, f"Сумма страниц сегментов ({total_covered}) ≠ всего ({total_pages})"

    return True, ""


def slice_pdf(
    pdf_path: Path,
    segments: list[dict],
    work_dir: Path,
) -> list[Path]:
    """
    Разрезает PDF согласно сегментам.
    Создаёт файлы в <work_dir>/<имя_оригинала>/ и возвращает список путей.

    work_dir — рабочая папка для нарезок (НЕ исходная папка с документами).
    Каллер должен передать сюда папку, отдельную от source_dir.
    """
    pdf_path = Path(pdf_path)
    work_dir = Path(work_dir)
    from sorter import sanitize_filename

    out_dir = work_dir / pdf_path.stem
    out_dir.mkdir(parents=True, exist_ok=True)

    src = fitz.open(str(pdf_path))
    out_paths = []

    try:
        for i, seg in enumerate(sorted(segments, key=lambda s: s["page_from"]), start=1):
            pf = seg["page_from"] - 1
            pt = seg["page_to"] - 1

            new_doc = fitz.open()
            new_doc.insert_pdf(src, from_page=pf, to_page=pt)

            doc_type = sanitize_filename(seg.get("doc_type", "Документ"))
            title = sanitize_filename(seg.get("title", ""))
            if title and title != doc_type:
                base_name = f"{i:02d}_{doc_type}_{title}"
            else:
                base_name = f"{i:02d}_{doc_type}"

            # Обрезаем на всякий
            base_name = base_name[:150]
            out_path = out_dir / f"{base_name}.pdf"

            new_doc.save(str(out_path))
            new_doc.close()
            out_paths.append(out_path)
    finally:
        src.close()

    return out_paths


def undo_slice(original_path: str, slice_parts: list[str]) -> None:
    """Удаляет нарезанные файлы и подпапку."""
    for part in slice_parts:
        try:
            Path(part).unlink(missing_ok=True)
        except Exception:
            pass

    # Пытаемся удалить пустую подпапку
    if slice_parts:
        parent = Path(slice_parts[0]).parent
        try:
            if parent.exists() and not any(parent.iterdir()):
                parent.rmdir()
        except Exception:
            pass
