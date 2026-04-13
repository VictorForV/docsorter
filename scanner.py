"""
Модуль сканирования файлов.
Рекурсивный обход папки, сбор поддерживаемых файлов.
"""

from pathlib import Path

SUPPORTED_EXTENSIONS = {
    # Изображения
    ".jpg", ".jpeg", ".png", ".tiff", ".tif", ".bmp", ".webp",
    # Документы
    ".pdf",
    ".docx", ".doc",
    ".xlsx", ".xls",
    # Текст
    ".txt", ".rtf",
}


def scan_folder(folder: Path) -> list[dict]:
    """
    Рекурсивно сканирует папку и возвращает список файлов.
    Каждый файл: {"path": Path, "name": str, "ext": str, "size": int, "rel_path": str}
    """
    folder = Path(folder)
    if not folder.exists() or not folder.is_dir():
        raise ValueError(f"Папка не найдена: {folder}")

    files = []
    for item in sorted(folder.rglob("*")):
        if not item.is_file():
            continue
        if item.name.startswith("."):
            continue
        ext = item.suffix.lower()
        if ext not in SUPPORTED_EXTENSIONS:
            continue
        files.append({
            "path": item,
            "name": item.name,
            "ext": ext,
            "size": item.stat().st_size,
            "rel_path": str(item.relative_to(folder)),
        })

    return files


def count_files(folder: Path) -> int:
    """Считает количество файлов (всех) рекурсивно."""
    return sum(1 for f in Path(folder).rglob("*") if f.is_file())


def get_file_type(ext: str) -> str:
    """Определяет тип файла для выбора метода обработки."""
    ext = ext.lower()
    if ext in {".jpg", ".jpeg", ".png", ".tiff", ".tif", ".bmp", ".webp"}:
        return "image"
    if ext == ".pdf":
        return "pdf"
    if ext in {".docx", ".doc"}:
        return "docx"
    if ext in {".xlsx", ".xls"}:
        return "xlsx"
    if ext in {".txt", ".rtf"}:
        return "text"
    return "unknown"
