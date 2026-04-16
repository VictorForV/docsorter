@echo off
echo === Сборка DocSorter ===

pip install -r requirements.txt

pyinstaller --noconfirm --onefile --windowed ^
    --name DocSorter ^
    --add-data "categories.json;." ^
    --hidden-import customtkinter ^
    --hidden-import httpx ^
    --hidden-import fitz ^
    --hidden-import docx ^
    --hidden-import openpyxl ^
    --hidden-import networkx ^
    --hidden-import matplotlib ^
    --hidden-import matplotlib.backends.backend_tkagg ^
    --collect-all customtkinter ^
    --collect-all matplotlib ^
    main.py

echo.
echo === Готово! Файл: dist\DocSorter.exe ===
pause
