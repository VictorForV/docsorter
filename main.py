"""
DocSorter — Сортировщик документов для юристов.
Точка входа приложения.
"""

import sys
import os

# Для PyInstaller: добавляем директорию exe в PATH
if getattr(sys, "frozen", False):
    os.chdir(os.path.dirname(sys.executable))

from gui import DocSorterApp


def main():
    app = DocSorterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
