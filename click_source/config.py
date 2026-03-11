# -*- coding: utf-8 -*-
"""
config - Константы и конфигурация макроса Click.
"""
from pathlib import Path

# Стили Word
STYLE_HEADING_APP = '7.32 ЗАГОЛОВОК ПРИЛОЖЕНИЙ (отображается в СОДЕРЖАНИИ'
STYLE_ILLUSTRATION = '7.32 Иллюстрация. Подпись к иллюстрации'
STYLE_ILLUSTRATION_APP = '7.32 Иллюстрация ПРИЛОЖЕНИЯ. Подпись к иллюстрации'
STYLE_ILLUSTRATION_LAYOUT = '7.32 Иллюстрация. Расположение иллюстрации'

# Имя выходного файла
OUTPUT_FILENAME = 'output.docx'


def get_icon_path():
    """Путь к иконке (рядом с исполняемым или в папке проекта)."""
    base = Path(__file__).parent
    for name in ('icon.ico', 'diaphragm.ico', 'click.ico'):
        p = base / name
        if p.exists():
            return str(p)
    return None
