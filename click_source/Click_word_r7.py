# -*- coding: utf-8 -*-
"""
Click_word_r7 - Вставка изображений (PNG/SVG) в документ Р7-Офис
через Document Builder API (аналог COM Word для Microsoft Office).

Требует: Р7-Офис с установленным Документ Конструктор (docconstructor).
"""
import base64
import logging
import subprocess
import tempfile
from os import path as os_path
from pathlib import Path

from config import OUTPUT_FILENAME

logger = logging.getLogger(__name__)


# Пути к docconstructor (проверять в порядке приоритета)
DOCCONSTRUCTOR_PATHS = [
    r'C:\Program Files\R7-Office\docconstructor.exe',
    r'C:\Program Files (x86)\R7-Office\docconstructor.exe',
    '/opt/r7-office/docconstructor',
    '/usr/bin/docconstructor',
]


def find_docconstructor():
    """Найти исполняемый файл Document Builder."""
    for p in DOCCONSTRUCTOR_PATHS:
        if os_path.isfile(p):
            return p
    return None


def process_directory_r7(rg2, path_wrd, format_value, position_value, text_value,
                        image_format='svg', docconstructor_path=None):
    """
    Вставка SVG/PNG в документ через Р7 Document Builder.
    Создаёт output.docx, совместимый с Р7-Офис.

    Args:
        rg2: список путей к rg2 файлам
        path_wrd: путь к шаблону .docx
        format_value: A4 или A3
        position_value: Книжная или Альбомная
        text_value: использовать стили приложения
        image_format: 'svg' или 'png'
        docconstructor_path: путь к docconstructor (опционально)
    """
    doccon = docconstructor_path or find_docconstructor()
    if not doccon:
        raise FileNotFoundError(
            'Р7 Document Builder (docconstructor) не найден. '
            'Установите Р7-Офис с Документ Конструктор.'
        )

    # Собираем пути к изображениям
    image_files = []
    for rg2_path in rg2:
        directory = os_path.dirname(rg2_path)
        base_name = Path(rg2_path).stem
        img_path = os_path.join(directory, base_name + f'.{image_format}')
        if os_path.exists(img_path):
            image_files.append(img_path)
        else:
            print(f'Файл не найден: {img_path}')

    if not image_files:
        raise ValueError('Нет найденных файлов изображений')

    # Размер в EMU (1 cm ≈ 36000 EMU)
    width_emu = 14 * 36000  # 14 см
    height_emu = 14 * 36000

    # MIME для Base64
    mime = 'image/svg+xml' if image_format == 'svg' else 'image/png'

    # Читаем изображения и кодируем в Base64
    images_data = []
    for img_path in image_files:
        with open(img_path, 'rb') as f:
            data = base64.b64encode(f.read()).decode('ascii')
        images_data.append({
            'src': f'data:{mime};base64,{data}',
            'caption': os_path.basename(img_path).replace('\\', ' / '),
        })

    template_path = os_path.abspath(path_wrd).replace('\\', '\\\\')
    output_dir = os_path.dirname(rg2[0])
    output_path = os_path.abspath(os_path.join(output_dir, OUTPUT_FILENAME)).replace('\\', '\\\\')

    # Генерируем скрипт (Base64 в файле — без ограничения длины командной строки)
    script_parts = [
        'builder.SetTmpFolder("DocBuilderTemp");',
        f'builder.OpenFile("{template_path}");',
        'var oDocument = Api.GetDocument();',
        f'var w = {width_emu}, h = {height_emu};',
        'var oParagraph, oImage;',
    ]

    for img in images_data:
        # Экранируем для JS: \ " и переводы строк
        src = img['src'].replace('\\', '\\\\').replace('"', '\\"').replace('\r', '').replace('\n', '')
        cap = img['caption'].replace('\\', '\\\\').replace('"', '\\"').replace('\r', ' ').replace('\n', ' ')
        script_parts.extend([
            'oParagraph = Api.CreateParagraph();',
            f'oImage = Api.CreateImage("{src}", w, h);',
            'oParagraph.AddDrawing(oImage);',
            'oDocument.Push(oParagraph);',
            'oParagraph = Api.CreateParagraph();',
            f'oParagraph.AddText("{cap}");',
            'oDocument.Push(oParagraph);',
            'oParagraph = Api.CreateParagraph();',
            'oParagraph.AddLineBreak();',
            'oDocument.Push(oParagraph);',
        ])

    script_parts.extend([
        f'builder.SaveFile("docx", "{output_path}");',
        'builder.CloseFile();',
    ])

    script_content = '\n'.join(script_parts)

    with tempfile.NamedTemporaryFile(
            mode='w', suffix='.docbuilder', delete=False, encoding='utf-8'
    ) as f:
        f.write(script_content)
        script_path = f.name

    try:
        result = subprocess.run(
            [doccon, script_path],
            capture_output=True,
            text=True,
            timeout=300,
            cwd=output_dir,
        )
        if result.returncode != 0:
            logger.error('Document Builder: %s', result.stderr or result.stdout)
            raise RuntimeError(f'Document Builder завершился с кодом {result.returncode}')
    finally:
        try:
            os_path.unlink(script_path)
        except OSError:
            pass

    logger.info('Документ сохранён: %s', output_path)


if __name__ == '__main__':
    process_directory_r7(
        rg2=[r'C:\test\file.rg2'],
        path_wrd=r'C:\test\Шаблон.docx',
        format_value='A4',
        position_value='Книжная',
        text_value=False,
        image_format='svg',
    )
