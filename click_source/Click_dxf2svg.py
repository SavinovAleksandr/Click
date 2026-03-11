# -*- coding: utf-8 -*-
"""
Click_dxf2svg - Конвертация DXF в SVG с использованием ezdxf.
"""
import logging
import re
from ezdxf import readfile, DXFStructureError
from ezdxf.addons.drawing import Frontend, RenderContext, svg, layout

from dxf_utils import get_drawing_size, remove_distant_texts

logger = logging.getLogger(__name__)


def dxf2svg(dxf_path, svg_path):
    """
    Конвертирует DXF-файл в SVG-файл.

    Args:
        dxf_path (str): Путь к входному DXF-файлу
        svg_path (str): Путь к выходному SVG-файлу
    """
    try:
        doc = readfile(dxf_path, encoding='windows-1251')
    except (DXFStructureError, IOError) as e:
        logger.error('Ошибка чтения DXF-файла: %s', e)
        return

    width, height = get_drawing_size(doc)
    min_point, max_point = (0, 0), (width, height)
    msp = doc.modelspace()

    remove_distant_texts(msp)

    context = RenderContext(doc)
    backend = svg.SVGBackend()
    frontend = Frontend(context, backend)
    frontend.draw_layout(msp, finalize=True)

    page = layout.Page(width, height, layout.Units.mm, layout.Margins.all(0))
    svg_output = backend.get_string(page)

    pattern = r'<rect\s+fill\s*=\s*"[^"]*"\s+x\s*=\s*"0"\s+y\s*=\s*"0"'
    replacement = '<rect fill="#ffffff" x="0" y="0"'
    new_content = re.sub(pattern, replacement, svg_output)
    with open(svg_path, 'w', encoding='utf-8') as fp:
        fp.write(new_content)


if __name__ == '__main__':
    dxf2svg(
        r'C:\Users\Gurev-AA\Desktop\тест\Григорий\да.dxf',
        r'C:\Users\Gurev-AA\Desktop\тест\Григорий\Pupupu.svg'
    )
