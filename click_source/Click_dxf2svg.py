# -*- coding: utf-8 -*-
"""
Click_dxf2svg - Конвертация DXF в SVG с использованием ezdxf.
"""
from ezdxf import readfile, bbox, DXFStructureError
from ezdxf.addons.drawing import Frontend, RenderContext, svg, layout
import re


def get_drawing_size(doc):
    """Возвращает размеры чертежа в единицах DXF."""
    msp = doc.modelspace()
    try:
        bounding_box = bbox.extents(msp)
        if not bounding_box.has_data:
            limits = doc.header.get('$LIMMAX', (100, 100))
            return limits[0] - 0, limits[1] - 0
        min_point, max_point = bounding_box.extmin, bounding_box.extmax
        width = max_point.x - min_point.x
        height = max_point.y - min_point.y
        return width, height
    except Exception:
        return 100, 100


def remove_distant_texts(msp, margin_percent=10):
    """
    Удаляет текстовые элементы, находящиеся далеко от границ чертежа.
    """
    geometric_entities = [e for e in msp if e.dxftype() in (
        'LINE', 'CIRCLE', 'LWPOLYLINE', 'POLYLINE', 'ARC', 'ELLIPSE', 'SPLINE'
    )]
    if not geometric_entities:
        return
    try:
        bounding_box = bbox.extents(geometric_entities)
        min_point, max_point = bounding_box.extmin, bounding_box.extmax
        width = max_point.x - min_point.x
        height = max_point.y - min_point.y
        margin_x = width * margin_percent / 100
        margin_y = height * margin_percent / 100
        min_x, max_x = min_point.x - margin_x, max_point.x + margin_x
        min_y, max_y = min_point.y - margin_y, max_point.y + margin_y
    except Exception:
        return

    texts_to_remove = []
    for entity in msp:
        if entity.dxftype() in ('TEXT', 'MTEXT'):
            try:
                entity_bbox = bbox.extents([entity])
                pos = entity_bbox.extmin
                if pos.x < min_x or pos.x > max_x or pos.y < min_y or pos.y > max_y:
                    texts_to_remove.append(entity)
            except Exception:
                pass

    for text in texts_to_remove:
        try:
            msp.delete_entity(text)
        except Exception:
            pass


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
        print('Error reading DXF file:', e)
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
