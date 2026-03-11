# -*- coding: utf-8 -*-
"""
Click_dxf2png - Конвертация DXF в PNG с использованием matplotlib и ezdxf.
"""
from os import path
import matplotlib
matplotlib.use('Agg')
from matplotlib import patches
import matplotlib.pyplot as plt
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
from ezdxf import bbox, readfile


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
        print('Не найдено геометрических элементов для определения границ')
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
    except Exception as e:
        print('Ошибка при вычислении границ:', e)
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


def create_flags(msp, dxf_file, dict_sch, ax):
    """Создание флагов (прямоугольников с подписями) на графике."""
    from ezdxf.bbox import Cache
    cache = Cache()
    msp_bbox = bbox.extents(msp, cache=cache)
    xmin, ymin = msp_bbox.extmin.x, msp_bbox.extmin.y
    xmax, ymax = msp_bbox.extmax.x, msp_bbox.extmax.y
    drawing_width = xmax - xmin
    drawing_height = ymax - ymin
    rect_width = drawing_width * 0.1
    rect_height = drawing_height * 0.025
    base_name = path.splitext(path.basename(dxf_file))[0]
    for i, (name, val) in enumerate(dict_sch.get(base_name, {}).items()):
        rect_x = xmin + (i % 10) * rect_width * 1.1
        rect_y = ymax - rect_height * (1 + i // 10)
        rect1 = patches.Rectangle((rect_x, rect_y), rect_width, rect_height,
                                  linewidth=0.5, edgecolor='black', facecolor='yellow')
        ax.add_patch(rect1)
        ax.text(rect_x + rect_width / 2, rect_y + rect_height / 2, str(val),
                ha='center', va='center', fontsize=8)


def convert_dxf_to_png(dxf_file, output_file, dict_sch=None, dpi=600):
    """
    Конвертирует DXF-файл в PNG-изображение.

    Args:
        dxf_file (str): Путь к DXF-файлу
        output_file (str): Имя выходного PNG-файла
        dict_sch (dict): Вложенный словарь сечений для создания флагов
        dpi (int): Разрешение изображения
    """
    if dict_sch is None:
        dict_sch = {}
    try:
        doc = readfile(dxf_file, encoding='windows-1251')
    except Exception as e:
        print('DXF-файл поврежден и не может быть конвертирован!', e)
        return

    width, height = get_drawing_size(doc)
    min_point, max_point = (0, 0), (width, height)
    msp = doc.modelspace()

    drawing_complexity = len([e for e in msp if e.dxftype() in ('TEXT', 'MTEXT', 'POLYLINE', 'LINE')])
    rec_dpi = min(1000, max(600, 700 + drawing_complexity))
    rec_lineweight = 0.1
    padding = 0.1
    base_scale = 1.0
    fig_width = max(5, width * base_scale / 100)
    fig_height = max(5, height * base_scale / 100)
    max_inches = 1200 / rec_dpi
    fig_width = min(fig_width, max_inches)
    fig_height = min(fig_height, max_inches)

    remove_distant_texts(msp)

    fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=rec_dpi)
    ax.set_facecolor('white')

    ctx = RenderContext(doc)
    backend = MatplotlibBackend(ax)
    Frontend(ctx, backend).draw_layout(msp, finalize=True)

    if dict_sch:
        try:
            create_flags(msp, dxf_file, dict_sch, ax)
        except Exception as e:
            print('Не удалось добавить прямоугольник с названием:', e)

    plt.savefig(output_file, dpi=rec_dpi, bbox_inches='tight', facecolor='white')
    plt.close()
