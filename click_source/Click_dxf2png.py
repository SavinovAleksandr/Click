# -*- coding: utf-8 -*-
"""
Click_dxf2png - Конвертация DXF в PNG с использованием matplotlib и ezdxf.
"""
import logging
from os import path
import matplotlib
matplotlib.use('Agg')
from matplotlib import patches
import matplotlib.pyplot as plt
from ezdxf.addons.drawing import RenderContext, Frontend
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
from ezdxf import bbox, readfile

from dxf_utils import get_drawing_size, remove_distant_texts

logger = logging.getLogger(__name__)


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
        logger.error('DXF-файл поврежден и не может быть конвертирован: %s', e)
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
            logger.warning('Не удалось добавить прямоугольник с названием: %s', e)

    plt.savefig(output_file, dpi=rec_dpi, bbox_inches='tight', facecolor='white')
    plt.close()
