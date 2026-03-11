# -*- coding: utf-8 -*-
"""
dxf_utils - Общие функции для работы с DXF (Click_dxf2png, Click_dxf2svg).
"""
import logging
from ezdxf import bbox

logger = logging.getLogger(__name__)

GEOMETRIC_TYPES = ('LINE', 'CIRCLE', 'LWPOLYLINE', 'POLYLINE', 'ARC', 'ELLIPSE', 'SPLINE')
TEXT_TYPES = ('TEXT', 'MTEXT')


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
    geometric_entities = [e for e in msp if e.dxftype() in GEOMETRIC_TYPES]
    if not geometric_entities:
        logger.warning('Не найдено геометрических элементов для определения границ')
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
        logger.warning('Ошибка при вычислении границ: %s', e)
        return

    texts_to_remove = []
    for entity in msp:
        if entity.dxftype() in TEXT_TYPES:
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
