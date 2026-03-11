# -*- coding: utf-8 -*-
"""
Click_Rastr - Интеграция с RastrWin3 для экспорта RG2 в DXF, затем в PNG/SVG.
Требует установленный RastrWin3 и Windows COM.
"""
from win32com.client import Dispatch
from pathlib import Path
from os import path as os_path
import winreg
from winreg import OpenKey, QueryValueEx, CloseKey, HKEY_CURRENT_USER, KEY_READ

from ezdxf import readfile, new
from Click_dxf2png import convert_dxf_to_png
from Click_dxf2svg import dxf2svg

Rastr = Dispatch('Astra.Rastr')
grf = Dispatch('Graph.GraphRastr')


def get_reg(name):
    """Получить значение из реестра RastrWin3."""
    try:
        registry_key = OpenKey(HKEY_CURRENT_USER, 'SOFTWARE\\RastrWin3', 0, KEY_READ)
        value, _ = QueryValueEx(registry_key, name)
        CloseKey(registry_key)
        return value
    except WindowsError:
        return None


def updategrf():
    """Обновить графику из Rastr."""
    grf.ReadFromRastrInt(
        Rastr,
        'node[ny,uhom,na,sta,sel,tip,_epn,_epg,_esh,_ekn],vetv[ip,iq,np,sta,sel,tip,_zbg,_zen,na,groupid],Generator[Num,Node,sta]',
        'graph_node,graph_vetv,graph_text,graph_figur,graph_nadp,graph_com,graph_settext,graph_kadr,graph_unom,graph_area,graph_params,graph2_block,graph2_drawitems,graph2_selvalues,graph2_path'
    )


def save_dxf(file, doc):
    """Сохранить DXF, удаляя круги."""
    text_style = doc.styles.get('Standard')
    text_style.dxf.font = 'Arial.ttf'
    doc.dxfversion = 'R2018'
    msp = doc.modelspace()
    for entity in msp:
        if entity.dxftype() == 'CIRCLE':
            msp.delete_entity(entity)
    doc.saveas(file)


def filter_specific_entities(file):
    """
    Создает копию документа новой версии с нужными типами объектов.
    Возвращает объект с типами элементов: линии, текст и т.д.
    """
    doc = readfile(file, encoding='windows-1251')
    filtered_doc = new('R2018')
    text_style = filtered_doc.styles.get('Standard')
    text_style.dxf.font = 'Arial.ttf'
    allowed_types = ('LINE', 'LWPOLYLINE', 'POLYLINE', 'TEXT', 'MTEXT', 'CIRCLE')
    msp_filtered = filtered_doc.modelspace()
    for entity in doc.modelspace():
        if entity.dxftype() in allowed_types:
            msp_filtered.add_entity(entity.copy())
    filtered_doc.saveas(file)
    return filtered_doc


def create_dxf(file, path_grf, com_value, path_sch=None):
    """
    Создает файлы PNG/SVG из файлов RG2.

    Args:
        file (str): путь к файлу режима *.rg2
        path_grf (str): путь к файлу графики *.grf
        com_value (str): выбор типа создаваемой картинки ('PNG' или 'SVG')
        path_sch (str): путь к файлу сечений *.sch (по умолчанию None)
    """
    dict_sch1 = {}
    dict_sch2 = {}

    dxf_path = os_path.splitext(file)[0] + '.dxf'
    png_path = os_path.splitext(file)[0] + '.png'
    svg_path = os_path.splitext(file)[0] + '.svg'

    path_shabl = Path(get_reg('UserFolder'))
    shabl_dir = path_shabl / 'SHABLON'
    shabl = shabl_dir / 'режим.rg2'
    shabl_sch = shabl_dir / 'сечения.sch'
    shabl_grf = shabl_dir / 'графика.grf'

    Rastr.ClearControl()
    Rastr.Load(1, file, str(shabl))
    Rastr.Load(1, path_grf, str(shabl_grf))

    kod = Rastr.rgm('')
    if kod == 0:
        updategrf()
        grf.ImportDXF(dxf_path)

    if path_sch:
        Rastr.Load(1, path_sch, str(shabl_sch))
        sch = Rastr.Tables('sechen')
        name_sch = sch.Cols('name')
        val_sch = sch.Cols('psech')
        sch.SetSel('sta = 1')
        i = sch.FindNextSel(-1)
        while i != -1:
            dict_sch2[round(val_sch.Z[i])] = name_sch.Z[i]
            dict_sch1[os_path.basename(os_path.splitext(file)[0])] = dict_sch2.copy()
            i = sch.FindNextSel(i)

    if com_value == 'PNG':
        convert_dxf_to_png(dxf_path, png_path, dict_sch1)
    else:
        dxf2svg(dxf_path, svg_path)


if __name__ == '__main__':
    create_dxf(
        r'C:\Users\Gurev-AA\Desktop\тест\Вот_ГЭС2\.28_1_Р_11_ФПТ_0.rg2',
        r'C:\Users\Gurev-AA\Desktop\тест\Графика_ВотГЭС_5.0.grf'
    )
