# -*- coding:  utf-8 -*-
import os
from math import tan, radians

import pythoncom
from win32com.client import Dispatch, gencache

PART_NUMBER = '06'
PART_NAME = 'Поручень крайний прямой секции'
DIM_MIN = 500
DIM_MAX = 2000

CODE1 = '002'
CODE2 = '99'

# DXF_FOLDER = f'D:\\youriy\\NCC\\Перила и Лотки\\Catalog\\.{CODE1}\\.{PART_NUMBER} - {PART_NAME}\\DXF'
DXF_FOLDER = os.path.expanduser(f'~\\Desktop\\Work\\!Перила и лотки\\Kompas\\3 Каталог СТО - ЦРНС.305112\\.{CODE1}\\.{CODE2}\\DXF\\.{PART_NUMBER} - {PART_NAME}')
if not os.path.exists(DXF_FOLDER):
    os.makedirs(DXF_FOLDER)

#  Подключим описание интерфейсов API5
api5 = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
kompas_object = api5.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(api5.KompasObject.CLSID, pythoncom.IID_IDispatch))

kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants

#  Подключим описание интерфейсов API7
api7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
application = api7.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(api7.IApplication.CLSID, pythoncom.IID_IDispatch))

application.Visible = True

api5_document = kompas_object.ActiveDocument2D()
iDocument = application.ActiveDocument
# iKompasDocument1 = kompas_api7_module.IKompasDocument1(iDocument)
doc2D1 = api7.IKompasDocument2D1(iDocument)

circles = []
arks = []
lines = []
ellipses = []

# Управление слоями
# d = kompas_object.GetParamStruct(9)
#
# n = api5_document.ksLayer(3)
#
api5_document.ksLayer(0)
#
# api5_document.ksGetObjParam(n, d, -1)
# d.state = 2
# api5_document.ksSetObjParam(n, d, -1)


def circle(x, y):
    circles.append(api5_document.ksCircle(x, y, 16.25, 1))


def ellipse(x, y, a, b=16.5):
    ellipse_params = api5.ksEllipseParam(kompas_object.GetParamStruct(kompas6_constants.ko_EllipseParam))
    ellipse_params.Init()
    ellipse_params.xc = x
    ellipse_params.yc = y
    ellipse_params.A = a
    ellipse_params.B = b
    ellipse_params.angle = 0
    ellipse_params.style = 1
    ellipses.append(api5_document.ksEllipse(ellipse_params))


def ark(x, y, radius, quarter):
    if quarter == 1:
        arks.append(api5_document.ksArcByPoint(x, y, radius, x + radius, y, x, y + radius, 1, 1))
    elif quarter == 2:
        arks.append(api5_document.ksArcByPoint(x, y, radius, x, y + radius, x - radius, y, 1, 1))
    elif quarter == 3:
        arks.append(api5_document.ksArcByPoint(x, y, radius, x - radius, y, x, y - radius, 1, 1))
    elif quarter == 4:
        arks.append(api5_document.ksArcByPoint(x, y, radius, x, y - radius, x + radius, y, 1, 1))
    elif quarter == 5:
        arks.append(api5_document.ksArcByPoint(x, y, radius, x, y - radius, x, y + radius, 1, 1))
    elif quarter == 6:
        arks.append(api5_document.ksArcByPoint(x, y, radius, x, y + radius, x, y - radius, 1, 1))
    elif quarter == 7:
        arks.append(api5_document.ksArcByPoint(x, y, radius, x + radius, y, x - radius, y, 1, 1))
    elif quarter == 8:
        arks.append(api5_document.ksArcByPoint(x, y, radius, x - radius, y, x + radius, y, 1, 1))


def oval(x, y, width=20, height=10):
    r = height / 2
    lines.append(api5_document.ksLineSeg(x - width / 2 + r, y - height / 2, x + width / 2 - r, y - height / 2, 1))
    ark(x + width / 2 - r, y, r, 5)
    lines.append(api5_document.ksLineSeg(x + width / 2 - r, y + height / 2, x - width / 2 + r, y + height / 2, 1))
    ark(x - width / 2 + r, y, r, 6)


def pillar_hole(x, y):
    lines.append(api5_document.ksLineSeg(x + 45, y - 23.5, x + 45, y + 23.5, 1))
    ark(x + 38.5, y + 23.5, 6.5, 1)
    lines.append(api5_document.ksLineSeg(x + 38.5, y + 30, x + 32, y + 30, 1))
    lines.append(api5_document.ksLineSeg(x + 32, y + 30, x + 30.3, y + 31.7, 1))
    lines.append(api5_document.ksLineSeg(x + 30.3, y + 31.7, x - 30.3, y + 31.7, 1))
    lines.append(api5_document.ksLineSeg(x - 30.3, y + 31.7, x - 32, y + 30, 1))
    lines.append(api5_document.ksLineSeg(x - 32, y + 30, x - 38.5, y + 30, 1))
    ark(x - 38.5, y + 23.5, 6.5, 2)
    lines.append(api5_document.ksLineSeg(x - 45, y + 23.5, x - 45, y - 23.5, 1))
    ark(x - 38.5, y - 23.5, 6.5, 3)
    lines.append(api5_document.ksLineSeg(x - 38.5, y - 30, x - 32, y - 30, 1))
    lines.append(api5_document.ksLineSeg(x - 32, y - 30, x - 30.3, y - 31.7, 1))
    lines.append(api5_document.ksLineSeg(x - 30.3, y - 31.7, x + 30.3, y - 31.7, 1))
    lines.append(api5_document.ksLineSeg(x + 30.3, y - 31.7, x + 32, y - 30, 1))
    lines.append(api5_document.ksLineSeg(x + 32, y - 30, x + 38.5, y - 30, 1))
    ark(x + 38.5, y - 23.5, 6.5, 4)


def pillar_left_hole(x, y):
    lines.append(api5_document.ksLineSeg(x + 45, y - 23.5, x + 45, y + 23.5, 1))
    ark(x + 38.5, y + 23.5, 6.5, 1)
    lines.append(api5_document.ksLineSeg(x + 38.5, y + 30, x + 32, y + 30, 1))
    lines.append(api5_document.ksLineSeg(x + 32, y + 30, x + 30.3, y + 31.7, 1))
    lines.append(api5_document.ksLineSeg(x + 30.3, y + 31.7, x, y + 31.7, 1))
    lines.append(api5_document.ksLineSeg(x, y + 31.7, x, y + 44, 1))
    lines.append(api5_document.ksLineSeg(x, y - 44, x, y - 31.7, 1))
    lines.append(api5_document.ksLineSeg(x, y - 31.7, x + 30.3, y - 31.7, 1))
    lines.append(api5_document.ksLineSeg(x + 30.3, y - 31.7, x + 32, y - 30, 1))
    lines.append(api5_document.ksLineSeg(x + 32, y - 30, x + 38.5, y - 30, 1))
    ark(x + 38.5, y - 23.5, 6.5, 4)


def pillar_right_hole(x, y):
    lines.append(api5_document.ksLineSeg(x, y + 31.7, x, y + 44, 1))
    lines.append(api5_document.ksLineSeg(x, y - 44, x, y - 31.7, 1))
    lines.append(api5_document.ksLineSeg(x, y + 31.7, x - 30.3, y + 31.7, 1))
    lines.append(api5_document.ksLineSeg(x - 30.3, y + 31.7, x - 32, y + 30, 1))
    lines.append(api5_document.ksLineSeg(x - 32, y + 30, x - 38.5, y + 30, 1))
    ark(x - 38.5, y + 23.5, 6.5, 2)
    lines.append(api5_document.ksLineSeg(x - 45, y + 23.5, x - 45, y - 23.5, 1))
    ark(x - 38.5, y - 23.5, 6.5, 3)
    lines.append(api5_document.ksLineSeg(x - 38.5, y - 30, x - 32, y - 30, 1))
    lines.append(api5_document.ksLineSeg(x - 32, y - 30, x - 30.3, y - 31.7, 1))
    lines.append(api5_document.ksLineSeg(x - 30.3, y - 31.7, x, y - 31.7, 1))


def rectangle(x, y, length, width, left=False, right=False):
    length += x
    vert_side = width + y
    lines.append(api5_document.ksLineSeg(x, y, length, y, 1))
    if right:
        pillar_right_hole(length, y + width/2)
    else:
        lines.append(api5_document.ksLineSeg(length, y, length, vert_side, 1))
    lines.append(api5_document.ksLineSeg(length, vert_side, x, vert_side, 1))
    if left:
        pillar_left_hole(x, y + width/2)
    else:
        lines.append(api5_document.ksLineSeg(x, vert_side, x, y, 1))


def rectangle_rounded(x, y, length, width, radius):
    length += x
    width += y
    lines.append(api5_document.ksLineSeg(x + radius, y, length - radius, y, 1))
    ark(length - radius, y + radius, radius, 4)
    lines.append(api5_document.ksLineSeg(length, y + radius, length, width - radius, 1))
    ark(length - radius, width - radius, radius, 1)
    lines.append(api5_document.ksLineSeg(length - radius, width, x + radius, width, 1))
    ark(x + radius, width - radius, radius, 2)
    lines.append(api5_document.ksLineSeg(x, width - radius, x, y + radius, 1))
    ark(x + radius, y + radius, radius, 3)


def rectangle_views(length, width=40, angle1=90, angle2=90, x=0, y=0, left=False, right=False, through=True):
    left_ofset = round(width / tan(radians(angle1)), 6)
    right_ofset = round(width / tan(radians(angle2)), 6)

    side1 = side2 = length

    if left_ofset < 0:
        x -= left_ofset
        side1 += left_ofset
    elif left_ofset > 0:
        side2 -= left_ofset

    if right_ofset > 0:
        side2 -= right_ofset
    elif right_ofset < 0:
        side1 += right_ofset

    width2 = width
    if width == 58:
        width2 = 88

    # front view
    lines.append(api5_document.ksLineSeg(x, y, x + side1, y, 1))
    lines.append(api5_document.ksLineSeg(x + side1, y, x + side1 - right_ofset, y + width, 1))
    lines.append(api5_document.ksLineSeg(x + side1 - right_ofset, y + width, x + left_ofset, y + width, 1))
    lines.append(api5_document.ksLineSeg(x + left_ofset, y + width, x, y, 1))

    # bottom view
    y += width + 100
    rectangle(x, y, side1, width2, left, right)

    # rear view
    y += width2 + 100
    lines.append(api5_document.ksLineSeg(x, y + width, x + side1, y + width, 1))
    lines.append(api5_document.ksLineSeg(x + side1, y + width, x + side1 - right_ofset, y, 1))
    lines.append(api5_document.ksLineSeg(x + side1 - right_ofset, y, x + left_ofset, y, 1))
    lines.append(api5_document.ksLineSeg(x + left_ofset, y, x, y + width, 1))

    # top view
    y += width + 100
    if not through:
        left = False
        right = False
    rectangle(x + left_ofset, y, side2, width2, left, right)


def delete_objects():
    for item in circles:
        api5_document.ksDeleteObj(item)

    for item in arks:
        api5_document.ksDeleteObj(item)

    for item in lines:
        api5_document.ksDeleteObj(item)

    for item in ellipses:
        api5_document.ksDeleteObj(item)

    circles.clear()
    arks.clear()
    lines.clear()
    ellipses.clear()


def regular_qty_and_offset(length: int) -> (int, int):
    quantity = (length - 151) // 182
    left_offset = (length - quantity * 182) / 2
    return quantity, left_offset


def gap_qty_and_offset(length: int) -> (int, float, int):
    x = length - 106
    repetition, k = divmod(x, 182)
    quantity = repetition + 1
    if k < 75:
        left_offset = (x - 182 * (repetition - 1) - 16) / 2
        repetition -= 1
    elif k > 166:
        left_offset = (k - 16) / 2
        quantity += 1
    else:
        left_offset = k
    return quantity, float(left_offset), repetition


overhang_length = 250
# folder = ''

# # скошенная крайняя секция
left_offset = overhang_length - 116

folder = f'.{overhang_length}'
if not os.path.exists(os.path.join(DXF_FOLDER, folder)):
    os.makedirs(os.path.join(DXF_FOLDER, folder))

for n in range(DIM_MIN, DIM_MAX + 1, 10):
    ''' Балка 40х40 '''
    # qty, offset = regular_qty_and_offset(n)
    # vertical_offset = 40*3 + 300 + 20
    # rectangle_views(n)
    # for i in range(qty + 1):
    #     circle(offset + i*182, vertical_offset)

    ''' Балка 40х40 деформационного шва '''
    # qty, offset, repetitions = gap_qty_and_offset(n)
    # vertical_offset = 40*3 + 300 + 20
    # rectangle_views(n)
    # circle(offset, vertical_offset)
    # for i in range(repetitions + 1):
    #     circle(n - 106 - i*182, vertical_offset)

    ''' Поручень трехбалки '''
    # rectangle_views(n, width=58, left=True, right=True, through=False)
    # oval(20, 29)
    # oval(20, 58 + 100 + 88 + 100 + 29)
    # oval(n - 20, 29)
    # oval(n - 20, 58 + 100 + 88 + 100 + 29)

    ''' Поручень двойной трехбалки '''
    # rectangle_views(n, width=58, left=True, right=True, through=False)
    # pillar_hole(n / 2, 58 + 100 + 44)
    # oval(20, 29)
    # oval(n / 2 - 20, 29)
    # oval(n / 2 + 20, 29)
    # oval(n - 20, 29)
    # oval(20, 58 + 100 + 88 + 100 + 29)
    # oval(n / 2 - 20, 58 + 100 + 88 + 100 + 29)
    # oval(n / 2 + 20, 58 + 100 + 88 + 100 + 29)
    # oval(n - 20, 58 + 100 + 88 + 100 + 29)

    ''' Поручень крайний трехбалки для скошенного свеса длиной overhang_length '''
    # full_length = n + left_offset
    # rectangle_views(full_length, width=58, angle1=-55, right=True, through=False)
    # pillar_hole(left_offset, 58 + 100 + 44)
    # oval(left_offset - 20, 29)
    # oval(left_offset + 20, 29)
    # oval(full_length - 20, 29)
    # oval(left_offset - 20, 58 + 100 + 88 + 100 + 29)
    # oval(left_offset + 20, 58 + 100 + 88 + 100 + 29)
    # oval(full_length - 20, 58 + 100 + 88 + 100 + 29)

    ''' Поручень крайний трехбалки для прямого свеса длиной overhang_length '''
    full_length = n + overhang_length
    rectangle_views(full_length, width=58, angle1=-45, right=True, through=False)
    pillar_hole(overhang_length, 58 + 100 + 44)
    oval(overhang_length - 20, 29)
    oval(overhang_length + 20, 29)
    oval(full_length - 20, 29)
    oval(overhang_length - 20, 58 + 100 + 88 + 100 + 29)
    oval(overhang_length + 20, 58 + 100 + 88 + 100 + 29)
    oval(full_length - 20, 58 + 100 + 88 + 100 + 29)

    ''' Поручень двухбалки '''
    # rectangle_views(n, width=58, left=True, right=True, through=False)
    # qty, offset = regular_qty_and_offset(n - 88)
    # offset += 44
    # vertical_offset = 58 + 100 + 44
    # for i in range(qty + 1):
    #     circle(offset + i*182, vertical_offset)

    ''' Поручень двухбалки двойной '''
    # rectangle_views(n, width=58, left=True, right=True, through=False)
    # pillar_hole(n / 2, 58 + 100 + 44)
    # qty, offset = regular_qty_and_offset(int(n / 2 - 88))
    # offset += 44
    # vertical_offset = 58 + 100 + 44
    # for i in range(qty + 1):
    #     circle(offset + i * 182, vertical_offset)
    #     circle(offset + n / 2 + i * 182, vertical_offset)

    ''' Поручень крайний двухбалки для скошенного свеса длиной overhang_length '''
    # full_length = n + left_offset
    # rectangle_views(full_length, width=58, angle1=-55, right=True, through=False)
    # pillar_hole(left_offset, 58 + 100 + 44)
    # qty, offset = regular_qty_and_offset(n - 88)
    # offset += 44
    # vertical_offset = 58 + 100 + 44
    # for i in range(qty + 1):
    #     circle(left_offset + offset + i*182, vertical_offset)

    ''' Поручень крайний двухбалки для прямого свеса длиной overhang_length '''
    # full_length = n + overhang_length
    # rectangle_views(full_length, width=58, angle1=-45, right=True, through=False)
    # pillar_hole(overhang_length, 58 + 100 + 44)
    # qty, offset = regular_qty_and_offset(n - 88)
    # offset += 44
    # vertical_offset = 58 + 100 + 44
    # for i in range(qty + 1):
    #     circle(overhang_length + offset + i * 182, vertical_offset)

    ''' Поручень крайний двойной для скошенного свеса длиной overhang_length '''
    # full_length = n + left_offset
    # rectangle_views(full_length, width=58, angle1=-55, right=True, through=False)
    # vertical_offset = 58 + 100 + 44
    # pillar_hole(left_offset, vertical_offset)
    # pillar_hole(left_offset + n / 2, vertical_offset)
    # qty, offset = regular_qty_and_offset(int(n / 2 - 88))
    # offset += 44
    # for i in range(qty + 1):
    #     circle(left_offset + offset + i * 182, vertical_offset)
    #     circle(left_offset + offset + n / 2 + i * 182, vertical_offset)

    ''' Поручень крайний двойной для прямого свеса длиной overhang_length '''
    # full_length = n + overhang_length
    # rectangle_views(full_length, width=58, angle1=-55, right=True, through=False)
    # vertical_offset = 58 + 100 + 44
    # pillar_hole(overhang_length, vertical_offset)
    # pillar_hole(overhang_length + n / 2, vertical_offset)
    # qty, offset = regular_qty_and_offset(int(n / 2 - 88))
    # offset += 44
    # for i in range(qty + 1):
    #     circle(overhang_length + offset + i * 182, vertical_offset)
    #     circle(overhang_length + offset + n / 2 + i * 182, vertical_offset)

    ''' Балка двухбалки '''
    # rectangle_views(n, width=58, left=True, right=True)
    # qty, offset = regular_qty_and_offset(n - 88)
    # offset += 44
    # vertical_offset = 58 * 2 + 300 + 88 + 44
    # for i in range(qty + 1):
    #     circle(offset + i*182, vertical_offset)

    ''' Балка двухбалки двойная '''
    # rectangle_views(n, width=58, left=True, right=True)
    # vertical_offset = 58 + 100 + 44
    # pillar_hole(n / 2, vertical_offset)
    # vertical_offset = 58 * 2 + 300 + 88 + 44
    # pillar_hole(n / 2, vertical_offset)
    # qty, offset = regular_qty_and_offset(int(n / 2 - 88))
    # offset += 44
    # for i in range(qty + 1):
    #     circle(offset + i * 182, vertical_offset)
    #     circle(offset + n / 2 + i * 182, vertical_offset)

    ''' Балка крайняя двухбалки для свеса длиной overhang_length '''
    # full_length = n + overhang_length
    # rectangle_views(full_length, width=58, angle1=45, right=True)
    # vertical_offset = 58 * 2 + 300 + 88 + 44
    # pillar_hole(overhang_length, 58 + 100 + 44)
    # pillar_hole(overhang_length, vertical_offset)
    # qty, offset = regular_qty_and_offset(n - 88)
    # offset += 44
    # for i in range(qty + 1):
    #     circle(overhang_length + offset + i * 182, vertical_offset)

    ''' Балка крайняя двухбалки для свеса длиной overhang_length '''
    # full_length = n + overhang_length
    # rectangle_views(full_length, width=58, angle1=45, right=True)
    # vertical_offset = 58 * 2 + 300 + 88 + 44
    # pillar_hole(overhang_length, 58 + 100 + 44)
    # pillar_hole(overhang_length, vertical_offset)
    # pillar_hole(overhang_length + n / 2, 58 + 100 + 44)
    # pillar_hole(overhang_length + n / 2, vertical_offset)
    # qty, offset = regular_qty_and_offset(int(n / 2 - 88))
    # offset += 44
    # for i in range(qty + 1):
    #     circle(overhang_length + offset + i * 182, vertical_offset)
    #     circle(overhang_length + offset + n / 2 + i * 182, vertical_offset)

    file_path = os.path.join(DXF_FOLDER, folder, f'ЦРНС.305112.{CODE1}.{CODE2}.{PART_NUMBER}.{overhang_length}-{"{:0>4}".format(str(n + overhang_length))}.dxf')
    # file_path = os.path.join(DXF_FOLDER, folder, f'ЦРНС.305112.{CODE1}.{CODE2}.{PART_NUMBER}-{"{:0>4}".format(str(n))}.dxf')
    api5_document.ksSaveToDXF(file_path)

    delete_objects()



