# -*- coding:  utf-8 -*-
#https://forum.ascon.ru/index.php/topic,32760.0.html
import os

import pythoncom
from win32com.client import Dispatch, gencache


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

# Управление слоями
# s = api5_document.ksLayer(3)
# api5_document.ksLayer(0)
# api5_document.ksDeleteObj(s)

lmin = 300
lmax = 302
width = 40

drw_obj = []

for i in range(lmin, lmax + 1):
    for item in drw_obj:
        api5_document.ksDeleteObj(item)

    drw_obj.clear()

    y = 0
    drw_obj.append(api5_document.ksLineSeg(0, y, i, y, 1))
    drw_obj.append(api5_document.ksLineSeg(i, y, i, width, 1))
    drw_obj.append(api5_document.ksLineSeg(i, width, 0, width, 1))
    drw_obj.append(api5_document.ksLineSeg(0, y, 0, width, 1))

    y = width + 100
    drw_obj.append(api5_document.ksLineSeg(0, y, i, y, 1))
    drw_obj.append(api5_document.ksLineSeg(i, y, i, y + width, 1))
    drw_obj.append(api5_document.ksLineSeg(i, y + width, 0, y + width, 1))
    drw_obj.append(api5_document.ksLineSeg(0, y, 0, y + width, 1))

    y += width + 100
    drw_obj.append(api5_document.ksLineSeg(0, y, i, y, 1))
    drw_obj.append(api5_document.ksLineSeg(i, y, i, y + width, 1))
    drw_obj.append(api5_document.ksLineSeg(i, y + width, 0, y + width, 1))
    drw_obj.append(api5_document.ksLineSeg(0, y, 0, y + width, 1))

    y += width + 100
    drw_obj.append(api5_document.ksLineSeg(0, y, i, y, 1))
    drw_obj.append(api5_document.ksLineSeg(i, y, i, y + width, 1))
    drw_obj.append(api5_document.ksLineSeg(i, y + width, 0, y + width, 1))
    drw_obj.append(api5_document.ksLineSeg(0, y, 0, y + width, 1))

    drw_obj.append(api5_document.ksCircle(i/3, width/2, 16.25, 1))
    drw_obj.append(api5_document.ksCircle(i/3 + 182, width/2, 16.25, 1))

    print(drw_obj)





# print(dir(api7))
# # print(dir(doc2D1))
# print(doc2D1.VariablesCount(0))
# # doc2D1.AddVariable('l', 1, None)
# # doc2D1.AddVariable('k', 55, None)
# # print(doc2D1.VariablesCount(0))
#
# print(doc2D1.Variable(False, 'l'))
# print(doc2D1.Variable(False, 'k'))
#
# for n in range(134, 185):
#     myvars = doc2D1.Variables(0)
#     myvars[0].Expression = n
#
# # myvars[1].Expression = 'l/5'
# #
#     doc2D1.RebuildDocument()
# #
# # myvars = doc2D1.Variables(0)
# # print(myvars[0].Name, myvars[0].Value)
# # print(myvars[1].Value)
#
# # MODELS_FOLDER = os.path.expanduser('~\\Documents\\test_dxf')
# #
# # if not os.path.exists(MODELS_FOLDER):
# #     os.makedirs(MODELS_FOLDER)
#
#     api5_document.ksSaveToDXF(f'c:\\test_dxf\\{n}.dxf')

