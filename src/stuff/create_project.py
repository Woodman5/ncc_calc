import win32com.client as wclient
import pythoncom

import os
from shutil import copy2

from swconst import SwConstants

MODELS_TEMPLATE_FOLDER = os.path.abspath('original_files\\models')
PRJ_FOLDER = os.path.abspath('projects')
MODELS_FOLDER = '3d_models'
DXF_FOLDER = 'dxf'

arg1 = wclient.VARIANT(pythoncom.VT_DISPATCH, None)
arg2 = wclient.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, -1)

copied_files = []
new_model_files = {}


def copyfiles(name, src, prjname):
    copy_to = os.path.join(models_path, f'{"_".join(name.split(".")[:-1])}_{prjname}.{name.split(".")[-1]}')
    copy2(src, copy_to)
    copied_files.append(src)
    print(f'{name} copied')
    return copy_to


while True:
    prj_name = input("Название проекта: ")
    if not os.path.exists(os.path.join(PRJ_FOLDER, prj_name)):
        os.makedirs(os.path.join(PRJ_FOLDER, prj_name, DXF_FOLDER))
        models_path = os.path.join(PRJ_FOLDER, prj_name, MODELS_FOLDER)
        os.makedirs(models_path)

        # connecting to SW
        swYearLastDigit = 4
        sw = wclient.Dispatch(
            "SldWorks.Application.%d" % (20 + (swYearLastDigit - 2)))  # e.g. 20 is SW2012,  23 is SW2015

        sw.CloseAllDocuments(True)

        sw.DocumentVisible(True, 2)  # assembly
        sw.DocumentVisible(False, 1)  # parts

        # processing all assembly files in 3d_models folder
        model_files = list(filter(lambda x: x.lower().endswith('.sldasm'), os.listdir(MODELS_TEMPLATE_FOLDER)))
        for i, v in enumerate(model_files):
            copy_from = os.path.join(MODELS_TEMPLATE_FOLDER, v)
            new_path = copyfiles(v, copy_from, prj_name)

            model_files_from = list(sw.GetDocumentDependencies2(copy_from, True, True, False))[1::2]
            print(model_files_from)

            assembly = sw.OpenDoc6(new_path, 2, 1, "", arg2, arg2)

            for file in model_files_from:
                print('file:', file)  # file: N:\DD\MOSTI\PROGRAM\PO\Pillar Cookie.SLDPRT
                model_name = os.path.basename(file[:file.rindex('.')])
                print('model_name:', model_name)  # model_name: Pillar Cookie

                if file not in copied_files:
                    new_model_files[model_name] = copyfiles(os.path.basename(file), file, prj_name)

                for each in assembly.GetComponents(True):
                    # print('each_name:', each.name2[:each.name2.rindex('-')])
                    if model_name == each.name2[:each.name2.rindex('-')]:
                        comp_name = f'{each.name2}@{os.path.basename(new_path)[:os.path.basename(new_path).rindex(".")]}'
                        print('comp_name:', comp_name)
                        assembly.Extension.SelectByID2(comp_name, "COMPONENT", 0, 0, 0, False, 0, arg1, 0)
                        print(new_model_files)
                        t = assembly.ReplaceComponents(new_model_files[model_name], "", True, True)
                        print(t)
                        assembly.ClearSelection2(True)
                        break

            assembly.EditRebuild3
            assembly.Save3(1, arg2, arg2)

        sw.CloseAllDocuments(True)

        # sw.DocumentVisible(True, 2)
        sw.DocumentVisible(True, 1)

        break
