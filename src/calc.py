# -*- coding: utf-8 -*-
# !C:\Python\Python37

import sys  # sys нужен для передачи argv в QApplication
import os
import re
import psutil
from shutil import copy2
import pprint
from math import pi

from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMessageBox

import win32com.client as wclient
import pythoncom

# from src.design import designtabs  # Это наш конвертированный файл дизайна
from src.design import design  # Это наш конвертированный файл дизайна
from src.helpers.handrail import Handrail
from src.helpers import config
from src.solidworks.swconst import SwConstants

pp = pprint.PrettyPrinter(width=60, compact=True)
regex = r"[^\d-]+"
MODELS_FOLDER = os.path.expanduser('~\\Documents\\Perila\\PO')
TEMPLATES_FOLDER = os.path.expanduser('~\\Documents\\Perila\\PO\\templates')
DXF_FOLDER = 'dxf'

arg1 = wclient.VARIANT(pythoncom.VT_DISPATCH, None)
arg2 = wclient.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, -1)

copied_files = []
new_model_files = {}


class NccApp(QtWidgets.QMainWindow, design.Ui_MainWindow):
    def __init__(self):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле design.py
        super().__init__()

        self.handrail = Handrail()

        self.setupUi(self)  # Это нужно для инициализации

        self.prjName.editingFinished.connect(lambda: self.heading_process('prj_name', self.prjName.text()))
        self.prjManager.editingFinished.connect(lambda: self.heading_process('prj_manager', self.prjManager.text()))

        self.bridgeLength.valueChanged[int].connect(lambda x: self.validate(setattr(self.handrail, 'blength', x)))
        self.pillarDist.valueChanged[int].connect(lambda x: self.validate(setattr(self.handrail, 'pillar_dist', x)))
        self.lengths.editingFinished.connect(lambda: self.text_process('lenth_list', self.lengths.text()))
        self.gapWdths.editingFinished.connect(lambda: self.text_process('gap_wdth', self.gapWdths.text()))
        self.beforeGapsL.editingFinished.connect(lambda: self.text_process('before_gaps_l', self.beforeGapsL.text()))
        self.beforeGapsR.editingFinished.connect(lambda: self.text_process('before_gaps_r', self.beforeGapsR.text()))
        self.pillarDistList.editingFinished.connect(
            lambda: self.text_process('pillar_dist_list', self.pillarDistList.text()))
        self.gapQty.valueChanged[int].connect(lambda x: self.validate(setattr(self.handrail, 'gap_qty', x)))
        self.beforeGapL.valueChanged[int].connect(lambda x: self.validate(setattr(self.handrail, 'before_gap_l', x)))
        self.beforeGapR.valueChanged[int].connect(lambda x: self.validate(setattr(self.handrail, 'before_gap_r', x)))
        self.overhangEndL.valueChanged[int].connect(
            lambda x: self.validate(setattr(self.handrail, 'overhang_end_l', x)))
        self.overhangEndR.valueChanged[int].connect(
            lambda x: self.validate(setattr(self.handrail, 'overhang_end_r', x)))
        self.useDoubleSections.stateChanged[int].connect(lambda x: self.set_bool(x, 'use_double_sections'))
        self.newFittings.stateChanged[int].connect(lambda x: self.set_bool(x, 'new_fitting'))
        self.beam.currentIndexChanged[int].connect(
            lambda x: self.validate(setattr(self.handrail, 'beem_dist', int(config.beam_dist[x]))))

        self.poConfig.textChanged.connect(self.process_conf_text)

        self.ral1.addItems(config.ral_colors)
        self.ral2.addItems(config.ral_colors)
        self.ral1.setCurrentIndex(-1)
        self.ral2.setCurrentIndex(-1)

        self.beam.addItems(config.beam_dist)
        self.beam.setCurrentIndex(1)

        self.beforeGapL.setValue(self.handrail.before_gap_l)
        self.beforeGapR.setValue(self.handrail.before_gap_r)

        self.serviceMessages.setText('')
        self.resultMessage.setText('')
        self.gapsMarker.setText('')
        self.lengthMarker.setText('')
        self.leftMarker.setText('')
        self.rightMarker.setText('')
        self.pillarMarker.setText('')
        self.resultMarker.setText('')

        # self.beforeGapR_bool.setStyleSheet("background-color: rgb(85, 170, 127);")
        self.lengths_bool.hide()
        self.bridgeLength_bool.hide()
        self.gapQty_bool.hide()
        self.beforeGapL_bool.hide()
        self.beforeGapR_bool.hide()
        self.overhangEndL_bool.hide()
        self.overhangEndR_bool.hide()
        self.pillar_bool.hide()

        self.btnGenerate.setEnabled(False)
        self.btnSave.setEnabled(False)

        self.btnGenerate.clicked.connect(self.create_project)
        self.btnSave.clicked.connect(self.save_bridge)
        self.btnLoad.clicked.connect(self.load_conf)
        self.leftCopy.clicked.connect(lambda: self.data_copy('before_gap_l'))
        self.rightCopy.clicked.connect(lambda: self.data_copy('before_gap_r'))
        self.pillarCopy.clicked.connect(lambda: self.data_copy('pillar_dist'))

        self.data = {
            'gap_wdth': (self.label_8.text(), 'gapWdths', 'gapsMarker'),
            'lenth_list': (self.label_13.text(), 'lengths', 'lengthMarker'),
            'before_gaps_l': (self.label_9.text(), 'beforeGapsL', 'leftMarker'),
            'before_gaps_r': (self.label_10.text(), 'beforeGapsR', 'rightMarker'),
            'before_gap_l': (self.label_9.text(), 'beforeGapL', 'before_gaps_l'),
            'before_gap_r': (self.label_10.text(), 'beforeGapR', 'before_gaps_r'),
            'prj_name': (self.label.text(), 'prjName'),
            'prj_manager': (self.label_2.text(), 'prjManager'),
            'gap_qty': (self.label_7.text(), 'gapQty'),
            'overhang_end_l': (self.label_11.text(), 'overhangEndL'),
            'overhang_end_r': (self.label_12.text(), 'overhangEndR'),
            'blength': (self.label_4.text(), 'bridgeLength'),
            'pillar_dist': (self.label_14.text(), 'pillarDist', 'pillar_dist_list'),
            'pillar_dist_list': (self.label_14.text(), 'pillarDistList', 'pillarMarker'),
        }

        self.handrail.par_folder = os.path.expanduser('~\\Documents\\Perila')
        self.folderLabel.setText(f'Папка сохранения проектов: {self.handrail.par_folder}')

    def set_serv_message(self, text=None):
        self.serviceMessages.clear()
        if text:
            self.serviceMessages.setStyleSheet("color: red;")
            self.serviceMessages.setText(text)
            self.serviceMessages.adjustSize()

    def set_info_message(self, text=None, alert=False):
        self.folderLabel.clear()
        self.folderLabel.setStyleSheet('')
        if alert:
            self.folderLabel.setStyleSheet("color: red;")
        if text:
            self.folderLabel.setText(text)
            self.folderLabel.adjustSize()

    def set_text(self, name, key=True):
        fieldtext = '  '.join(str(x) for x in getattr(self.handrail, name))
        fieldlabel = '1' if key else ''
        count = 2
        key = 0
        while True:
            key = fieldtext.find(' ', key)
            if key == -1:
                break
            fieldlabel = fieldlabel.ljust(key + 2, ' ')
            fieldlabel += str(count)
            count += 1
            key += 2
        curwidth = self.__dict__[self.data[name][1]].width()
        needwidth = len(fieldtext) * 7 + 30
        self.__dict__[self.data[name][1]].setText(fieldtext)
        if curwidth < needwidth - 30:
            self.__dict__[self.data[name][1]].setFixedWidth(len(fieldtext) * 7 + 30)
        self.__dict__[self.data[name][2]].setText(fieldlabel)
        self.__dict__[self.data[name][2]].adjustSize()

    def save_bridge(self, path=None):
        if not (self.handrail.prj_name and self.handrail.prj_manager):
            self.set_info_message('Укажите название проекта и имя менеджера', True)
            return
        saved = self.handrail.save_conf(path)
        if saved:
            self.set_info_message('Файл сохранен', True)
        return True

    def data_copy(self, name):
        value = getattr(self.handrail, name)
        j = self.handrail.gap_qty
        if name == 'pillar_dist':
            j += 1
        if isinstance(value, int):
            value = [value] * j
        setattr(self.handrail, self.data[name][2], value)
        self.set_text(self.data[name][2])
        self.validate()

    def heading_process(self, name, text):
        setattr(self.handrail, name, text)
        self.handrail.processConfig()
        self.btnSave.setEnabled(True)

    def show_result(self):
        text, label, stats = self.handrail.__str__()
        self.resultMessage.setText(text)
        self.resultMarker.setText(label)
        self.resultMarker.adjustSize()
        self.resultMessage.adjustSize()
        self.serviceMessages.clear()
        self.serviceMessages.setStyleSheet("")
        self.serviceMessages.setText(stats)
        self.serviceMessages.adjustSize()
        self.btnGenerate.setEnabled(True)
        self.btnSave.setEnabled(True)

    def validate_po_conf(self, data):
        pass

    def text_process(self, name, text):
        if text == '':
            setattr(self.handrail, name, [])
            self.set_text(name, False)
            return False
        text = (re.sub(regex, ' ', text)).strip()
        if text:
            ln_list = [int(x) for x in text.split(' ') if int(x) > 0]
            qty = self.handrail.gap_qty
            if not ln_list:
                self.set_serv_message('- Отрицательные числа недопустимы. '
                                      'Пример правильного написания: "450 687 127".\n')
            elif len(text.split(' ')) == 1 and len(ln_list) == 1:
                if name in ['lenth_list', 'pillar_dist_list']:
                    ln_list = [ln_list[0]] * (qty + 1)
                else:
                    ln_list = [ln_list[0]] * qty
                setattr(self.handrail, name, ln_list)
                self.set_serv_message()
                self.validate()
            elif len(ln_list) != qty and name not in ['lenth_list', 'pillar_dist_list'] or len(
                    ln_list) != qty + 1 and name in ['lenth_list', 'pillar_dist_list']:
                self.set_serv_message(f'- Количество введеных размеров в поле "{self.data[name][0]}" не соответствует '
                                      f'количеству швов.\n')
            elif len(ln_list) == qty and name not in ['lenth_list', 'pillar_dist_list'] or len(
                    ln_list) == qty + 1 and name in ['lenth_list', 'pillar_dist_list']:
                self.set_serv_message()
                setattr(self.handrail, name, ln_list)
                self.validate()
            self.set_text(name)
        else:
            self.set_serv_message('- Размеры указаны неверно. Пример правильного написания: "450 687 127".\n')

    def clearResults(self):
        self.resultMessage.clear()
        self.resultMarker.clear()
        self.set_info_message()
        self.btnGenerate.setEnabled(False)
        self.btnSave.setEnabled(False)

    def validate(self, x=None):
        self.clearResults()
        if not self.handrail.manual:
            message = ''
            hr = self.handrail
            # проверяем длину моста
            calc_length = sum([sum(hr.lenth_list), sum(hr.gap_wdth)])
            bridge_length = calc_length == hr.blength
            self.bridgeLength_bool.hide()
            if not bridge_length:
                self.bridgeLength_bool.setStyleSheet("background-color: red;")
                self.bridgeLength_bool.show()
                message += f'- Длина моста не равна сумме составляющих ({calc_length} мм) ' \
                           f'на {max(calc_length, hr.blength) - min(calc_length, hr.blength)} мм.\n'

            # проверяем свесы, ограничения есть в форме, но на всякий случай оставлю это тут
            self.overhangEndR_bool.hide()
            self.overhangEndL_bool.hide()
            overhang_r = 220 <= hr.overhang_end_r <= 434
            overhang_l = 220 <= hr.overhang_end_l <= 434
            if not overhang_l:
                self.overhangEndL_bool.setStyleSheet("background-color: red;")
                self.overhangEndL_bool.show()
                message += '- Левый свес должен быть от 220 до 434 мм.\n'
            if not overhang_r:
                self.overhangEndR_bool.setStyleSheet("background-color: red;")
                self.overhangEndR_bool.show()
                message += '- Правый свес должен быть от 220 до 434 мм.\n'

            # проверяем отступы
            self.beforeGapL_bool.hide()
            self.beforeGapR_bool.hide()
            if hr.before_gaps_l:
                for key in hr.before_gaps_l:
                    gap_l = 200 <= key <= 1000
                    if not gap_l:
                        break
            else:
                gap_l = 200 <= hr.before_gap_l <= 1000
            if hr.before_gaps_r:
                for key in hr.before_gaps_r:
                    gap_r = 200 <= key <= 1000
                    if not gap_r:
                        break
            else:
                gap_r = 200 <= hr.before_gap_r <= 1000

            if not gap_r:
                self.beforeGapR_bool.setStyleSheet("background-color: red;")
                self.beforeGapR_bool.show()
                message += '- Правый отступ должен быть от 200 до 1000 мм.\n'
            if not gap_l:
                self.beforeGapL_bool.setStyleSheet("background-color: red;")
                self.beforeGapL_bool.show()
                message += '- Левый отступ должен быть от 200 до 1000 мм.\n'

            # проверяем ширину дефшвов
            gap_bool = True
            self.gapQty_bool.hide()
            if len(hr.gap_wdth) != hr.gap_qty:

                self.gapQty_bool.setStyleSheet("background-color: red;")
                self.gapQty_bool.show()
                gap_bool = False
                message += '- Количество швов не совпадает с количеством в перечне их ширины.\n'
            elif hr.gap_qty > 0:
                for value in hr.gap_wdth:
                    gap_bool = 10 <= value <= 500
                    if not gap_bool:
                        self.gapQty_bool.setStyleSheet("background-color: red;")
                        self.gapQty_bool.show()
                        message += '- Ширина деформационного шва должна быть от 10 до 500 мм.\n'
                        break

            # проверяем отступы и швы вместе
            self.gapWdths.setStyleSheet("")
            gapl_qty_bool = True
            gapr_qty_bool = True
            gap_sum_bool = True
            if hr.gap_qty > 0 and gap_bool:
                if not self.beforeGapsL.text():
                    hr.before_gaps_l.clear()
                if not self.beforeGapsR.text():
                    hr.before_gaps_r.clear()
                if hr.before_gaps_r:
                    bef_r = hr.before_gaps_r
                else:
                    bef_r = [hr.before_gap_r] * hr.gap_qty
                if hr.before_gaps_l:
                    bef_l = hr.before_gaps_l
                else:
                    bef_l = [hr.before_gap_l] * hr.gap_qty
                if len(bef_l) != hr.gap_qty:
                    self.beforeGapL_bool.setStyleSheet("background-color: red;")
                    self.beforeGapL_bool.show()
                    gapl_qty_bool = False
                    message += '- Количество левых отступов не соответствует количеству швов.\n'
                elif len(bef_r) != hr.gap_qty:
                    self.beforeGapR_bool.setStyleSheet("background-color: red;")
                    self.beforeGapR_bool.show()
                    gapr_qty_bool = False
                    message += '- Количество правых отступов не соответствует количеству швов.\n'
                else:
                    for key, value in enumerate(hr.gap_wdth):
                        temp = value + bef_l[key] + bef_r[key]
                        gap_sum_bool = 410 <= temp <= 1500
                        if not gap_sum_bool:
                            self.beforeGapL_bool.setStyleSheet("background-color: red;")
                            self.beforeGapR_bool.setStyleSheet("background-color: red;")
                            self.gapWdths.setStyleSheet("color: red;")
                            self.beforeGapL_bool.show()
                            self.beforeGapR_bool.show()
                            message += f'- Сумма отступов и шва №{key + 1} должна быть не более 1.5 м, сейчас {temp} мм.\n'

            # проверяем расстояния между стойками
            self.pillar_bool.hide()
            pillar_dist_bool = True
            if not self.pillarDistList.text():
                hr.pillar_dist_list.clear()

            if hr.pillar_dist_list:
                pill = hr.pillar_dist_list
            else:
                pill = [hr.pillar_dist] * (hr.gap_qty + 1)

            if len(pill) == hr.gap_qty + 1:
                for key, value in enumerate(hr.pillar_dist_list):
                    pillar_dist_bool = 500 <= value <= 2000
                    if not pillar_dist_bool:
                        self.pillar_bool.setStyleSheet("background-color: red;")
                        self.pillar_bool.show()
                        message += f'- Расстояние между опорами на участке {key + 1} должно быть между 500 и 2000 мм.\n'
                        break
            else:
                self.pillar_bool.setStyleSheet("background-color: red;")
                self.pillar_bool.show()
                pillar_dist_bool = False
                message += '- Количество расстояний между опор не соответствует количеству швов.\n'

            # проверяем длины участков моста
            self.lengths_bool.hide()
            if len(hr.lenth_list) == hr.gap_qty + 1:
                if hr.gap_qty == 0:
                    bridge_length_bool = hr.lenth_list[0] - hr.overhang_end_r - hr.overhang_end_l >= 500
                    if not bridge_length_bool:
                        self.lengths_bool.setStyleSheet("background-color: red;")
                        self.lengths_bool.show()
                        message += f'- Длина одиночного участка перил не может быть ' \
                                   f'короче {hr.overhang_end_r + hr.overhang_end_l + 500} мм.\n'
                elif hr.gap_qty > 0:
                    overh_r = [hr.overhang_end_r]
                    overh_r.extend([0] * hr.gap_qty)
                    overh_l = [0] * hr.gap_qty
                    overh_l.extend([hr.overhang_end_l])
                    if hr.before_gaps_r:
                        bef_r = hr.before_gaps_r.copy()
                        bef_r.insert(0, 0)
                    else:
                        bef_r = [0]
                        bef_r.extend([hr.before_gap_r] * hr.gap_qty)
                    if hr.before_gaps_l:
                        bef_l = hr.before_gaps_l.copy()
                        bef_l.append(0)
                    else:
                        bef_l = [hr.before_gap_l] * hr.gap_qty
                        bef_l.append(0)
                    for key, value in enumerate(hr.lenth_list):
                        temp = overh_l[key] + overh_r[key] + bef_l[key] + bef_r[key]
                        if not key or key == len(hr.lenth_list) - 1:
                            temp += 500
                        bridge_length_bool = value >= temp
                        if not bridge_length_bool:
                            self.lengths_bool.setStyleSheet("background-color: red;")
                            self.lengths_bool.show()
                            message += f'- Длина участка {key + 1} должна быть не менее {temp} мм.\n'
                            break
            else:
                self.lengths_bool.setStyleSheet("background-color: red;")
                self.lengths_bool.show()
                bridge_length_bool = False
                message += '- Количество участков моста не соответствует количеству швов и не может быть меньше 1.\n'

            self.set_serv_message(message)

            if bridge_length and overhang_r and overhang_l and gap_l and gap_r and gap_bool and pillar_dist_bool \
                    and bridge_length_bool and gapl_qty_bool and gapr_qty_bool and gap_sum_bool:
                hr.calculate()
                self.show_result()
        else:
            self.process_conf_text()

    def set_bool(self, value, name):
        value = True if value > 0 else False
        setattr(self.handrail, name, value)
        self.validate()

    def load_conf(self):
        var_list = config.saved_parameters

        fname = QtWidgets.QFileDialog.getOpenFileName(
            self,
            'Открыть файл',
            self.handrail.par_folder,
            "Text Files (*.txt)"
        )
        file_path = fname[0]

        params = []
        bridge_config = []
        counter = len(var_list) - 1

        if file_path:
            with open(file_path, 'r', encoding='utf-8') as fin:
                for line in fin:
                    line = line.strip()
                    if line and not line.startswith('#'):
                        if counter > 0:
                            params.append(line)
                            counter -= 1
                        else:
                            bridge_config.append(line)
        print(params)
        print(bridge_config)
        self.poConfig.setPlainText('')
        self.handrail.manual_config = []

        try:

            for key, item in enumerate(params):
                if key in [0, 1, 10, 11]:
                    self.__dict__[var_list[key][1]].setText(item)
                    if key in [10, 11]:
                        self.text_process(var_list[key][0], item)
                    else:
                        setattr(self.handrail, var_list[key][0], item)
                if key in [8, 9]:
                    if item in ['yes', 'y', 'Yes', 'YES', '+', 'д', 'да', 'Да', 'ДА']:
                        status = True
                    elif item in ['no', 'n', 'No', 'NO', '-', 'н', 'нет', 'НЕТ', 'Нет']:
                        status = False
                    self.__dict__[var_list[key][1]].setChecked(status)
                    setattr(self.handrail, var_list[key][0], status)
                if key in [2, 3, 5, 6]:
                    self.__dict__[var_list[key][1]].setValue(int(item))
                    setattr(self.handrail, var_list[key][0], int(item))
                if key == 4:
                    self.__dict__[var_list[key][1]].setCurrentText(item)
                    setattr(self.handrail, var_list[key][0], int(item))
                if key in [7, 12, 13]:
                    data = item.split()
                    if len(data) > 1 and len(set(data)) > 1:
                        self.__dict__[var_list[key][1]].setText(item)
                        self.text_process(var_list[key][0], item)
                    elif len(data) == 1 or len(set(data)) == 1:
                        self.__dict__[var_list[key][2]].setValue(int(data[0]))
                        setattr(self.handrail, var_list[key][3], int(data[0]))

            if bridge_config:
                self.poConfig.setPlainText('\n\n'.join(bridge_config))
                self.handrail.manual_config = bridge_config
                self.validate_po_conf(bridge_config)

            self.validate()

        except Exception:
            self.set_serv_message('Файл составлен неправильно или поврежден')

    def process_conf_text(self):
        self.clearResults()
        self.set_serv_message()
        allowed_chars = ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '!', '_', 'x', ' ', '(', ')', '\n')
        text = self.poConfig.toPlainText()
        message = ''

        def x_pattern(pattern: str) -> list:
            data = [int(x) for x in pattern.split('x')]
            return [data[1]] * data[0]

        def bracket_pattern(br_pattern: str) -> list:
            temp = br_pattern.split('x')
            data = temp[1][1:-1].split(' ')
            prepared_data = []
            for value in data:
                if value.find('x') > -1:
                    prepared_data.extend(x_pattern(value))
                else:
                    prepared_data.append(int(value))
            return prepared_data * int(temp[0])

        def process_brackets(arg: list, count: int) -> list:
            key1, key2 = -1, -1
            while count > 0:
                for i, j in enumerate(arg):
                    if j.find('(') > -1:
                        key1 = i
                        continue
                    elif j.find(')') > -1:
                        key2 = i
                        break
                if key1 == key2 and key1 != -1:
                    arg[key1] = arg[key1][:-1].replace('(', '')
                elif key1 != -1 and key2 > key1:
                    for x in range(key1 + 1, key2 + 1):
                        arg[key1] = arg[key1] + ' ' + arg[x]
                    for x in range(key1 + 1, key2 + 1):
                        arg.pop(x)
                count -= 1
            return arg

        def make_list(arg: text) -> list:
            splited_data = arg.strip(' ').split(' ')
            while splited_data.count(''):
                splited_data.remove('')
            return splited_data

        if len(text) > 0:
            self.handrail.manual = True
            for char in text:
                if char not in allowed_chars:
                    message = 'Конфигурация содержит ошибки.\n'
                    break

            x_count = text.count('x')
            lb_count = text.count('(')
            rb_count = text.count(')')
            _count = text.count('_')

            if _count:
                self.gapQty.setValue(_count)
                self.handrail.gap_qty = _count

            if (_count and not self.handrail.gap_wdth) or _count > len(self.handrail.gap_wdth):
                message = message + 'Размеры деформационных швов не заданы.\n'

            if text[-1] == '_' or text[0] == '_':
                message = message + 'Символ "_" не может быть первым или последним символом.\n'

            if lb_count != rb_count:
                message = message + 'Количество открывающих и закрывающих скобок не совпадает.\n'
            else:
                k, k2, lrb_list = 0, 0, []
                for _ in range(lb_count):
                    index = text.find('(', k)
                    if index == len(text) - 1:
                        message = message + 'Открывающая скобка не может быть последней.\n'
                        break
                    if text[index + 1] == ' ':
                        message = message + 'После открывающей скобки не может следовать пробел.\n'
                        break
                    lrb_list.append(index)
                    index2 = text.find(')', k2)
                    lrb_list.append(index2)
                    k = index + 1
                    k2 = index2 + 1
                for i in range(len(lrb_list) - 1):
                    if lrb_list[i] > lrb_list[i + 1]:
                        message = message + 'Порядок следования скобок неверен. Вложенность не допускатся.\n'
                        break
                for i in range(0, len(lrb_list) - 1, 2):
                    if lrb_list[i + 1] - lrb_list[i] == 1:
                        message = message + 'Между скобками должны быть числа.\n'
                        break
                    for char in text[lrb_list[i] + 1:lrb_list[i + 1]]:
                        if char in ['!', '_', 'x', '(', ')', '\n']:
                            message = message + 'Между скобками должны быть числа.\n'
                            break

            k = 0
            for i in range(x_count):
                x = text.find('x', k)
                k = x + 1
                if x == 0:
                    message = message + 'Символ "х" должен предваряться целым числом.\n'
                    break
                elif x == len(text) - 1:
                    message = message + 'Символ "х" не может быть последним символом.\n'
                    break
                elif text[x - 1] in ['!', '_', 'x', ' ', '(', ')', '\n']:
                    message = message + 'Символ "х" должен предваряться целым числом.\n'
                    break
                elif text[x + 1] in ['!', '_', 'x', ' ', ')', '\n']:
                    message = message + 'После "х" должно быть целое число или открывающая скобка.\n'
                    break

            k = 0
            for i in range(_count):
                x = text.find('_', k)
                k = x + 1
                if text[x - 1] != ' ':
                    message = message + 'Символ "_" должен предваряться пробелом.\n'
                    break
                elif text[-1] != '_' and text[x + 1] != ' ':
                    message = message + 'После "_" должен быть пробел.\n'
                    break

            if text[-1] not in allowed_chars:
                message = message + 'Последний символ не допустим.\n'

            if message:
                self.set_serv_message(message)

            if not message and text[-1] in [' ', '\n']:
                if text.find('\n') > -1:
                    temp_list = text.split('\n')
                    while temp_list.count(''):
                        temp_list.remove('')
                    self.handrail.manual_config = temp_list
                    length_list = []
                    for item in temp_list:
                        temp_data = make_list(item)
                        length_list.append(process_brackets(temp_data, item.count('(')))
                else:
                    self.handrail.manual_config = [text]
                    temp_data = make_list(text)
                    length_list = [process_brackets(temp_data, text.count('('))]

                for key, value in enumerate(length_list):
                    temp_value = []
                    if value[-1] == '_':
                        self.set_serv_message('Строка не может заканчиваться символом "_".\n')
                        return
                    for dist in value:
                        if dist == '_' or dist.startswith('!'):
                            temp_value.append(dist)
                        elif dist.find('(') > -1:
                            temp_value.extend(bracket_pattern(dist))
                        elif dist.find('x') > -1:
                            temp_value.extend(x_pattern(dist))
                        else:
                            temp_value.append(int(dist))
                    length_list[key] = temp_value

                span_qty = len(length_list)

                _, _, bef_l, bef_r = self.handrail.make_gaps()
                if span_qty % 2 == 0:
                    self.set_serv_message('Количество строк конфигурации должно быть нечетным.\n')
                    return

                over_l = [self.handrail.overhang_end_l] * (span_qty // 2 + 1)
                over_r = [self.handrail.overhang_end_r] * (span_qty // 2 + 1)

                for i in range(1, span_qty, 2):
                    if len(length_list[i]) != 3:
                        self.set_serv_message('Четные строки должны состоять из 3 чисел.\n')
                        return
                    elif not 220 <= length_list[i][1] <= 434 or not 220 <= length_list[i][2] <= 434:
                        self.set_serv_message('Длина свеса должна быть от 200 до 434 мм.\n')
                        return
                    if length_list[i][1] + length_list[i][2] + 10 > length_list[i][0]:
                        self.set_serv_message('Сумма свесов перил не может быть больше '
                                              'расстояния между пролетами перил.\n')
                        return

                    over_l[i // 2 + 1] = length_list[i][2]
                    over_r[i // 2] = length_list[i][1]

                bridge = []
                gap_counter = 0

                for num in range(0, span_qty, 2):
                    next_dist = length_list[num + 1][0] if num < span_qty - 1 else 0
                    for key, dist in enumerate(length_list[num]):
                        if dist == '_':
                            if bridge[-1][0] == 'PO Reg.SLDASM':
                                if bridge[-1][1] == 1:
                                    bridge[-1][0] = 'PO Irreg R.SLDASM'
                                else:
                                    bridge[-1][1] -= 1
                                    bridge.append(
                                        ['PO Irreg R.SLDASM', 1, length_list[num][key - 1], 0, 0, 0, 0, 0, 0, 0, 0]
                                    )
                            elif bridge[-1][0] == 'PO End L.SLDASM':
                                bridge[-1][0] = 'PO End L Gap.SLDASM'

                            bridge[-1][7] = self.handrail.gap_wdth[gap_counter]
                            bridge[-1][3] = bef_l[gap_counter]
                            bridge[-1][8] = bef_r[gap_counter + 1]
                            gap_counter += 1

                        elif isinstance(dist, int):
                            if dist < 500 or dist > 2000:
                                self.set_serv_message('Расстояние между опорами на участке '
                                                      'должно быть между 500 и 2000 мм.\n')
                                return
                            if key == 0:
                                if len(length_list[num]) == 1:
                                    bridge.append(
                                        ['PO End LR.SLDASM', 1, dist, 0, 0, over_l[num // 2], over_r[num // 2], 0, next_dist, 0, 0]
                                    )
                                    continue
                                bridge.append(
                                    ['PO End L.SLDASM', 1, dist, 0, 0, over_l[num // 2], 0, 0, 0, 0, 0]
                                )
                            elif key == len(length_list[num]) - 1 and key > 0:
                                if length_list[num][key - 1] == '_':
                                    bridge.append(
                                        ['PO End R Gap.SLDASM', 1, dist, 0, bef_r[gap_counter], 0, over_r[num // 2], 0, next_dist, 0, 0]
                                    )
                                    continue
                                bridge.append(
                                    ['PO End R.SLDASM', 1, dist, 0, 0, 0, over_r[num // 2], 0, next_dist, 0, 0]
                                )
                            else:
                                if length_list[num][key - 1] == '_' and key > 1:
                                    bridge.append(
                                        ['PO Irreg L.SLDASM', 1, dist, 0, bef_r[gap_counter], 0, 0, 0, 0, 0, 0]
                                    )
                                    continue
                                if dist == bridge[-1][2] and bridge[-1][0] == 'PO Reg.SLDASM':
                                    bridge[-1][1] += 1
                                    continue
                                bridge.append(
                                    ['PO Reg.SLDASM', 1, dist, 0, 0, 0, 0, 0, 0, 0, 0]
                                )

                        else:  # здесь обработка калиток "!"
                            pass

                temp_len_list = [0]
                for item in bridge:
                    temp_len_list[-1] += (item[2] * item[1]) + item[3] + item[4] + item[5] + item[6] + item[8]
                    if item[7] != 0:
                        temp_len_list[-1] -= item[8]
                        temp_len_list.append(0)

                pp.pprint(temp_len_list)

                self.handrail.lenth_list = temp_len_list
                self.set_text('lenth_list')

                self.handrail.blength = sum(temp_len_list) + sum(self.handrail.gap_wdth)
                self.bridgeLength.setValue(self.handrail.blength)

                self.handrail.bridge = bridge
                pp.pprint(self.handrail.bridge)
                self.show_result()

        else:
            self.handrail.manual = False

    def sw_connect(self):
        # connecting to SolidWorks
        sldw = False
        for proc in psutil.process_iter(attrs=['name']):
            if proc.info['name'] == 'SLDWORKS.exe':
                sldw = True
        if sldw:
            key = QMessageBox.question(self, 'Внимание!', "Открытые документы Solidworks будут закрыты без сохранения.",
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if key == 65536:
                return None
        swYearLastDigit = 4
        sw = wclient.Dispatch(
            "SldWorks.Application.%d" % (20 + (swYearLastDigit - 2)))  # e.g. 20 is SW2012,  23 is SW2015
        sw.Visible = 1
        sw.CloseAllDocuments(True)

        sw.SetUserPreferenceToggle(SwConstants.swAutomaticScaling3ViewDrawings, False)

        return sw

    def copyfiles(self, models_path, *args):
        count = args[1]
        args = args[0]
        name = f'{"_".join(args[0].split(".")[:-1])}_{self.handrail.prj_name}'
        if args[0] not in config.common_parts:
            if args[0].endswith('SLDASM'):
                name = f'PO-{count}_{self.handrail.prj_name}_l{args[2]}'
            else:
                name = f'PO-{count}_{name}_l{args[2]}'
        copy_to = os.path.join(models_path, f'{name}.{args[0].split(".")[-1]}')
        src = os.path.join(MODELS_FOLDER, args[0])
        if copy_to not in copied_files:
            copy2(src, copy_to)
            copied_files.append(copy_to)
        return src, copy_to

    def make_blocks(self, sw, lens, path, fitting):
        lens = list(set(lens))
        blocks = {}
        part_path = os.path.join(MODELS_FOLDER, 'Bridge span.sldprt')
        part = sw.OpenDoc6(part_path, 1, 1, "", arg2, arg2)
        part = sw.ActiveDoc
        eqMgr = part.GetEquationMgr
        for length in lens:
            name = f'Bridge span {length}.sldprt'
            eqMgr.Equation(0, f'"l" = {length}')
            eqMgr.Equation(4, f'"newFitting" = {fitting}')
            eqMgr.EvaluateAll
            part.EditRebuild3
            new_path = os.path.join(path, name)
            part.SaveAs3(new_path, 0, 0)
            value_y = eqMgr.Value(1)
            value_x = eqMgr.Value(2)
            blocks[length] = (new_path, value_x, value_y)
        sw.CloseDoc(part.GetTitle)
        return blocks

    def to_dxf(self, sw, *args, long_overhand=False, part=False, view=config.views):
        fname = os.path.splitext(os.path.basename(args[0]))[0].split('_')
        with open('log.txt', 'a') as f:
            f.write(f'\n\nargs = {args}\n')
            f.write(f'old-fname = {fname}\n')
        part_bool = long_overhand and part
        if part:
            origin_name = fname[1]
        if part and not long_overhand and origin_name not in config.to_dxf_parts:
            return
        elif part_bool and origin_name not in config.to_dxf_ends:
            return
        dxf_temp = os.path.join(TEMPLATES_FOLDER, 'dxf.drwdot')
        drw = sw.NewDocument(dxf_temp, 12, config.drv_dim[0], config.drv_dim[1])  # create drawing
        model = sw.ActiveDoc
        with open('log.txt', 'a') as f:
            f.write(f'drw = {drw}\n')
            f.write(f'model = {model}\n')
        folder = args[1]
        if part:
            folder = os.path.join(args[1], DXF_FOLDER)
            if fname[1] in config.to_dxf_parts:
                fname[1] = config.to_dxf_parts[fname[1]]
            else:
                fname[1] = config.to_dxf_ends[fname[1]]
            fname.pop()
        fname = '_'.join(fname)
        dxf_path = os.path.join(folder, f'{fname}.dxf')
        with open('log.txt', 'a') as f:
            f.write(f'fname = {fname}\n')
            f.write(f'folder = {folder}\n')
            f.write(f'part-fname = {fname}\n')
            f.write(f'final-fname = {fname}\n')
            f.write(f'dxf_path = {dxf_path}\n')
            f.write(f'view = {view}\n')
        for v in view:
            drw_view = model.CreateDrawViewFromModelView3(args[0], v[0], v[1], v[2], 0)
            with open('log.txt', 'a') as f:
                f.write(f'v = {v}\n')
                f.write(f'drw_view = {drw_view}\n')
            drw_view.SetDisplayTangentEdges2(0)
            if part:
                drw_view.Angle = pi / 2
        model.SaveAs3(dxf_path, 0, 0)
        sw.CloseDoc(drw.GetTitle)
        return True

    def create_project(self):
        prj_name = self.handrail.prj_name
        man_name = self.handrail.prj_manager
        basepath = self.handrail.par_folder
        if not (prj_name and man_name):
            self.set_info_message('Укажите название проекта и имя менеджера', True)
            return
        else:
            self.set_info_message(f'Папка сохранения проектов: {basepath}')

        prj_path = os.path.join(basepath, f'{prj_name}_{man_name}')
        dxf_path = os.path.join(prj_path, DXF_FOLDER)
        models_path = os.path.join(prj_path, '3d_models')

        self.set_info_message('Подключаюсь к Solidworks...')

        sw = self.sw_connect()
        if not sw:
            self.set_info_message(f'Папка сохранения проектов: {basepath}')
            return

        prj_path_bool = os.path.exists(prj_path)
        dxf_path_bool = os.path.exists(dxf_path)
        models_path_bool = os.path.exists(models_path)
        if not prj_path_bool:
            os.makedirs(dxf_path)
            os.makedirs(models_path)
            self.save_bridge(prj_path)
        else:
            key = QMessageBox.question(self, 'Внимание!',
                                       "Папки проекта уже существуют. Будут удалены все файлы, кроме текстовых.",
                                       QMessageBox.Yes | QMessageBox.No,
                                       QMessageBox.Yes)
            if key == 16384:
                folders = [dxf_path, models_path]
                if not dxf_path_bool:
                    os.makedirs(dxf_path)
                    folders.remove(dxf_path)
                if not models_path_bool:
                    os.makedirs(models_path)
                    folders.remove(models_path)
                for folder in folders:
                    files = os.listdir(folder)
                    for file in files:
                        if not file.endswith('.txt'):
                            path = os.path.join(folder, file)
                            if os.path.isfile(path):
                                os.remove(path)
            else:
                QMessageBox.information(self, ' ', "Измените название проекта", QMessageBox.Ok)
                return

        self.set_info_message('Копирую файлы моделей, перестраиваю сборки...')

        #  Убираем дубли секций чтобы не копировать и не создавать одни и те же модели
        br_clear = set()
        for item in self.handrail.bridge:
            temp = list(item[:9])
            temp[1] = 0
            br_clear.add(tuple(temp))

        parts = {}
        indexes = []
        counter = 1
        for section in br_clear:
            copy_from, new_path = self.copyfiles(models_path, section, counter)

            model_files_from = list(sw.GetDocumentDependencies2(copy_from, True, True, False))[1::2]

            for i, x in enumerate(self.handrail.bridge):
                temp = list(x[:9])
                temp[1] = 0
                if section == tuple(temp):
                    indexes.append(i)
                    current_sector = self.handrail.bridge[i]

            for index in indexes:
                parts[index] = [new_path]

            assembly = sw.OpenDoc6(new_path, 2, 1, "", arg2, arg2)
            assembly_name = os.path.basename(new_path)
            active_assembly = sw.ActiveDoc
            active_assembly.ViewZoomtofit2()

            for_dxf = []

            for file in model_files_from:
                model_name = os.path.basename(file[:file.rindex('.')])
                model_name += f'.{file.split(".")[-1].lower()}'

                _, new_model_files[model_name] = self.copyfiles(models_path, [model_name, *section[1:]], counter)

                for each in assembly.GetComponents(True):
                    short_name = model_name[:model_name.rindex('.')]
                    if short_name == each.name2[:each.name2.rindex('-')]:
                        comp_name = f'{each.name2}@{assembly_name[:assembly_name.rindex(".")]}'
                        assembly.Extension.SelectByID2(comp_name, "COMPONENT", 0, 0, 0, False, 0, arg1, 0)
                        t = assembly.ReplaceComponents(new_model_files[model_name], "", True, True)
                        assembly.ClearSelection2(True)
                        for_dxf.append(new_model_files[model_name])
                        break

            done = []
            for part in assembly.GetComponents(True):
                name = part.name2[:part.name2.rindex('-')]
                if name not in done:
                    status = part.GetSuppression
                    if status != 0:
                        model = part.GetModelDoc2
                        eqpart = model.GetEquationMgr
                        done.append(name)
                    if status and eqpart.GetCount > 0:
                        eqDict = {}
                        for key in range(eqpart.GetCount):
                            eqDict[eqpart.Equation(key).split('=')[0].strip('"')] = key
                        for ind, var in eqDict.items():
                            if ind in config.sw_variables_part:
                                value = eqpart.Equation(var).split('=')[-1].split('@')[0]
                                value += f'@{assembly_name}"'
                                bef = eqpart.Equation(var)
                                eqpart.Equation(var, f'"{ind}" = {value}')
                                aft = eqpart.Equation(var)
                        eqpart.EvaluateAll
            done.clear()

            eqMgr = assembly.GetEquationMgr
            eqDict = {}
            for key in range(eqMgr.GetCount):
                eqDict[eqMgr.Equation(key).split('=')[0].strip('"')] = key
            for ind, var in eqDict.items():
                if ind == 'beamToBeamDist':
                    eqMgr.Equation(var, f'"{ind}" = {self.handrail.beem_dist}')
                elif ind == 'newFitting':
                    value = 1 if self.handrail.new_fitting else 0
                    eqMgr.Equation(var, f'"{ind}" = {value}')
                elif ind in config.sw_variables_assembly:
                    value = current_sector[config.sw_variables_assembly[ind]]
                    if current_sector[9] and ind == 'l':
                        value /= 2
                    eqMgr.Equation(var, f'"{ind}" = {value}')
            eqMgr.EvaluateAll
            assembly.EditRebuild3
            x_offset = eqMgr.Value(eqDict['assemXOffset'])
            for index in indexes:
                parts[index].append(x_offset)
            assembly.LightweightAllResolved
            assembly.Save3(4, arg2, arg2)
            if section[5] > 252 or section[6] > 252:
                long = True
            else:
                long = False
            for_dxf = list(set(for_dxf))
            for file in for_dxf:
                self.to_dxf(sw, file, prj_path, long_overhand=long, part=True)
            for_dxf.clear()
            self.to_dxf(sw, new_path, prj_path, view=(("*Front", config.drv_dim[0] / 2, config.drv_dim[1] / 2),))
            indexes.clear()
            counter += 1

        copied_files.clear()

        self.set_info_message(f'Модели скопированны в {models_path}', True)

        assembly_template = os.path.join(TEMPLATES_FOLDER, 'assembly.asmdot')
        doc = sw.NewDocument(assembly_template, 0, 0, 0)  # create assembly document
        main_assembly = sw.ActiveDoc
        main_assembly.SetTitle2(f'{prj_name}_{man_name}')
        main_assembly_path = os.path.join(models_path, f'{prj_name}_{man_name}.sldasm')
        main_assembly.ShowNamedView2("*Perila", -1)
        main_assembly_ext = main_assembly.Extension
        main_assembly_ext.SaveAs(main_assembly_path, 0, 4, arg1, arg2, arg2)

        current_x = 0
        for key, each in enumerate(self.handrail.bridge):
            name = parts[key][0]
            z_offset = 0
            if each[0].startswith('PO End') and not self.handrail.new_fitting:
                z_offset = 2 / 1000
            for i in range(each[1]):
                current_x += round(parts[key][1] / 1000, 4)
                cur_comp = main_assembly.AddComponent4(name, "", current_x, 0, z_offset)
                current_x += round((each[2] - parts[key][1] + each[3] + each[7] + each[8]) / 1000, 4)
        main_assembly.ViewZoomtofit2()
        fitting = 1 if self.handrail.new_fitting else 0
        blocks = self.make_blocks(sw, self.handrail.lenth_list, models_path, fitting)
        for item in blocks.values():
            block_part = sw.OpenDoc6(item[0], 1, 1, "", arg2, arg2)
        current_assembly = sw.ActivateDoc(f'{prj_name}_{man_name}')
        current_x = 0
        self.handrail.gap_wdth.append(0)
        for ind, lens in enumerate(self.handrail.lenth_list):
            overhand = self.handrail.overhang_end_l if not ind else 0
            current_x += round((blocks[lens][1] - overhand) / 1000, 4)
            current_y = blocks[lens][2] / 1000 * -1
            cur_comp = current_assembly.AddComponent4(blocks[lens][0], "", current_x, 0, current_y)
            current_x += round((blocks[lens][1] + self.handrail.gap_wdth[ind]) / 1000, 4)
        current_assembly.LightweightAllResolved
        current_assembly.Save3(4, arg2, arg2)
        self.handrail.gap_wdth.pop()
        self.to_dxf(sw, main_assembly_path, prj_path, view=(("*Bottom", config.drv_dim[0] / 2, config.drv_dim[1] / 2),))

    def closeEvent(self, e):
        result = QtWidgets.QMessageBox.question(self, "Confirm Dialog", "Сохранить конфигурацию?",
                                                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Cancel,
                                                QtWidgets.QMessageBox.Yes)
        if result == QtWidgets.QMessageBox.Yes:
            if self.save_bridge():
                e.accept()
            else:
                e.ignore()
        elif result == QtWidgets.QMessageBox.Cancel:
            e.ignore()
        else:
            e.accept()


def main():
    app = QtWidgets.QApplication(sys.argv)  # Новый экземпляр QApplication
    window = NccApp()  # Создаём объект класса ExampleApp
    window.show()  # Показываем окно
    app.exec_()  # и запускаем приложение

