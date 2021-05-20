# -*- coding: utf-8 -*-
import pprint
from . import config
import os

pp = pprint.PrettyPrinter(width=60, compact=True)


class Handrail(object):
    def __init__(self):
        super().__init__()

        self.blength = 0
        self.pillar_dist = 1500
        self.pillar_dist_list = []
        self.lenth_list = []
        self.bridge = []
        self.gap_qty = 0
        self.gap_wdth = []
        self.before_gap_l = 260
        self.before_gap_r = 260
        self.before_gaps_l = []
        self.before_gaps_r = []
        self.overhang_end_l = 250
        self.overhang_end_r = 250
        self.beem_dist = 703
        self.use_double_sections = False
        self.colors = []
        self.doubles = []
        self.new_fitting = True

        self.manual = False
        self.manual_config = []

        self.prj_name = ''
        self.prj_manager = ''
        self.par_folder = ''

    def __str__(self):
        label = '1'
        stats = ''
        if self.bridge:
            text = '['
            count = 0
            for key, item in enumerate(self.bridge):
                qty = f' {item[1]}x' if item[0] == 'PO Reg.SLDASM' else ''
                overl = f' {item[5]} |' if item[5] != 0 else ''
                overr = f' {item[6]} |' if item[6] != 0 else ''
                bef_l = f' {item[3]} |' if item[3] != 0 else ''
                bef_r = f' {item[4]} |' if item[4] != 0 else ''
                length = f'{qty} {item[2]} |' if item[2] != 0 else ''

                if item[0] in config.end_sections and key < len(self.bridge) - 1:
                    overr = overr[:-1]
                    gap = f']    {item[8]}    ['
                else:
                    gap = ''

                text += bef_r + overl + length + overr + gap + bef_l
                if item[0] in config.gap_sections:
                    text = text[:-1]
                    text += f'\u2510    {item[7]}    \u250c '
                    count += 1
            text = text[:-1] + ']'

            count = 2
            key_start = 0
            sumarr = []
            sections_qty = 0
            regular_qty = 0

            for i, j in enumerate(self.bridge):
                if j[0] in config.gap_sections:
                    sumarr.append(j[3] + j[7] + j[8])
                sections_qty += j[1]
                if j[0] == 'PO Reg.SLDASM':
                    regular_qty += j[1]

            stats += f'Всего секций\u0009\u0009\u0009\u0009{sections_qty} / 100%\n' \
                     f'Регулярных секций\u0009\u0009\u0009{regular_qty} / {round(regular_qty / sections_qty * 100, 2)}%\n'

            if self.use_double_sections:
                stats += f'Из них двойных\u0009\u0009\u0009\u0009{sum(self.doubles)} / {round(sum(self.doubles) / sections_qty * 100, 2)}%\n'

            stats += f'Нерегулярных секций\u0009\u0009\u0009{sections_qty - regular_qty} / {round((sections_qty - regular_qty) / sections_qty * 100, 2)}%\n\n' \
                     f'Максимально возможное количество двойных секций\u0009{sum(self.doubles)}'

            while True:
                key_end = text.find('\u2510', key_start)
                if key_end == -1:
                    break
                num_pos = text.find('\u250c', key_start)
                ind = text.rfind('|', key_start, key_end)
                label = label.ljust(ind, ' ')
                label += '| ' + str(sumarr[count - 2])
                label = label.ljust(num_pos, ' ')
                label += str(count)
                count += 1
                key_start = num_pos + 1

        else:
            text = 'Расчет не проводился'
            label = ''
        return text, label, stats

    def makelist(self, i, j, k, *args):
        templist = list(config.matrix[i][j][k])
        items = [list(x) for x in templist]
        args = args[0]
        for key, item in enumerate(items):
            for ind, _ in enumerate(item):
                if ind == 0:
                    continue
                item[ind] = args[key][ind]
            if item[1] != 0:
                self.bridge.append(item)

    def make_gaps(self):
        over_l = [self.overhang_end_l]
        if self.gap_qty > 0:
            over_l.extend([0] * self.gap_qty)
            over_r = [0] * self.gap_qty
            over_r.append(self.overhang_end_r)
        else:
            over_r = [self.overhang_end_r]

        if self.before_gaps_r:
            bef_r = self.before_gaps_r.copy()
            bef_r.insert(0, 0)
        else:
            bef_r = [0]
            bef_r.extend([self.before_gap_r] * self.gap_qty)
        if self.before_gaps_l:
            bef_l = self.before_gaps_l.copy()
            bef_l.append(0)
        else:
            bef_l = [self.before_gap_l] * self.gap_qty
            bef_l.append(0)

        return over_l, over_r, bef_l, bef_r

    def calculate(self):
        self.bridge.clear()
        self.doubles.clear()

        over_l, over_r, bef_l, bef_r = self.make_gaps()

        if not self.pillar_dist_list:
            self.pillar_dist_list = [self.pillar_dist] * (self.gap_qty + 1)

        temp = len(self.lenth_list)

        for key, value in enumerate(self.lenth_list):
            current_overhang = bef_l[key] + over_l[key] + bef_r[key] + over_r[key]
            print(current_overhang)
            reg_qty, irreg_len = divmod(value - current_overhang, self.pillar_dist_list[key])
            print(reg_qty, irreg_len)

            if temp == 1:
                k = 0
            elif temp > 1 and key == 0:
                k = 1
            elif 0 < key < temp - 1:
                k = 2
            elif key == temp - 1:
                k = 3

            if k == 1 or k == 2:
                gap = self.gap_wdth[key]
                nextoffset = bef_r[key + 1]
            else:
                gap, nextoffset = 0, 0

            if reg_qty == 0 and irreg_len >= 500 or reg_qty == 1 and irreg_len == 0:
                i, j = 0, 0
                length = irreg_len if irreg_len != 0 else self.pillar_dist_list[key]
                args = [
# имя, кол-во, длина, до шва слева, до шва справа, левый свес, правый свес, шов, смещение до следующего, запас, запас
                    ['', 1, length, bef_l[key], bef_r[key], over_l[key], over_r[key], gap, nextoffset, 0, 0]
                ]
            elif reg_qty > 0:
                i, j = 1, 0
                x = irreg_len + self.pillar_dist_list[key]
                qty = reg_qty - 1
                offset_l, offset_r = bef_l[key], bef_r[key]
                print(offset_l, offset_r)
                if irreg_len == 0:  # сюда можно попасть только при количестве >= 2
                    length_l = length_r = self.pillar_dist_list[key]
                    qty = reg_qty - 2
                elif 0 < irreg_len < 500:
                    length_r = length_l = x // 2
                    if x % 2 == 1 and k != 0:
                        if offset_r == 0:
                            offset_l += 1
                        else:
                            offset_r += 1
                            if self.bridge:
                                self.bridge[-1][8] = offset_r
                    if x % 2 == 1 and k == 0:
                        length_l += 1
                elif 500 <= irreg_len:
                    length_l = self.pillar_dist_list[key]
                    length_r = irreg_len

                if k == 3:
                    length_l, length_r = length_r, length_l
                args = [
                    ['', 1, length_l, 0, offset_r, over_l[key], 0, 0, 0, 0, 0],
                    ['', qty, self.pillar_dist_list[key], 0, 0, 0, 0, 0, 0, 0, 0],
                    ['', 1, length_r, offset_l, 0, 0, over_r[key], gap, nextoffset, 0, 0],
                ]
                self.doubles.append(qty // 2)

            elif reg_qty == 0 and irreg_len < 500:
                i, j = 0, 1
                length = 0
                x = value // 2
                offset_l = x
                offset_r = value - x
                if self.bridge:
                    self.bridge[-1][8] = offset_r
                args = [
                    ['', 1, length, offset_l, offset_r, over_l[key], over_r[key], gap, nextoffset, 0, 0]
                ]

            self.makelist(i, j, k, args)

        if self.use_double_sections:
            inserts, i = [], []
            for k, item in enumerate(self.bridge):
                if item[0] == 'PO Reg.SLDASM' and item[1] > 1:
                    if item[1] % 2 == 0:
                        item[1] //= 2
                        item[2] *= 2
                        item[9] = 1
                    else:
                        temp = item.copy()
                        temp[1] //= 2
                        temp[2] *= 2
                        temp[9] = 1
                        item[1] = 1
                        inserts.append(temp)
                        i.append(k)
            inserts.reverse()
            i.reverse()
            for ind, val in enumerate(i):
                self.bridge.insert(val, inserts[ind])

        self.bridge = [tuple(x) for x in self.bridge]

        pp.pprint(self.bridge)
        print('')
        return self.bridge

    def convert_bool(self, param):
        if param:
            return 'yes'
        return 'no'

    def processConfig(self):
        gaps = '0'
        gaps_l = self.before_gap_l
        gaps_r = self.before_gap_r
        sectors = self.pillar_dist
        manual = ''
        len_list = '0'

        if self.gap_wdth:
            gaps = ' '.join(str(x) for x in self.gap_wdth)

        if self.before_gaps_l:
            gaps_l = ' '.join(str(x) for x in self.before_gaps_l)

        if self.before_gaps_r:
            gaps_r = ' '.join(str(x) for x in self.before_gaps_r)

        if self.pillar_dist_list:
            sectors = ' '.join(str(x) for x in self.pillar_dist_list)

        if self.lenth_list:
            len_list = ' '.join(str(x) for x in self.lenth_list)

        heading = f'# Название проекта\n{self.prj_name}\n\n' \
                  f'# Менеджер\n{self.prj_manager}\n\n# Цвета, пока не реализовано\n\n\n'

        common_params = f'# Количество деформационных швов\n{self.gap_qty}\n\n' \
                        f'# Общая длина моста\n{self.blength}\n\n' \
                        f'# Ширина "лесенки", 703 или 728 мм\n{self.beem_dist}\n\n' \
                        f'# Свес концевой левой секции\n{self.overhang_end_l}\n\n' \
                        f'# Свес концевой правой секции\n{self.overhang_end_r}\n\n' \
                        f'# Базовое расстояние между стойками\n{sectors}\n\n' \
                        f'# Допустимые значения для чекбоксов:\n' \
                        f'# yes, y, Yes, YES, +, д, да, Да, ДА\n' \
                        f'# no, n, No, NO, -, н, нет, НЕТ, Нет\n' \
                        f'# Использовать двойные секции, yes/no\n{self.convert_bool(self.use_double_sections)}\n\n' \
                        f'# Использовать верхние фитинги, yes/no\n{self.convert_bool(self.new_fitting)}\n\n\n'

        params = f'# Ширина деформационных швов\n{gaps}\n\n' \
                 f'# Длины участков моста\n{len_list}\n\n' \
                 f'# Отступ от стоек до швов слева\n{gaps_l}\n\n' \
                 f'# Отступ от стоек до швов справа\n{gaps_r}'

        if self.manual:
            manual = "\n".join(self.manual_config)

        return heading + common_params + params + config.manual + manual + '\n'

    def save_conf(self, path=None):
        if not path:
            file_name = f'{self.prj_name}_{self.prj_manager}.txt'
            folder = os.path.join(self.par_folder, f'{self.prj_name}_{self.prj_manager}')
            file_path = os.path.join(folder, file_name)

            if not os.path.exists(folder):
                os.makedirs(folder)
        else:
            file_path = path

        text = self.processConfig()

        with open(file_path, 'w', encoding='utf-8') as fout:
            fout.write(text)

        return True



