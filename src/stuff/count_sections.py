import pprint
import re

regex = r"[^\d-]+"

pp = pprint.PrettyPrinter(width=38, compact=True)

blength = int(input('Длина моста в мм: '))
gap_qty = int(input('Количество деформационных швов: '))
print('Укажите ширину деформационных швов слева направо через пробел.\nЕсли все одинаковые, укажите ширину 1 раз.')
while True:
    text = input('Ширина, мм: ')
    text = (re.sub(regex, ' ', text)).strip()
    if text:
        gap_wdth = [int(x) for x in text.split(' ') if int(x) > 0]
        if len(text.split(' ')) == 1:
            gap_wdth = [gap_wdth[0]] * gap_qty
            break
        if len(gap_wdth) != gap_qty:
            print(f'Количество размеров в {gap_wdth} не совпадает с количеством швов.\nВведите данные еще раз.')
            continue
        break

before_gap_l = int(input('Отступ от стоек до швов слева в мм (стандартно 250 мм): '))
before_gap_r = int(input('Отступ от стоек до швов справа в мм (стандартно 250 мм): '))

while True:
    text = input('Эти размеры фиксированы и не могут быть измененны? y/n ')
    if text == 'y':
        bgl_fix = True
        break
    elif text == 'n':
        bgl_fix = False
        break
    continue

pillar_dist = int(input('Расстояние между опорами в регулярной секции в мм (стандартно 1500 мм): '))
overhang_end_l = int(input('Свес концевой левой секции в мм (стандартно 252 мм): '))
overhang_end_r = int(input('Свес концевой правой секции в мм(стандартно 252 мм): '))

print('Укажите длины участков моста до и между дефшвами слева направо')
lenth_list = []
for i in range(gap_qty + 1):
    len4 = blength - sum(lenth_list) - sum(gap_wdth[:i])
    if i < gap_qty:
        while True:
            len1 = int(input(f'Длина участка {i + 1} в мм: '))
            len2 = overhang_end_l + before_gap_l
            len3 = overhang_end_r + before_gap_r
            if i == 0 and len1 < len2 or i == gap_qty and len4 < len3:
                print('Первый или последний участок не может быть короче суммы длин отступа и свеса')
                continue
            lenth_list.append(len1)
            break
        print(f'До правого конца моста осталось {len4} мм')
        continue
    print(f'Последний участок длиной {len4} мм добавлен автоматически')
    lenth_list.append(len4)
pp.pprint(lenth_list)

bridge = []

for i, value in enumerate(lenth_list):
    if i == 0:
        reg_qty, irreg_len = divmod(value - overhang_end_l - before_gap_l, pillar_dist)
        print(f'Участок моста 1: кол-во регулярных - {reg_qty}, остаток - {irreg_len}')
        if reg_qty == 0:
            if irreg_len < 500:
                before_gap_l_temp = before_gap_l
                if bgl_fix:
                    overhang_end_l += irreg_len
                else:
                    x = irreg_len // 2
                    overhang_end_l += x
                    before_gap_l_temp += irreg_len - x
                bridge.append(('L_end_1pillow_gap_section', 1, 0, overhang_end_l, before_gap_l_temp))
            if irreg_len >= 500:
                bridge.append(('L_end_2pillows_gap_section', 1, irreg_len, overhang_end_l, before_gap_l))

        if reg_qty == 1:
            if irreg_len < 500:
                pillar_dist1 = (pillar_dist + irreg_len) // 2
                pillar_dist2 = pillar_dist + irreg_len - pillar_dist1
                bridge.append(('L_end_section', 1, pillar_dist1, overhang_end_l, 0))
                bridge.append(('L_gap_section', 1, pillar_dist2, 0, before_gap_l))
            if irreg_len >= 500:
                bridge.append(('L_end_section', 1, pillar_dist, overhang_end_l, 0))
                bridge.append(('L_gap_section', 1, irreg_len, 0, before_gap_l))

        if reg_qty >= 2:
            if irreg_len < 500:
                pillar_dist1 = (pillar_dist + irreg_len) // 2
                pillar_dist2 = pillar_dist + irreg_len - pillar_dist1
                bridge.append(('L_end_section', 1, pillar_dist, overhang_end_l, 0))
                if reg_qty - 2 == 0:
                    bridge.append(('Ireg_section', 1, pillar_dist1, 0, 0))
                    bridge.append(('L_gap_section', 1, pillar_dist2, 0, before_gap_l))
                else:
                    bridge.append(('Reg_section', reg_qty - 2, pillar_dist, 0, 0))
                    bridge.append(('Ireg_section', 1, pillar_dist1, 0, 0))
                    bridge.append(('L_gap_section', 1, pillar_dist2, 0, before_gap_l))
            if irreg_len >= 500:
                bridge.append(('L_end_section', 1, pillar_dist, overhang_end_l, 0))
                bridge.append(('Reg_section', reg_qty - 1, pillar_dist, 0, 0))
                bridge.append(('L_gap_section', 1, irreg_len, 0, before_gap_l))

    elif i == len(lenth_list) - 1:
        reg_qty, irreg_len = divmod(value - overhang_end_r - before_gap_r, pillar_dist)
        print(f'Последний участок моста: кол-во регулярных - {reg_qty}, остаток - {irreg_len}')
        if reg_qty == 0:
            if irreg_len < 500:
                before_gap_r_temp = before_gap_r
                if bgl_fix:
                    overhang_end_r += irreg_len
                else:
                    x = irreg_len // 2
                    overhang_end_r += x
                    before_gap_r_temp += irreg_len - x
                bridge.append(('R_end_1pillow_gap_section', 1, 0, before_gap_r_temp, overhang_end_r))
            if irreg_len >= 500:
                bridge.append(('R_end_2pillows_gap_section', 1, irreg_len, before_gap_r, overhang_end_r))

        if reg_qty == 1:
            if irreg_len < 500:
                pillar_dist1 = (pillar_dist + irreg_len) // 2
                pillar_dist2 = pillar_dist + irreg_len - pillar_dist1
                bridge.append(('R_gap_section', 1, pillar_dist1, before_gap_r, 0))
                bridge.append(('R_end_section', 1, pillar_dist2, 0, overhang_end_r))
            if irreg_len >= 500:
                bridge.append(('R_gap_section', 1, irreg_len, before_gap_r, 0))
                bridge.append(('R_end_section', 1, pillar_dist, 0, overhang_end_r))

        if reg_qty >= 2:
            if irreg_len < 500:
                pillar_dist1 = (pillar_dist + irreg_len) // 2
                pillar_dist2 = pillar_dist + irreg_len - pillar_dist1
                if reg_qty - 2 == 0:
                    bridge.append(('R_gap_section', 1, pillar_dist1, before_gap_r, 0))
                    bridge.append(('Ireg_section', 1, pillar_dist2, 0, 0))
                else:
                    bridge.append(('R_gap_section', 1, pillar_dist1, before_gap_r, 0))
                    bridge.append(('Ireg_section', 1, pillar_dist2, 0, 0))
                    bridge.append(('Reg_section', reg_qty - 2, pillar_dist, 0, 0))
                bridge.append(('R_end_section', 1, pillar_dist, 0, overhang_end_r))
            if irreg_len >= 500:
                bridge.append(('R_gap_section', 1, irreg_len, before_gap_r, 0))
                bridge.append(('Reg_section', reg_qty - 1, pillar_dist, 0, 0))
                bridge.append(('R_end_section', 1, pillar_dist, 0, overhang_end_r))

    else:
        reg_qty, irreg_len = divmod(value - before_gap_r - before_gap_l, pillar_dist)
        print(f'Участок моста {i + 1}: кол-во регулярных - {reg_qty}, остаток - {irreg_len}')
        if reg_qty == 0:
            pass
        if reg_qty == 1:
            pass
        if reg_qty >= 2:
            if irreg_len < 500:
                pillar_dist1 = (pillar_dist + irreg_len) // 2
                pillar_dist2 = pillar_dist + irreg_len - pillar_dist1
                bridge.append(('R_gap_section', 1, pillar_dist1, before_gap_r, 0))
                # if reg_qty - 2 == 0:
                #     bridge.append(('Ireg_section', 1, pillar_dist1, 0, 0))
                #     bridge.append(('L_gap_section', 1, pillar_dist2, 0, before_gap_l))
                # else:
                bridge.append(('Reg_section', reg_qty - 1, pillar_dist, 0, 0))
                bridge.append(('L_gap_section', 1, pillar_dist2, 0, before_gap_l))
            if irreg_len >= 500:
                bridge.append(('R_gap_section', 1, pillar_dist, before_gap_r, 0))
                bridge.append(('Reg_section', reg_qty - 1, pillar_dist, 0, 0))
                bridge.append(('L_gap_section', 1, irreg_len, 0, before_gap_l))

print('\nРезультат:\nНазвание секции | количество | длина между опор | вылет слева | вылет справа')
pp.pprint(bridge)

