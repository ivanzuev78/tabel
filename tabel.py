import openpyxl as opx
from random import randint as r
import datetime
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Font


def count_hour(mon,year):
    day = 1
    count_time = 0
    while True:
        try:
            if datetime.datetime(year, mon, day).isoweekday() < 6:
                count_time += 8
            day += 1
        except:
            break
    return count_time


def workday(d, m, y, holiday):

    while True:
        try:
            if d in holiday:
                yield False
            else:
                yield True if datetime.datetime(y, m, d).isoweekday() < 6 else False
        except:
            break
        d += 1


def fill_tabel(all_time, workdaygen):
    all_stroki = []
    for ind, val in enumerate(all_time):
        all_time[ind] = int(val)
    while True:
        if sum(all_time) == 0:
            print('good exit')
            break
        try:
            curent_day = []
            if next(workdaygen):
                for ind, top in enumerate(all_time):
                    if top > 0:
                        if ind == len(all_time) - 1:
                            hour = 8 - sum(curent_day)
                            curent_day.append(hour)
                            all_time[ind] -= hour
                        elif sum(curent_day) < 8:
                            if sum(all_time[ind+1:]) > 8 or (sum(curent_day) + sum(all_time[ind+1:]) > 8):
                                if top > 8:
                                    hour =r(1, 8 - sum(curent_day))
                                else:
                                    hour = top
                                all_time[ind] -= hour
                                curent_day.append(hour)
                            elif sum(all_time[ind+1:]) < 8:
                                hour = 8 - sum(curent_day) - sum(all_time[ind+1:])
                                all_time[ind] -= hour
                                curent_day.append(hour)
                        else:
                            curent_day.append(0)
                    else:
                        curent_day.append(0)
                all_stroki.append(curent_day)
            else:
                all_stroki.append([' - ' for _ in range(len(all_time))])
        except:
            # print('exception')
            break
    return all_stroki


def make_tabel_func(user, month, year, topic, topic_val, holiday):

    topdict = {}  # Словарь (тема: кол-во часов)
    wb = opx.Workbook()  # Создаём рабочую книгу
    ws = wb.active  # Запоминаем активный лист

    # Границы для форматирования ячеек
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))

    defalt_border = Border(left=Side(style=None),
                           right=Side(style=None),
                           top=Side(style=None),
                           bottom=Side(style=None))
    #  Размер шрифта 13
    ft = Font(name='Times New Roman',
              size=13,
              bold=False,
              italic=False,
              vertAlign=None,
              underline='none',
              strike=False,
              color='FF000000')
    # Номер месяца строкой, что бы вставлять 01, 02 и тд.
    if month < 10:
        monprint = '0' + str(month)
    else:
        monprint = str(month)

    # Словарь месяцев 1: "январь" и т.д.
    month_list = {1: 'январь', 2: 'февраль', 3: 'март', 4: 'апрель', 5: 'май', 6: 'июнь',
                  7: 'июль', 8: 'август', 9: 'сентябрь', 10: 'октябрь', 11: 'ноябрь', 12: 'декарбь'}

    # Количество активных тем
    topcount = 0
    for i, top in enumerate(topic):
        if top:
            ws[f'{chr(66+topcount)}4'].value = top
            topdict[top] = int(topic_val[i])
            topcount += 1


    # Делаем шапку
    ws.merge_cells(f'A1:{chr(66+topcount)}1')  # Объединяем ячейки для названия
    ws.merge_cells(f'A2:{chr(66+topcount)}2')  # Объединяем ячейки для названия
    ws.merge_cells(f'B3:{chr(66 + topcount - 1)}3')  # Объединяем ячейки для номера инвентарнго объекта

    ws['A1'].value = f'Табель учета затрат времени сотрудника ОВНТ {user}'  # Добавляет в выбраную ячейку указанное значение
    ws['A2'].value = f'на «{month_list[month]}» {year} года по инвентарным объектам'
    ws['B3'].value = 'Номер инвентарного объекта'

    ws[f'{chr(66+topcount)}4'].value = 'Сумма'
    ws['A4'].value = 'Дата'


    w_day = workday(1, month, year, holiday)
    all_day = []

    for i in topic_val:
        if i:
            all_day.append(i)
    sum_total = [0 for _ in range(len(all_day))]
    tabel = fill_tabel(all_day, w_day)

    for ind, i in enumerate(tabel):
        day = [ind + 1]
        if type(i[0]) == int:
            for index, k in enumerate(i):
                sum_total[index] += k
            sums = [sum(i)]
            ws.append(day + i + sums)
        else:
            ws.append(day + i + [' - '])
    ws.append(['Итого'] + sum_total + [sum(sum_total)])


    for row in ws:  # Проходим по всем строкам на листе
        for cell in row:  # Проходим по всем ячейкам в строке
            cell.alignment = opx.styles.Alignment(horizontal='center')  # Выравниваем ячейку по центру
            cell.border = thin_border   # создаем границы клеткам
            cell.font = ft

    for col in range(65, 67 + topcount):
        for row in range(1,3):
            ws[f'{chr(col)}{row}'].border = defalt_border


    for col in range(66, 67 + topcount):
        if col == 65 + topcount:
            ws.column_dimensions[f'{chr(col)}'].width = 20  # делаем шире столбец с Производством
        else:
            pass
            ws.column_dimensions[f'{chr(col)}'].width = 20 - 2 * topcount

    ws.merge_cells(f'A41:{chr(66 + topcount)}41')
    ws['A41'].value = 'Начальник ОВНТ                                                                Керестень А.А.'
    ws['A41'].font = ft
    ws['A41'].alignment = opx.styles.Alignment(horizontal='center')

    wb.save(f'{monprint}.Табель_{user}.xlsx')  # Сохраняем книгу


if __name__ == '__main__':
    # make_tabel_func('Ivan')
    # print(count_hour(1,2020))
    sum_ = [0, 0, 0, 0, 0]
    for i in fill_tabel(1, 2020, ['q','w','e', 'a'], [50, 50, 34, ]):
        try:
            for ind, k in enumerate(i):
                sum_[ind] += k
        except:
            pass
        print(i)
    print(sum_)