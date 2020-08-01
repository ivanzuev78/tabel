import openpyxl as opx
from random import randint
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
    if d in holiday:
        return False
    return True if datetime.datetime(y, m, d).isoweekday() < 6 else False


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

    # Заполняем таблицу
    day = 1
    nulbefore = 0
    while True:
        try:
            if workday(day, month, year, holiday):
                row_to_add = [day]
                if nulbefore:
                    row_to_add += [0 for _ in range(nulbefore)]
                for i in range(topcount):
                    if topdict[ws[f'{chr(66 + i)}4'].value] > 8:
                        row_to_add.append(8)
                        topdict[ws[f'{chr(66 + i)}4'].value] -= 8
                        for _ in range(i,topcount - 1):
                            row_to_add.append(0)
                        break
                    elif topdict[ws[f'{chr(66 + i)}4'].value] > 0:
                        row_to_add.append(topdict[ws[f'{chr(66 + i)}4'].value])
                        row_to_add.append(8-topdict[ws[f'{chr(66 + i)}4'].value])
                        topdict[ws[f'{chr(67 + i)}4'].value] -= topdict[ws[f'{chr(66 + i)}4'].value]
                        topdict[ws[f'{chr(66 + i)}4'].value] = 0
                        nulbefore += 1
                        for _ in range(i,topcount - 2):
                            row_to_add.append(0)
                        break

                row_to_add.append(8)
            else:
                row_to_add = [day] + ['-' for _ in range(topcount + 1)]
            ws.append(row_to_add)
            day += 1
        except:
            break

    row_to_add = ['Итого']
    for k in range(topcount + 1):
        sum_top = 0
        for i in range(5, 5 + day):
            try:
                sum_top += int(ws[f'{chr(66 + k)}{i}'].value)
            except:
                pass
        row_to_add.append(sum_top)
    ws.append(row_to_add)



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
    print(count_hour(1,2020))