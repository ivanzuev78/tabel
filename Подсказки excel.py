import openpyxl as opx
from random import randint
import datetime
from openpyxl.styles.borders import Border, Side

# Добавляет строку в конец листа. Каждое значение в списке в ячейки попорядку
# ws.append(['Дата', '3416', '3419', 'Производство', 'Итого'])


wb = opx.Workbook()  # Создаём рабочую книгу
ws = wb.active  # Запоминаем активный лист

thin_border = Border(left=Side(style='thin'),
                     right=Side(style='thin'),
                     top=Side(style='thin'),
                     bottom=Side(style='thin'))

ws['B3'].border = thin_border