import openpyxl, pytz
from datetime import datetime, time

book = openpyxl.open("C:/Users/Noizz/Desktop/learn/Excel отчёт/Сводка по объектам SM.xlsx", read_only=False)
sheet = book.active

for row in range(2, sheet.max_row + 1):  # Начинаем с 2-ой строки, так как они считаются с единицы, а в ней у нас строка 'КОНТР СРОК' и её не привести к формату даты (так как буквы)!
    zno = sheet[row][1].value
    object = sheet[row][5].value
    control_time = sheet[row][6].value
    group = sheet[row][8].value
    employee = sheet[row][9].value
    date = control_time
    # format = "%y/%m/%d %H:%M:%S"
    # date_obj = datetime.strptime(control_time, format)
    dt = datetime.now(pytz.timezone('Europe/Moscow'))
    dt = dt.replace(microsecond=0)
    if date > datetime.now(pytz.timezone('Europe/Moscow')):
        print(zno, object, control_time, group, employee)
