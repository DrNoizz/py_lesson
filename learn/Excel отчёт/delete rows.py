import openpyxl
from openpyxl import Workbook
from datetime import datetime, timedelta, time

book = openpyxl.open("C:/Users/Noizz/Desktop/learn/Excel отчёт/Сводка по объектам SM.xlsx", read_only=False)
sheet = book.active
# for row in range(1, 13):
#     control_time = sheet[row][6].value
#     print(control_time)


for row in range(2, sheet.max_row + 1):  # Начинаем с 2-ой строки, так как они считаются с единицы, а в ней у нас строка 'КОНТР СРОК' и её не привести к формату даты (так как буквы)!
    control_time = sheet[row][6].value
    date = control_time
    # format = "%y/%m/%d %H:%M:%S"
    # date_obj = datetime.strptime(control_time, format)
    dt = datetime.now()
    dt = dt.replace(microsecond=0)
    if date > datetime.now():
        print(control_time)
