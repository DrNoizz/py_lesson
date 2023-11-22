import shutil
import os

result = shutil.disk_usage('C:/')
gb = 10 **9
with open('file.txt', 'a', encoding='utf-8') as f:

    print(f"Всего места: {result.total/gb:.2f}" + ' ' + 'Gb', file=f)
    print(f"Места использовано: {result.used/gb:.2f}" + ' ' + 'Gb', file=f)
    print(f"Места осталось: {result.free/gb:.2f}" + ' ' + 'Gb', file=f)
