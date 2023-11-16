import pandas as pd


file = "C:/Users/Noizz/Desktop/learn/Excel отчёт/Сводка по объектам SM.xlsx"
df = pd.read_excel(file)
# снимаем ограничения на вывод по ширине и кол-ву строк
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.max_colwidth', None)
pd.options.display.expand_frame_repr = False

df.sort_values(by = ["ОБЪЕКТ", "КОНТР СРОК"])  # Если нужно сортировать по убыванию: df.sort_values(by = ["ОБЪЕКТ", "КОНТР СРОК"], ascending = False),
# если по первому столбцу по возрастанию, а по второму - по убыванию: ascending = [False, True]
df[['НОМЕР ЗНО','ОБЪЕКТ','КОНТР СРОК', 'ГРУППА', 'СОТРУДНИК']]
