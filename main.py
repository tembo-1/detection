from search import Search

fields = Search('excel.xlsx', 'catalog.json')

with open("freqsy.txt", "wt", encoding="utf-8") as file:
    for index, key in enumerate(fields.search(), start=1):
        file.writelines(f"{index} {key}\n")

# todo:
#  ? Передавать название файла параметром (`argparse`) или из конфигурации (`configparser`, `tomllib`, etc.)
#  + Записать в  excel
#  + Взять 200 записей
#  + Создать БД
