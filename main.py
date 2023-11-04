from search import Search
            
fields = Search('excel.xlsx', 'catalog.json')

for key in fields.search():
    print(key)

