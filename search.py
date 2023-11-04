from pycatsearch import catalog
import json
import openpyxl

class Search:
    def __init__(self, excel_path, json_path):
        self.json_path = json_path
        self.excel = openpyxl.load_workbook(excel_path)

        self.__prepareJsonData()

    def __prepareJsonData(self):
        with open(self.json_path) as file:
            self.data = json.load(file)['catalog']

    def search(self):
        sheet = self.excel.active  

        self.accord = []

        for row in sheet.iter_rows(min_row=2, min_col=1, max_col=3):
            excel_freq = row[0].value

            if excel_freq is None:
                break
            
            for tag, array in self.data.items():
                flag = False
                for freq in array['lines']:
                    if (freq['frequency'] + 0.1 >= excel_freq and freq['frequency'] - 0.1 <= excel_freq):
                        self.accord.append({array['name'] : freq['frequency']})
                        flag = True
                        break
                if (flag):
                    break 

        return self.accord  