import json
import xlsxwriter
from utils import write_head_data
from os import system
import requests
from bs4 import BeautifulSoup

with open("src/assets/codes.json", encoding="utf8") as file_:
    data = json.load(file_)


book = xlsxwriter.Workbook("src/price-fluctuation.xlsx")

down_trend_format = book.add_format({
    "bg_color": "#FFC7CE",
    "font_color": "#9C0006"
})

up_trend_format = book.add_format({
    "bg_color": "#C6EFCE",
    "font_color": "#006100"
})

sheet = book.add_worksheet()
sheet.set_column("B:B", 15)

titles = ["Code Name", "1 Day", "7 Day", "20 Day", "60 Day"]

price_fluctuation = {}
row_index = 0
for key in data.keys():
    sheet.write(row_index, 0, key)
    row_index += 1

    write_head_data(sheet, row_index, titles)
    row_index += 1
    
    start_row_index = row_index

    for info in data[key]:
        response = requests.get(f"https://finance.yahoo.com/quote/{info['code']}/history", headers={'USER-AGENT': 'Mozilla/5.0'})
        soup = BeautifulSoup(response.text, "html.parser")
        code_data = soup.select("tr.BdT td")[:-3]

        close_data = []
        data_index = 0
        while len(close_data) <= 61:
            try:
                close_data.append(float(code_data[4+7*data_index].text.replace(",", "")))
            except:
                print(f"{info['code']} no -- {code_data[7*data_index].text} -- data")
            
            data_index += 1

        if close_data != []:
            values = [
                info["name"],
                round((close_data[0] - close_data[1]) / close_data[1], 2),
                round((close_data[0] - close_data[7]) / close_data[7], 2),
                round((close_data[0] - close_data[20]) / close_data[20], 2),
                round((close_data[0] - close_data[60]) / close_data[60], 2)
            ]

            write_head_data(sheet, row_index, values)
            row_index += 1

            sheet.conditional_format(f"C{start_row_index+1}:F{row_index}", {
                "type": "cell",
                "criteria": "<",
                "value": 0,
                "format": down_trend_format
            })

            sheet.conditional_format(f"C{start_row_index+1}:F{row_index}", {
                "type": "cell",
                "criteria": ">=",
                "value": 0,
                "format": up_trend_format
            })

    row_index += 1

book.close()
system(".\src\price-fluctuation.xlsx")
