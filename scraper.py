from openpyxl import load_workbook
import requests
import os
import re
from jinja2 import Environment, FileSystemLoader, select_autoescape

url = 'https://www.alko.fi/INTERSHOP/static/WFS/Alko-OnlineShop-Site/-/Alko-OnlineShop/fi_FI/Alkon%20Hinnasto%20Tekstitiedostona/alkon-hinnasto-tekstitiedostona.xlsx'
response = requests.get(url)
#with open('alkon-hinnasto-tekstitiedostona.xlsx', 'wb') as file:
#    file.write(response.content)

wb = load_workbook(filename ='alkon-hinnasto-tekstitiedostona.xlsx')
ws = wb.active

package_path = os.path.dirname(os.path.abspath(__file__))
templates_dir = os.path.join(package_path, "templates")
env = Environment(
    loader=FileSystemLoader(templates_dir),
    autoescape=select_autoescape(["html", "xml", "j2"])
)
template = env.get_template("template.j2")

array = []

f = open('index.html', 'w')

for index, row in enumerate(ws.iter_rows(min_row=5, values_only=True)):
    if row[21] == None or float(row[21]) == 0:
        continue
    # 0-5, 8, 21
    array.append(['https://www.alko.fi/tuotteet/' + row[0], row[1], row[2], row[3], row[4], row[5], row[8], row[21], str(round(float(row[5]) / float(row[21]) * 100, 2))])

os.remove('alkon-hinnasto-tekstitiedostona.xlsx')
f.write(template.render(table=re.sub('None', '\'\'', str(array))))
f.close()