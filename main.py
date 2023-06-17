import pandas as pd
import openpyxl
from openpyxl import Workbook
from urllib.parse import quote_plus
import requests
from bs4 import BeautifulSoup

def check(e):
    if e:
        raise e

def is_alpha(c):
    return (97 <= ord(c) and ord(c) <= 122) or (65 <= ord(c) and ord(c) <= 90)

# open file
args = 'companies.csv'
try:
    data = pd.read_csv(args)
except Exception as e:
    check(e)

print(data)
# create spreadsheet
wb = Workbook()
ws = wb.active
ws.title = "SOSCA"
ws.append(["Entity #", "Registration Date", "Status", "Entity Name", "Jurisdiction", "Agent", "# Of Entries", "Link"])

# iterate over records
for index, row in data.iterrows():
    # fetch page
    url = "https://businesssearch.sos.ca.gov/CBS/SearchResults?filing=&SearchType=LPLLC&SearchCriteria={}&SearchSubType=Keyword".format(quote_plus(row[0]))
    print(f"Loading {url}")

    try:
        res = requests.get(url)
        check(res.status_code != 200)
        soup = BeautifulSoup(res.text, 'html.parser')
        entity_table = soup.find(id='enitityTable')
        tds = entity_table.find_all('td') if entity_table else []

        if len(tds) != 0:
            for i, td in enumerate(tds):
                cell_value = td.get_text().strip()
                if i % 6 == 3:
                    off = 0
                    nl = cell_value[1:].find('\n')
                    while not is_alpha(cell_value[off + nl]):
                        off += 1
                    cell_value = cell_value[nl + off + 1:]
                ws.append([cell_value])
        else:
            ws.append([row[0]])

        ws.append([url])

    except Exception as e:
        check(e)

# save the workbook
wb.save("./SOSCA.xlsx")
