from bs4 import BeautifulSoup
import urllib.request
import datetime
import xlwt
from xlwt import Workbook

# Get the current date
x = datetime.datetime.now()

# Craete Excel Workbook
wb = Workbook()
sheetName = "Libor-" + x.strftime("%b") + " " + x.strftime("%d")
sheet1 = wb.add_sheet(str(sheetName), cell_overwrite_ok=True)

# Scrape WSJ Web Page
wsj = "https://www.wsj.com/market-data/bonds"

page = urllib.request.urlopen(wsj)

soup = BeautifulSoup(page, "lxml")

libor_table = soup.find_all(
    'table', class_="WSJTables--table--1SdkiG8p ")


# Generate lists
A = []
B = []
C = []
D = []
E = []
F = []
G = []

for row in libor_table[0].find_all("tr"):
    cells = row.find_all("td")
    if len(cells) == 5:  # Extract Body of Table
        A.append(cells[0].find(text=True))
        C.append(cells[1].find(text=True))
        D.append(cells[2].find(text=True))
        E.append(cells[3].find(text=True))
        F.append(cells[4].find(text=True))

# Write Rates to Excel File
r = 1
for i in A:
    col = 0
    sheet1.write(0, col, 'LIBOR')
    sheet1.write(r, col, A[r-1])
    r += 1

r = 1
for i in C:
    col = 1
    sheet1.write(0, col, 'LATEST')
    sheet1.write(r, col, C[r-1])
    r += 1

r = 1
for i in D:
    col = 2
    sheet1.write(0, col, 'WK AGO')
    sheet1.write(r, col, D[r-1])
    r += 1

r = 1

for i in E:
    col = 3
    sheet1.write(0, col, 'HIGH')
    sheet1.write(r, col, E[r-1])
    r += 1

r = 1
for i in F:
    col = 4
    sheet1.write(r, col, F[r-1])
    sheet1.write(0, col, 'LOW')
    r += 1


wb.save('WSJ Libor.xls')
