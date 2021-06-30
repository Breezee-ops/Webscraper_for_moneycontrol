
import requests
import bs4 as bs
import xlsxwriter as xl

# variables

y = 0
ROW = 2
COL = 0
l = 0
en = 0

# html parsers

r_profitandloss = requests.get('https://www.moneycontrol.com/financials/oilandnaturalgascorporation/consolidated-profit-lossVI/ONG#ONG')
soup_prof = bs.BeautifulSoup(r_profitandloss.content,'lxml')
table_prof = soup_prof.find(class_='mctable1').find_all('tr')[1:]

# finding entries based on table count

for row in table_prof:
    y += 1
    cell = [i.text for i in row.find_all('td')]
    if y == 8:
        revenue = cell
    elif y == 28:
        net_profit = cell
    elif y == 36:
        div = cell

# enumerating

df_rev =(list(x) for x in enumerate(revenue))
df_prof =(list(n) for n in enumerate(net_profit))
df_div =(list(j) for j in enumerate(div))

# working excel

workbook = xl.Workbook('investment data.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write(0, 0, "oil and ntural gas corporation")
worksheet.write(1, 1, '2021')
worksheet.write(1, 2, '2020')
worksheet.write(1, 3, '2019')
worksheet.write(1, 4, '2018')
worksheet.write(1, 5, '2017')

for index, entry in df_rev:
    worksheet.write(ROW, COL + index, entry)
en+=1

for index, entry in df_prof:
    worksheet.write(ROW + en, COL + index, entry)
en+=1

for index, entry in df_div:
    worksheet.write(ROW + en, COL + index, entry)
en+=1

workbook.close()