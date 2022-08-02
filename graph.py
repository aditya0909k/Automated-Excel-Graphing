import pandas as pd
from matplotlib import pyplot as plt
import xlwings as xw
import sys, os

#path = os.path.dirname(sys.executable)

file = input('filename: ')

#read in dataframe and workbook
df = pd.read_excel(f'{file}', skiprows=1, skipfooter=1)
wb = xw.Book(f'{file}')

#create pivot table, sort in order top to bottom, create plot
table = df.pivot_table(values=['Amount'], index='Name').sort_values(by='Amount')
ax = table.plot(kind='barh', figsize=(12, 9), color='#32a852', 
                width=.60, legend=False, stacked=True) 
ax.xaxis.set_major_formatter('${x:1.0f}')

#make plot look better
plt.title("Profit by Employee/Company", fontname='Times New Roman', fontsize=15) 
plt.xlabel("")
plt.ylabel("")
plt.xticks(fontname='Times New Roman', fontsize=12)
plt.yticks(fontname='Times New Roman', fontsize=12)
plt.bar_label(ax.containers[0], labels=[f'${p:1.2f}' for p in table['Amount']], 
                padding=3, fontname='Times New Roman', fontsize=12)
plt.xlim(None, 1.15*max(table.values))

#get plot's figure for export
fig = plt.gcf()

#export plot to excel sheet
sheet = wb.sheets[0]
pic = sheet.pictures.add(fig, name="Profit Summary", update=True, 
                    left=sheet.range("I2").left, top=sheet.range("I2").top)
