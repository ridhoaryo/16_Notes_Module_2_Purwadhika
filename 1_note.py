import xlrd
# cara install:
# pip install xlrd
# py -m pip install xlrd
# conda install xlrd

# cara install beberapa package
# py -m pip install namaPackage1 namaPackage2 namaPackage3

file = xlrd.open_workbook('file.xlsx')
# sheet = file.sheet_by_index(0)
sheet = file.sheet_by_name('data_orang')
# print(f'Baris {sheet.nrows}, Kolom {sheet.ncols}')


# print(sheet.cell_value(0,0))
# print(sheet.cell_value(0,1))
# print(sheet.cell_value(0,2))
# print(sheet.cell_value(0,3))
cols = []
for i in range(sheet.ncols):
    # print(sheet.cell_value(0,i))
    cols.append(sheet.cell_value(0,i))
# print(cols)

# print(f'Menggunakan method row_values(index): {sheet.row_values(0)}')

# How to make json file from python list of list
header_list = sheet.row_values(0) # separate the header
value_list = [] # make a list for values
for i in range(1,sheet.nrows):
    value_list.append(sheet.row_values(i))

for_json = []
for i in range(len(header_list)-1):
    y = dict(zip(header_list,value_list[i])) # make dict from header pairing with each list in value_list
    for_json.append(y)

# print(header_list)
# print(value_list)  
# print(for_json)

import json
with open('file.json', 'w') as file:
    json.dump(for_json, file)

import csv
with open('file_csv.csv', 'w', newline='') as file:
    write = csv.DictWriter(file, header_list)
    write.writeheader()
    write.writerows(for_json)

