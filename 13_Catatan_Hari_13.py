# data = open('contoh_aja.txt', 'r')
# data.read()
# data.readlines()

# data = open('contoh_aja.txt', 'w')
# data.write('abcdefg')

# ----- BIKIN FILE BARU -----

# data2 = open('file.py', 'w')
# data2.write('print(\'ini dibuat dari python\')')

# ----- MEMBACA FILE DARI FILE .TXT -----
# data_def = open('data.txt', 'r')
# x = data_def.read()

# ----- menulis text menjadi fungsi dari file data.txt ke file data.py
# data_def2 = open('data.py', 'w')
# data_def2.write(x)

# data_abc = open('data_andi_budi_caca.txt', 'r')
# text = (data_abc.read()).replace(',','')
# print(text)
# text_list = text.split()
# print(text_list)
# text_list.sort(reverse=True)
# print(text_list)

# output = open('data_andi_budi_caca_reverse.txt', 'w')
# for i in text_list:
#     output.write(i + '\n')

# ----- Menjadikan sebuah csv ke dalam dictionary dalam python
# data_diri = open('data_diri.csv', 'r')
# x = data_diri.read().split('\n')

# list_header = x[0]
# list_header = list_header.split(';')
# list_content = x[1:4]

# content = []
# for i in list_content:
#     a = i.split(';')
#     content.append(a)

# data_diri_dict = []
# for i in range(len(list_header)):
#     y = dict(zip(list_header,content[0::][i]))
#     data_diri_dict.append(y)
# print(data_diri_dict)

# ----- menjadikan json list, menjadi dictionary python
# json = open('data_diri.json', 'r')
# print(json.read())

# ----- USING CSV MODULE -----
# import csv

# list_data = []
# with open('data_diri.csv', 'r') as x:
#     baca = csv.DictReader(x)
#     for data in baca:
#         list_data.append(dict(data))
# print(list_data)

# data = [
#     {'nama': 'Andi', 'usia':22, 'kota':'Jakarta'},
#     {'nama': 'Budi', 'usia':22, 'kota':'Bandung'},
#     {'nama': 'Caca', 'usia':22, 'kota':'Jakarta'}
# ]

# with open('data_baru.csv', 'w', newline='') as new_data:
#     kolom = ['nama', 'usia', 'kota']
#     tulis = csv.DictWriter(new_data, fieldnames=kolom)
#     tulis.writeheader()
#     tulis.writerows(data)

# ----- USING JSON MODULE -----
import json
with open('data_diri.json', 'r') as dataku:
    output = json.load(dataku)

print(output)

with open('data_diri_copy.json', 'w') as data_copy:
    json.dump(output, data_copy)