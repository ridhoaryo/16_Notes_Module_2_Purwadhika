import json
import xlsxwriter
import csv
# ================ IMPORT FROM JSON TO XLSX ================
with open('file.json', 'r') as my_file:
    output = json.load(my_file)

header = list(output[0].keys())
data = []
for i in output:
    data.append(list(i.values()))
unite = []
for i in data:
    unite.append(i)
unite.insert(0,header)

# ================ IMPORT FROM CSV TO XLSX ================
list_of_data = []
with open('file_csv.csv', 'r') as csv_file:
    out_csv = csv.DictReader(csv_file)
    for d in out_csv:
        list_of_data.append(dict(d))
# print(list_of_data)
header_csv = list(list_of_data[0].keys())
data_csv = []
for a in list_of_data:
    data_csv.append(list(a.values()))
# print(data_csv)
unite_csv = []
for b in data_csv:
    unite_csv.append(b)
unite_csv.insert(0, header_csv)
# print(unite_csv)

# ================ MAKE WORKBOOK AND SHEET ================

new_file = xlsxwriter.Workbook('file_duplicate_from_json_csv.xlsx')
sheet = new_file.add_worksheet('data_orang')
sheet2 = new_file.add_worksheet('data_orang_2')
sheet3 = new_file.add_worksheet('data_orang_3')

# ================ WRITE FOR SHEET 1 ================
for i in range(len(data)+1):
    for j in range(len(header)):
        sheet.write(i,j,unite[i][j])

# ================ WRITE FOR SHEET 2 ================
for e in range(len(data_csv)+1):
    for f in range(len(header_csv)):
        sheet2.write(e,f,unite_csv[e][f])

# NOTE: IF YOU WANT TO ADD ANOTHER SHEET, REMEMBER TO CLOSE XLSX FILE FIRST
# ================ WRITE FOR SHEET 3 (SAME AS SHEET 2) ================
for g in range(len(data_csv)+1):
    for h in range(len(header_csv)):
        sheet3.write(g,h,unite_csv[g][h])


new_file.close()
