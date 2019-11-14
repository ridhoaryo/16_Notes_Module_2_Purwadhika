import xlrd
import xlsxwriter

# xlsx ==> new xlsx
opened_file = xlrd.open_workbook('file.xlsx')
sheet = opened_file.sheet_by_index(0)
data = []
for i in range(sheet.nrows):
    data.append(sheet.row_values(i))
# print(data)

header = data[0]
value = list(data[1:])

# print(header)
print(value)

new_file = xlsxwriter.Workbook('file_duplicate.xlsx')
sheet = new_file.add_worksheet('data_orang')

for i in range(len(value)+1):
    for j in range(len(header)):
        sheet.write(i,j,data[i][j])

# new_file.close()

