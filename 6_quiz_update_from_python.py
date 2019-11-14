import xlrd
import xlsxwriter

# =========== READ FILE ===========
update_file = xlrd.open_workbook('input_file.xlsx')
sheet_old = update_file.sheet_by_name('data_pelanggan')

# =========== PRINT OLD DATA AGAIN. BECAUSE IF NOT, NEW DATA WILL OVERWRITE THE OLD ONE ===========
old_data = []
for row in range(sheet_old.nrows):
    old_data.append(sheet_old.row_values(row))

old_data_header = old_data[0]
old_data_value = old_data[1:]

input_data = xlsxwriter.Workbook('input_file.xlsx')
sheet1 = input_data.add_worksheet('data_pelanggan')

for i in range(len(old_data_value)+1):
    for j in range(len(old_data_header)):
        sheet1.write(i,j,old_data[i][j])

# =========== THIS TIME, ASK USER WHETHER THEY WANT TO ADD NEW DATA OR NOT ===========
answer = input('Do you want to input data?(y/n): ').lower()

# =========== IN HERE, WE GRAB THE NUMBER OF THE LAST ROW BY SUBTRACTING OLD NROWS BY ONE ===========
n = sheet_old.nrows-1

# =========== IN HERE, WE GRAB THE LAST ROW THAT HAS BEEN USED, SO THE DATA WON'T
# OVERWRITE THE LAST ROW FROM OLD FILE ===========
r = sheet_old.nrows

while answer == 'y':
    c = 0
    n += 1
    nama = input('Input name: ')
    pekerjaan = input('Input occupancy: ')
    list_of_input = [n,nama,pekerjaan]
    for i in range(0,len(list_of_input)):
        sheet1.write(r,c,list_of_input[i])
        c += 1
    r += 1
    answer = input('Do you want to input data?(y/n): ').lower()
else:
    # =========== IF USER DOESN'T WANT TO ADD NEW DATA, THEN CLOSE() =========== 
    input_data.close()