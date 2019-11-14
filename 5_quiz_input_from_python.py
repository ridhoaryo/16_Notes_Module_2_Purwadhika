import xlsxwriter

input_data = xlsxwriter.Workbook('input_file.xlsx')
sheet1 = input_data.add_worksheet('data_pelanggan')
sheet1.write(0,0,'No')
sheet1.write(0,1,'Nama')
sheet1.write(0,2,'Pekerjaan')

answer = input('Do you want to input data?(y/n): ').lower()
n = 0
r = 1
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
    input_data.close()
