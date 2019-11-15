# Note and Quiz Day 16 Purwadhika Module 2 Job Connector Data Science

Pada repo ini, saya akan menjelaskan tentang mini project saya, yaitu memasukkan data baru dan mengupdate data ke Excel melalui Python.

## 1. Membuat file excel awal
```
import xlsxwriter
```
Sebelumnya kita perlu untuk mengimport `xlsxwriter` terlebih dahulu. Xlsxwriter adalah library agar kita dapat menulis ke dalam file Excel.

```
input_data = xlsxwriter.Workbook('input_file.xlsx')
sheet1 = input_data.add_worksheet('data_pelanggan')
sheet1.write(0,0,'No')
sheet1.write(0,1,'Nama')
sheet1.write(0,2,'Pekerjaan')
```

Kemudian kita buat Workbooknya dengan file `input_file.xlsx`. Setelah itu kita buat object `sheet1` sebagai wadah kita untuk menulis data nantinya. `add_worksheet` adalah method untuk membuat sheet baru, kita akan membuat sheet dengan nama `data_pelanggan`.

Setelah itu, kita membuat di dalam sheet `data_pelanggan` nama kolom terlebih dahulu, yaitu 'No', 'Nama', 'Pekerjaan'.

```
answer = input('Do you want to input data?(y/n): ').lower()
n = 0
r = 1
```
Pada step ini, kita minta input kepada user apakah ingin menambah data baru, sambil mendeklarasi bahwa `n = 0` dan `r = 1`. `n` adalah nomer terakhir yang terpakai. Dan `r` adalah row terakhir yang terpakai.

```
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
```

Jika jawaban user `y` maka kita deklarasi bahwa:
- `c` = kolom dimulai dari 0
- `n += 1` = nomor di tambah 1
- input nama = masukkan nama orang
- input pekerjaan = masukkan pekerjaan orang
- list_of_input = masukkan n, nama dan pekerjaan ke dalam satu list.
- `sheet1.write(r,c,list_of_input[i]` = masukkan data pada row saat itu, column saat itu dan data dari list_of_input dengan index `i` ke dalam cell yang dituju.
- Kemudian tambahkan column dengan 1, keluar dari for loop dan tambahkan row dengan 1 juga.
- tanyakan apakah user ingin menambah data lagi?
- jika sudah tidak ada, `input_data.close()` untuk menyimpan data awal.

## 2. Membuat file untuk mengupdate
```
update_file = xlrd.open_workbook('input_file.xlsx')
sheet_old = update_file.sheet_by_name('data_pelanggan')
```
Baca terlebih dahulu file yang lama.

```
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
```

Print ulang file yang lama. Karena jika kita langsung menulis data baru, data lama akan hilang.

```
answer = input('Do you want to input data?(y/n): ').lower()
n = sheet_old.nrows-1
r = sheet_old.nrows
```
Seperti langkah awal, kita tanyakan apakah user ingin memasukkan data baru lagi, sambil kita mendeklarasi n dari `nrows-1` dan `nrows`

```
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
```

Selebihnya sama dengan langkah ketika memasukkan data awal.