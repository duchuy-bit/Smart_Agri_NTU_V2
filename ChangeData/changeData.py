
import gspread
import csv
from gspread_formatting import *
import datetime
# from gspread_formatting import *

#-----------------------------EXCEL----------------------------

date = []
time = []
temperature1 = []
temperature2 = []
temperature3 = []
humidity1 = []
humidity2 = []
humidity3 = []
light1 = []
light2 = []
light3 = []
line_count = 0

#      -----  đếm số row   ------- 
with open('dataset.csv') as csv_file:
    print('  ------- loading excel success')
    csv_reader = csv.reader(csv_file, delimiter=',')
    for row in csv_reader:        
        # dừng đọc khi dữ liệu của cell ở row đó = rỗng
        if row[1] == '':
            print(' ---  file null')
            break     
        line_count += 1
    print(f' ----- Processed {line_count} lines.')

length = line_count
print(' ---- length=')
print(length)

        # =-=-=-=-=-=-=-=----- lấy dữ liệu 10 dòng cuối --------=-=-=-=-=-=-=
n = 0
with open('dataset.csv') as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=',')
    line_count = 0
    for row in csv_reader:

        # dừng đọc khi dữ liệu của cell ở row đó = rỗng
        if row[1] == '':
            break

        if line_count == length:
            break    

        line_count += 1

        if line_count > length - length + 1:            
            date.append(row[1])
            time.append(row[2])
            humidity1.append(row[3])
            humidity2.append(row[4])
            humidity3.append(row[5])
            temperature1.append(row[6])
            temperature2.append(row[7])
            temperature3.append(row[8])
            light1.append(row[9])
            light2.append(row[10])
            light3.append(row[11])
            n +=1

print('n=')
print(n)

i=0
while i < n:
    # temperature1[i] = float(temperature1[i])
    # temperature2[i] = float(temperature2[i])
    # temperature3[i] = float(temperature3[i])
    # humidity1[i] = float(humidity1[i])
    # humidity2[i] = float(humidity2[i])
    # humidity3[i] = float(humidity3[i])
    # light1[i] = float(light1[i])
    # light2[i] = float(light2[i])
    # light3[i] = float(light3[i])
    
    # temperature1[i] = str(temperature1[i])
    # temperature2[i] = str(temperature2[i])
    # temperature3[i] = str(temperature3[i])
    # humidity1[i] = str(humidity1[i])
    # humidity2[i] = str(humidity2[i])
    # humidity3[i] = str(humidity3[i])
    # light1[i] = str(light1[i])
    # light2[i] = str(light2[i])
    # light3[i] = str(light3[i])
    print(f'{date[i]} {time[i]} {humidity1[i]} {humidity2[i]} {humidity3[i]} : {temperature1[i]} {temperature2[i]} {temperature3[i]} :  {light1[i]} {light2[i]} {light3[i]}')
    i+=1

# -----------------------------Google Sheet---------------
gs = gspread.service_account("gsheet.json")
sht = gs.open_by_key('1KsE0-DmUPqbbJKJVLwpHzm3BNssRZH7Enmvi6OkVF7k')
worksheet = sht.get_worksheet(0)

# print(worksheet.row_count)

list_of_lists = worksheet.get_all_values()


length_list = len(list_of_lists)

i=0
checked = 0
print('-----')
print(str(list_of_lists[len(list_of_lists) - 1]))
while i < n:
    # temperature[i] = float(temperature[i])
    # humidity[i] = float(humidity[i])
    # light[i] = float(light[i])
    
    # temperature1[i] = float(temperature1[i])
    # temperature2[i] = float(temperature2[i])
    # temperature3[i] = float(temperature3[i])
    # humidity1[i] = float(humidity1[i])
    # humidity2[i] = float(humidity2[i])
    # humidity3[i] = float(humidity3[i])
    # light1[i] = float(light1[i])
    # light2[i] = float(light2[i])
    # light3[i] = float(light3[i])
    
    insertRow = [date[i],time[i],humidity1[i],humidity2[i],humidity3[i],float(temperature1[i]),float(temperature2[i]),float(temperature3[i]),light1[i],light2[i],light3[i]]
    if str(date[i]) == list_of_lists[length_list-1][0] and str(time[i]) == list_of_lists[length_list-1][1]:
        print('trung nhau')
        checked=1
        print(insertRow)
        j = i+1
        while j < n:
            print('danh sach cac hang can chen')
            insertRow = [date[j],time[j],humidity1[j],humidity2[j],humidity3[j],temperature1[j],temperature2[j],temperature3[j],light1[j],light2[j],light3[j]]
            print(insertRow)
            length_list +=1
            worksheet.insert_row(insertRow,length_list,value_input_option="USER_ENTERED")
            j += 1
        break
    i+=1
List_Add = []
if(checked == 0):
    print('Insert All')
    # insertRow = [date[i],time[i],temperature[i],humidity[i],light[i]]
    j = 0
    while j < n:
        print('danh sach cac hang can chen')
        insertRow = [date[j],time[j],humidity1[j],humidity2[j],humidity3[j],temperature1[j],temperature2[j],temperature3[j],light1[j],light2[j],light3[j]]
        List_Add.append(insertRow)
        print(insertRow)
        length_list += 1
        # worksheet.insert_row(insertRow,length_list,value_input_option="USER_ENTERED")
        print('Length List ')
        print(length_list)
        j += 1
i=-0
length_list = 1
#xoa
# for cpfs in range(len(List_Add)):
#     i+=1
#     # insertRow = List_Add[cpfs]
#     insertRow = [date[cpfs],time[cpfs],temperature[cpfs],humidity[cpfs],light[cpfs]]
#     length_list +=1
#     worksheet.insert_row(insertRow,length_list,value_input_option="USER_ENTERED")
#     # print(List_Add[cpfs])


insertRow = [date[1],time[1]]

worksheet.insert_row(insertRow,2,value_input_option="USER_ENTERED")

sht.values_update(
    'DataSet!A2', 
    params={'value_input_option': "USER_ENTERED"}, 
    body={'values': List_Add},
    
)














# worksheet.insert_row(List_Add,length_list,value_input_option="USER_ENTERED")
# # temperature[n-1] = float(temperature[n-1])

# worksheet.update_cell(len, 1, date[n-1])


# insertRow = [date[n-1],time[n-1],temperature[n-1],humidity[n-1],light[n-1]]
# worksheet.insert_row(insertRow,len+1,value_input_option="USER_ENTERED")


# print()



# fmt = cellFormat(
#     numberFormat=numberFormat(type='NUMBER', pattern='####.#')
#     )

# format_cell_range(worksheet, 'C38:D39', fmt)



# fmt = cellFormat(numberFormat=numberFormat(type='NUMBER', pattern='####.#'))
# format_cell_range(worksheet, 'C38:C39', cell_format= NumberFormat(type="NUMBER",pattern='#####.#'))
# worksheet.update_cells('C38', value_input_option = 'USER_ENTERED')

# worksheet.format("C38",)
# row_list = list_of_lists[len-1]

# current_cell=row_list[i].replace(',', '') #remove the commas from any numbers


# worksheet.format('A38',
#     {
#         # "requests":{
#             "cell":{
#                 "userEnteredFormat": {
#             "numberFormat": {
#                 "type": "NUMBER",
#                 "pattern": "#,##0",
#             },
#             "backgroundColor": {
#                 "red": 0.0,
#                 "green": 0.4,
#                 "blue": 0.4
#             },
#             }
#             }
#         # }
#     }
# )



