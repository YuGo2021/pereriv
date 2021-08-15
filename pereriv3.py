import re
import datetime as dt
import time
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Border, Side

def fio(fio_cell):
    # преобразование ФИО в Ф И.О.
    if fio_cell == None:
         return ""
    else:
        fio_regex = r"([А-ЯЁ]{1}[а-яё]+)\s{1}(\(([А-ЯЁ]{1}[а-яё]+)\)\s{1})?([А-ЯЁ]{1}[а-яё]+)(\s{1}([А-ЯЁ]{1}[а-яё]+))?"
        fio_test = re.search(fio_regex, fio_cell)
        fio_new = ""
        if fio_test is not None:
            #print(fio_test.group(0))
            #print(fio_test.group(1))
            #print(fio_test.group(2))
            #print(fio_test.group(3))
            #print(fio_test.group(4))
            
            if fio_test.group(6) is not None:
                fio_new = fio_test.group(1)+" "+fio_test.group(4)[0]+". "+fio_test.group(6)[0]+"."
            else:
                fio_new = fio_test.group(1)+" "+fio_test.group(4)[0]+"."
            #print(fio_new)
    return fio_new
    
def oper_div(sheet_name): # поиск в столбце с ФИО разделителей по сменам
    for i in range(1, sheet_name.max_row + 1):
        if sheet_name.cell(row = i, column = 6).value == "ОПЕРАТОРЫ 5/2":
            i_5_2 = i
        #print(i_5_2)
        if sheet_name.cell(row = i, column = 6).value== "ОПЕРАТОРЫ 2/2":
            i_2_2 = i
        #print(i_2_2)
        if sheet_name.cell(row = i, column = 6).value== "НЕПОЛНЫЙ РАБОЧИЙ ДЕНЬ" or sheet_name.cell(row = i, column = 6).value == "ОПЕРАТОРЫ  неполного рабочего  дня":
            i_1_2 = i
        #print(i_1_2)
    return(i_5_2, i_2_2, i_1_2)
    
def fio_full(fio_cell):
    # преобразование ФИО в Ф И.О.
    if fio_cell == None:
         return ""
    else:
        fio_regex = r"([А-ЯЁ]{1}[а-яё]+)\s{1}(\(([А-ЯЁ]{1}[а-яё]+)\)\s{1})?([А-ЯЁ]{1}[а-яё]+)(\s{1}([А-ЯЁ]{1}[а-яё]+))?"
        fio_test = re.search(fio_regex, fio_cell)
        fio_new = ""
        if fio_test is not None:
            #print(fio_test.group(0))
            #print(fio_test.group(1))
            #print(fio_test.group(2))
            #print(fio_test.group(3))
            #print(fio_test.group(4))
            
            if fio_test.group(6) is not None:
                fio_new = fio_test.group(0)
            else:
                fio_new = fio_test.group(1)+" "+fio_test.group(4)
            #print(fio_
    return fio_new

def shift(shift_cell):
    time_regex = r"((([0]{1})?([1-9]{1})\:([0-9]{2}))|(([1-2]{1}[0-9]{1})\:([0-9]{2})))\s?-\s?((([0]{1})?([1-9]{1})\:([0-9]{2}))|((([1-2]{1})([0-9]{1}))\:([0-9]{2})))"
    shift_test = re.search(time_regex, shift_cell)
    shift_start = ""
    shift_end = ""
    if shift_test is not None:
        if shift_test.group(4) is not None: 
            shift_start = f"{shift_test.group(4)}:{shift_test.group(5)}"
        else:
            shift_start = shift_test.group(1)
        
        if shift_test.group(10) is not None: 
            shift_end = f"{shift_test.group(12)}:{shift_test.group(13)}"
        else:
            shift_end = shift_test.group(9)
        #print(shift_start)
        
    return (shift_start , shift_end, shift_test.group(5), shift_test.group(13))

def time_shift(cell):
    time_regex = r"(\d{1,2})[\-\:]{1}(\d{1,2})"
    time_test = re.search(time_regex, cell)
    time_hour = ""
    time_min = ""
    if time_test is not None:
        time_hour = time_test.group(1)
        time_min = time_test.group(2)
    return (time_hour, time_min)
    
def sting_no(sheet):
    for i in range(1, sheet_grafik1.max_row + 1):
        if sheet_grafik1.cell(row = i, column = 6).value == "ОПЕРАТОРЫ 5/2":
            i_5_2 = i
            #print(i_5_2)
        if sheet_grafik1.cell(row = i, column = 6).value== "ОПЕРАТОРЫ 2/2":
            i_2_2 = i
            #print(i_2_2)
        if sheet_grafik1.cell(row = i, column = 6).value== "НЕПОЛНЫЙ РАБОЧИЙ ДЕНЬ" or sheet_grafik1.cell(row = i, column = 6).value == "ОПЕРАТОРЫ  неполного рабочего  дня":
            i_1_2 = i
            #print(i_1_2)
    return (i_5_2, i_2_2, i_1_2)

def find_cell(k, sheet, CC_name, shift):
    wb_per = openpyxl.load_workbook("Исключения.xlsx")
    sheet_per = wb_per.active
    
    n1 = ""
    n2 = ""
    n3 = ""
    n4 = ""
    #n5 = ""
    n6 = 0
    n7 = 0
    color = 0
    oper = ""
    oper = fio_full(sheet.cell(row = k, column = 6).value)
    #print(oper)
    n1 = "Смирновка"
    n2 = fio(oper)
    n3 = "5/2"
    #print(type(sheet.cell(row = k, column = 5).value))
    #temp = sheet.cell(row = k, column = 5).value
    #print(shift(temp)[0])
    #n4 = shift(sheet.cell(row = k, column = 5).value)[0]
    #n5 = shift(sheet.cell(row = k, column = 5).value)[1]
    cell = sheet.cell(row = k, column = 9).value
    cell_plus_1 = sheet.cell(row = k+1, column = 9).value
    
    if (type(cell) == int or type(cell) == float) and (type(cell_plus_1) == int or type(cell_plus_1) == float):
        n4 = "с+доп"
    elif (type(cell) == int or type(cell) == float) and cell_plus_1 is None:
        n4 = "c"
    else:
        n4 = "доп"
    
    if  (type(cell) == int or type(cell) == float) and (type(cell_plus_1) == int or type(cell_plus_1) == float):
        n6 = sheet.cell(row = k, column = 9).value + sheet.cell(row = k+1, column = 9).value
    elif (type(cell) == int or type(cell) == float) and cell_plus_1 is None:
        n6 = sheet.cell(row = k, column = 9).value
    else:
        n6 = sheet.cell(row = k+1, column = 9).value
 
    
#     if sheet.cell(row = k, column = 9).value is not None and sheet.cell(row = k+1, column = 9).value is not None:
#         n4 = "с+доп"
#     elif sheet.cell(row = k, column = 9).value is not None and sheet.cell(row = k+1, column = 9).value is None:
#         n4 = "c"
#     else:
#         n4 = "доп"
    
#     if sheet.cell(row = k, column = 9).value is not None and sheet.cell(row = k+1, column = 9).value is not None:
#         n6 = sheet.cell(row = k, column = 9).value + sheet.cell(row = k+1, column = 9).value
#     elif sheet.cell(row = k, column = 9).value is not None and sheet.cell(row = k+1, column = 9).value is None:
#         n6 = sheet.cell(row = k, column = 9).value
#     else:
#         n6 = sheet.cell(row = k+1, column = 9).value
    
    if n4 == "c":
        if n6 == 11:
            n7 = 75
            
        elif n6 > 11:
            n7 = 75
            color = 1
        elif n6 == 8:
            n7 = 65
        elif 8< n6 < 11:
            n7 = 70
            color = 1
        elif n6 == 3:
            n7 = 15
            color = 1
        elif 3 < n6 < 6:
            n7 = 20
            color = 1
        elif n6 == 6:
            n7 = 30
            color = 1
        elif 6 < n6 < 8:
            n7 = 60
            color = 1
        else:
            n7 = ""
            color = 1
            
    else:
        if n6 == 11:
            n7 = 75
        elif 8 < n6 < 11:
            n7 = 70
            
        elif n6 > 11:
            n7 = 75
            
        elif n6 == 8:
            n7 = 65
        elif 3 < n6 < 6:
            n7 = 20
            
        elif n6 == 3:
            n7 = 15
            
        elif n6 == 6:
            n7 = 30
            
        elif 6 < n6 < 8:
            n7 = 60
            
        else:
            n7 = ""
        color = 1
    for i in range(2, sheet_per.max_row + 1):
        if n2 == sheet_per.cell(row = i, column = 1).value:
            n7 = sheet_per.cell(row = i, column = 2).value
            if n4 == "c":
                color = 0
            else:
                color = 1
    return (n1,n2,n3,n4,n6,n7,oper,color)

def sum_cell(line1,line2,sheet):
    sum_lines = 0
    for z in range(4, sheet.max_column + 1):
        for line in range(line1, line2):
            if sheet.cell(row = line, column = z).value is not None:
                sum_lines += sheet.cell(row=line, column=z).value
    return sum_lines

def time_chek(cell):
    #print(type(cell))
    if type(cell) == dt.time:
        time_start = cell.strftime("%H:%M")
        #print(time_start)
        if time_start[0] == "0":
            time_start = time_start[1:]
    else:    
        time_start = cell
        if time_start[5:7] == ":00":
            time_start = time_start[0:4]
        
        if time_start[0] == "0":
            time_start = time_start[1:]
    return time_start


    
wb_grafik_per = openpyxl.load_workbook(f"перерывы_сборка.xlsx")
for s in range(len(wb_grafik_per.sheetnames)):
    if wb_grafik_per.sheetnames[s] == 'перерывы':
        break
wb_grafik_per.active = s

print("Готовим таблицы для рассылки...")
sheet_per = wb_grafik_per.active

thin = Side(border_style="thin", color="000000")
double = Side(border_style="double", color="ff0000")
medium = Side(border_style="medium", color="000000")

if wb_grafik_per.sheetnames.count('Смирновка') == 0:
    wb_grafik_per.create_sheet(title = 'Смирновка', index = 0)
sheet_sm = wb_grafik_per['Смирновка']
sheet_sm['A1'] = "ФИО"
sheet_sm['B1'] = "Время работы"
sheet_sm['C1'] = "Перерыв 1"
sheet_sm['D1'] = "Перерыв 2"
sheet_sm['E1'] = "Перерыв 3"
sheet_sm['F1'] = "Перерыв 4"
sheet_sm['G1'] = "Перерыв 5"

if wb_grafik_per.sheetnames.count('Высота') == 0:
    wb_grafik_per.create_sheet(title = 'Высота', index = 0)
sheet_vis = wb_grafik_per['Высота']
sheet_vis['A1'] = "ФИО"
sheet_vis['B1'] = "Время работы"
sheet_vis['C1'] = "Перерыв 1"
sheet_vis['D1'] = "Перерыв 2"
sheet_vis['E1'] = "Перерыв 3"
sheet_vis['F1'] = "Перерыв 4"
sheet_vis['G1'] = "Перерыв 5"

if wb_grafik_per.sheetnames.count('Киров') == 0:
    wb_grafik_per.create_sheet(title = 'Киров', index = 0)
sheet_kir = wb_grafik_per['Киров']
sheet_kir['A1'] = "ФИО"
sheet_kir['B1'] = "Время работы"
sheet_kir['C1'] = "Перерыв 1"
sheet_kir['D1'] = "Перерыв 2"
sheet_kir['E1'] = "Перерыв 3"
sheet_kir['F1'] = "Перерыв 4"
sheet_kir['G1'] = "Перерыв 5"

if wb_grafik_per.sheetnames.count('НСК') == 0:
    wb_grafik_per.create_sheet(title = 'НСК', index = 0)
sheet_nsk = wb_grafik_per['НСК']
sheet_nsk['A1'] = "ФИО"
sheet_nsk['B1'] = "Время работы"
sheet_nsk['C1'] = "Перерыв 1"
sheet_nsk['D1'] = "Перерыв 2"
sheet_nsk['E1'] = "Перерыв 3"
sheet_nsk['F1'] = "Перерыв 4"
sheet_nsk['G1'] = "Перерыв 5"

if wb_grafik_per.sheetnames.count('Ростов') == 0:
    wb_grafik_per.create_sheet(title = 'Ростов', index = 0)
sheet_rost = wb_grafik_per['Ростов']
sheet_rost['A1'] = "ФИО"
sheet_rost['B1'] = "Время работы"
sheet_rost['C1'] = "Перерыв 1"
sheet_rost['D1'] = "Перерыв 2"
sheet_rost['E1'] = "Перерыв 3"
sheet_rost['F1'] = "Перерыв 4"
sheet_rost['G1'] = "Перерыв 5"

if wb_grafik_per.sheetnames.count('НиНо') == 0:
    wb_grafik_per.create_sheet(title = 'НиНо', index = 0)
sheet_nn = wb_grafik_per['НиНо']
sheet_nn['A1'] = "ФИО"
sheet_nn['B1'] = "Время работы"
sheet_nn['C1'] = "Перерыв 1"
sheet_nn['D1'] = "Перерыв 2"
sheet_nn['E1'] = "Перерыв 3"
sheet_nn['F1'] = "Перерыв 4"
sheet_nn['G1'] = "Перерыв 5"

# формируем рассылку для Смирновки
no = 2
for kol in range(4,sheet_per.max_column + 1 ):
     
    if sheet_per.cell(row = 1, column = kol).value == 'Смирновка':
        sheet_sm.cell(row = no, column = 1).value = sheet_per.cell(row = 2, column = kol).value
        sheet_sm.cell(row = no, column = 2).value = f"{time_chek(sheet_per.cell(row = 4, column = kol).value)}-{time_chek(sheet_per.cell(row = 5, column = kol).value)}"
        per_berin = 10
        per_end = 10
        no_sm = 3
        for st in range(10, sheet_per.max_row + 1):
            if sheet_per.cell(row = st, column = kol).value == 5 and st > per_end:
                per_berin = st
                for st_per in range(st, sheet_per.max_row + 1):
                    if sheet_per.cell(row = st_per, column = kol).value != 5:
                        per_end = st_per - 1
                        break
                #print(per_berin, per_end)
                for st_beg in range(per_berin, per_berin - 12, -1):
                    if sheet_per.cell(row = st_beg, column = 1).value is not None:
                        hour_begin = time_shift(shift(sheet_per.cell(row = st_beg, column = 1).value)[0])[0]
                        break
                             
                min_begin = time_shift(sheet_per.cell(row = per_berin, column = 2).value)[0]
                for st_end in range(per_end, per_end - 12, -1):
                    if sheet_per.cell(row = st_end, column = 1).value is not None:
                        hour_end = time_shift(shift(sheet_per.cell(row = st_end, column = 1).value)[0])[0]
                        break
                            
                min_end = time_shift(sheet_per.cell(row = per_end, column = 2).value)[1]
                if min_end == "60":
                    min_end = "00"
                    hour_end = int(hour_end) + 1 
                elif min_end == "0":
                    min_end = "00"
                elif min_end == "5":
                    min_end = "05"
                
                if min_begin == "0":
                    min_begin = "00"
                elif min_begin == "5":
                    min_begin = "05"
                elif min_begin == "60":
                    min_begin = "00"
                    hour_begin = int(hour_begin) + 1 
                
                sheet_sm.cell(row = no, column = no_sm).value = f"{hour_begin}:{min_begin}-{hour_end}:{min_end}"
                no_sm += 1
        no += 1
for i in range(1, sheet_sm.max_column + 1):
    for j in range(1, sheet_sm.max_row + 1):
        sheet_sm.cell(row = j, column = i).border = Border(top=thin, left=thin, right=thin, bottom=thin)

# формируем рассылку для Высоты
no = 2
for kol in range(4,sheet_per.max_column + 1 ):
     
    if sheet_per.cell(row = 1, column = kol).value == 'Высота':
        sheet_vis.cell(row = no, column = 1).value = sheet_per.cell(row = 2, column = kol).value
        sheet_vis.cell(row = no, column = 2).value = f"{time_chek(sheet_per.cell(row = 4, column = kol).value)}-{time_chek(sheet_per.cell(row = 5, column = kol).value)}"
        per_berin = 10
        per_end = 10
        no_sm = 3
        for st in range(10, sheet_per.max_row + 1):
            if sheet_per.cell(row = st, column = kol).value == 5 and st > per_end:
                per_berin = st
                for st_per in range(st, sheet_per.max_row + 1):
                    if sheet_per.cell(row = st_per, column = kol).value != 5:
                        per_end = st_per - 1
                        break
                #print(per_berin, per_end)
                for st_beg in range(per_berin, per_berin - 12, -1):
                    if sheet_per.cell(row = st_beg, column = 1).value is not None:
                        hour_begin = time_shift(shift(sheet_per.cell(row = st_beg, column = 1).value)[0])[0]
                        break
                             
                min_begin = time_shift(sheet_per.cell(row = per_berin, column = 2).value)[0]
                for st_end in range(per_end, per_end - 12, -1):
                    if sheet_per.cell(row = st_end, column = 1).value is not None:
                        hour_end = time_shift(shift(sheet_per.cell(row = st_end, column = 1).value)[0])[0]
                        break
                            
                min_end = time_shift(sheet_per.cell(row = per_end, column = 2).value)[1]
                if min_end == "60":
                    min_end = "00"
                    hour_end = int(hour_end) + 1 
                elif min_end == "0":
                    min_end = "00"
                elif min_end == "5":
                    min_end = "05"
                
                if min_begin == "0":
                    min_begin = "00"
                elif min_begin == "5":
                    min_begin = "05"
                elif min_begin == "60":
                    min_begin = "00"
                    hour_begin = int(hour_begin) + 1 
                
                sheet_vis.cell(row = no, column = no_sm).value = f"{hour_begin}:{min_begin}-{hour_end}:{min_end}"
                no_sm += 1
        no += 1
for i in range(1, sheet_vis.max_column + 1):
    for j in range(1, sheet_vis.max_row + 1):
        sheet_vis.cell(row = j, column = i).border = Border(top=thin, left=thin, right=thin, bottom=thin)        

# формируем рассылку для Кирова
no = 2
for kol in range(4,sheet_per.max_column + 1 ):
     
    if sheet_per.cell(row = 1, column = kol).value == 'Киров':
        sheet_kir.cell(row = no, column = 1).value = sheet_per.cell(row = 2, column = kol).value
        sheet_kir.cell(row = no, column = 2).value = f"{sheet_per.cell(row = 4, column = kol).value}-{sheet_per.cell(row = 5, column = kol).value}"
        per_berin = 10
        per_end = 10
        no_sm = 3
        for st in range(10, sheet_per.max_row + 1):
            if sheet_per.cell(row = st, column = kol).value == 5 and st > per_end:
                per_berin = st
                for st_per in range(st, sheet_per.max_row + 1):
                    if sheet_per.cell(row = st_per, column = kol).value != 5:
                        per_end = st_per - 1
                        break
                #print(per_berin, per_end)
                for st_beg in range(per_berin, per_berin - 12, -1):
                    if sheet_per.cell(row = st_beg, column = 1).value is not None:
                        hour_begin = time_shift(shift(sheet_per.cell(row = st_beg, column = 1).value)[0])[0]
                        break
                             
                min_begin = time_shift(sheet_per.cell(row = per_berin, column = 2).value)[0]
                for st_end in range(per_end, per_end - 12, -1):
                    if sheet_per.cell(row = st_end, column = 1).value is not None:
                        hour_end = time_shift(shift(sheet_per.cell(row = st_end, column = 1).value)[0])[0]
                        break
                            
                min_end = time_shift(sheet_per.cell(row = per_end, column = 2).value)[1]
                if min_end == "60":
                    min_end = "00"
                    hour_end = int(hour_end) + 1 
                elif min_end == "0":
                    min_end = "00"
                elif min_end == "5":
                    min_end = "05"
                
                if min_begin == "0":
                    min_begin = "00"
                elif min_begin == "5":
                    min_begin = "05"
                elif min_begin == "60":
                    min_begin = "00"
                    hour_begin = int(hour_begin) + 1 
                
                sheet_kir.cell(row = no, column = no_sm).value = f"{hour_begin}:{min_begin}-{hour_end}:{min_end}"
                no_sm += 1
        no += 1
for i in range(1, sheet_kir.max_column + 1):
    for j in range(1, sheet_kir.max_row + 1):
        sheet_kir.cell(row = j, column = i).border = Border(top=thin, left=thin, right=thin, bottom=thin)

# формируем рассылку для Новосибирска
no = 2
for kol in range(4,sheet_per.max_column + 1 ):
     
    if sheet_per.cell(row = 1, column = kol).value == 'НСК':
        sheet_nsk.cell(row = no, column = 1).value = sheet_per.cell(row = 2, column = kol).value
        sheet_nsk.cell(row = no, column = 2).value = f"{int(time_shift(sheet_per.cell(row = 4, column = kol).value)[0])+4}:{time_shift(sheet_per.cell(row = 4, column = kol).value)[1]}-{int(time_shift(sheet_per.cell(row = 5, column = kol).value)[0])+4}:{time_shift(sheet_per.cell(row = 5, column = kol).value)[1]}"
        per_berin = 10
        per_end = 10
        no_sm = 3
        for st in range(10, sheet_per.max_row + 1):
            if sheet_per.cell(row = st, column = kol).value == 5 and st > per_end:
                per_berin = st
                for st_per in range(st, sheet_per.max_row + 1):
                    if sheet_per.cell(row = st_per, column = kol).value != 5:
                        per_end = st_per - 1
                        break
                #print(per_berin, per_end)
                for st_beg in range(per_berin, per_berin - 12, -1):
                    if sheet_per.cell(row = st_beg, column = 1).value is not None:
                        hour_begin = time_shift(shift(sheet_per.cell(row = st_beg, column = 1).value)[0])[0]
                        break
                             
                min_begin = time_shift(sheet_per.cell(row = per_berin, column = 2).value)[0]
                for st_end in range(per_end, per_end - 12, -1):
                    if sheet_per.cell(row = st_end, column = 1).value is not None:
                        hour_end = time_shift(shift(sheet_per.cell(row = st_end, column = 1).value)[0])[0]
                        break
                            
                min_end = time_shift(sheet_per.cell(row = per_end, column = 2).value)[1]
                if min_end == "60":
                    min_end = "00"
                    hour_end = int(hour_end) + 1 
                elif min_end == "0":
                    min_end = "00"
                elif min_end == "5":
                    min_end = "05"
                
                if min_begin == "0":
                    min_begin = "00"
                elif min_begin == "5":
                    min_begin = "05"
                elif min_begin == "60":
                    min_begin = "00"
                    hour_begin = int(hour_begin) + 1 
                
                sheet_nsk.cell(row = no, column = no_sm).value = f"{int(hour_begin)+4}:{min_begin}-{int(hour_end)+4}:{min_end}"
                no_sm += 1
        no += 1
for i in range(1, sheet_nsk.max_column + 1):
    for j in range(1, sheet_nsk.max_row + 1):
        sheet_nsk.cell(row = j, column = i).border = Border(top=thin, left=thin, right=thin, bottom=thin)

# формируем рассылку для Ростов
no = 2
for kol in range(4,sheet_per.max_column + 1 ):
     
    if sheet_per.cell(row = 1, column = kol).value == 'Ростов':
        sheet_rost.cell(row = no, column = 1).value = sheet_per.cell(row = 2, column = kol).value
        sheet_rost.cell(row = no, column = 2).value = f"{sheet_per.cell(row = 4, column = kol).value}-{sheet_per.cell(row = 5, column = kol).value}"
        per_berin = 10
        per_end = 10
        no_sm = 3
        for st in range(10, sheet_per.max_row + 1):
            if sheet_per.cell(row = st, column = kol).value == 5 and st > per_end:
                per_berin = st
                for st_per in range(st, sheet_per.max_row + 1):
                    if sheet_per.cell(row = st_per, column = kol).value != 5:
                        per_end = st_per - 1
                        break
                #print(per_berin, per_end)
                for st_beg in range(per_berin, per_berin - 12, -1):
                    if sheet_per.cell(row = st_beg, column = 1).value is not None:
                        hour_begin = time_shift(shift(sheet_per.cell(row = st_beg, column = 1).value)[0])[0]
                        break
                             
                min_begin = time_shift(sheet_per.cell(row = per_berin, column = 2).value)[0]
                for st_end in range(per_end, per_end - 12, -1):
                    if sheet_per.cell(row = st_end, column = 1).value is not None:
                        hour_end = time_shift(shift(sheet_per.cell(row = st_end, column = 1).value)[0])[0]
                        break
                            
                min_end = time_shift(sheet_per.cell(row = per_end, column = 2).value)[1]
                if min_end == "60":
                    min_end = "00"
                    hour_end = int(hour_end) + 1 
                elif min_end == "0":
                    min_end = "00"
                elif min_end == "5":
                    min_end = "05"
                
                if min_begin == "0":
                    min_begin = "00"
                elif min_begin == "5":
                    min_begin = "05"
                elif min_begin == "60":
                    min_begin = "00"
                    hour_begin = int(hour_begin) + 1 
                
                sheet_rost.cell(row = no, column = no_sm).value = f"{hour_begin}:{min_begin}-{hour_end}:{min_end}"
                no_sm += 1
        no += 1
for i in range(1, sheet_rost.max_column + 1):
    for j in range(1, sheet_rost.max_row + 1):
        sheet_rost.cell(row = j, column = i).border = Border(top=thin, left=thin, right=thin, bottom=thin)
 
# формируем рассылку для Нижнего Новгорода
no = 2
for kol in range(4,sheet_per.max_column + 1 ):
     
    if sheet_per.cell(row = 1, column = kol).value == 'НиНо':
        sheet_nn.cell(row = no, column = 1).value = sheet_per.cell(row = 2, column = kol).value
        sheet_nn.cell(row = no, column = 2).value = f"{sheet_per.cell(row = 4, column = kol).value}-{sheet_per.cell(row = 5, column = kol).value}"
        per_berin = 10
        per_end = 10
        no_sm = 3
        for st in range(10, sheet_per.max_row + 1):
            if sheet_per.cell(row = st, column = kol).value == 5 and st > per_end:
                per_berin = st
                for st_per in range(st, sheet_per.max_row + 1):
                    if sheet_per.cell(row = st_per, column = kol).value != 5:
                        per_end = st_per - 1
                        break
                #print(per_berin, per_end)
                for st_beg in range(per_berin, per_berin - 12, -1):
                    if sheet_per.cell(row = st_beg, column = 1).value is not None:
                        hour_begin = time_shift(shift(sheet_per.cell(row = st_beg, column = 1).value)[0])[0]
                        break
                             
                min_begin = time_shift(sheet_per.cell(row = per_berin, column = 2).value)[0]
                for st_end in range(per_end, per_end - 12, -1):
                    if sheet_per.cell(row = st_end, column = 1).value is not None:
                        hour_end = time_shift(shift(sheet_per.cell(row = st_end, column = 1).value)[0])[0]
                        break
                            
                min_end = time_shift(sheet_per.cell(row = per_end, column = 2).value)[1]
                if min_end == "60":
                    min_end = "00"
                    hour_end = int(hour_end) + 1 
                elif min_end == "0":
                    min_end = "00"
                elif min_end == "5":
                    min_end = "05"
                
                if min_begin == "0":
                    min_begin = "00"
                elif min_begin == "5":
                    min_begin = "05"
                elif min_begin == "60":
                    min_begin = "00"
                    hour_begin = int(hour_begin) + 1 
                
                sheet_nn.cell(row = no, column = no_sm).value = f"{hour_begin}:{min_begin}-{hour_end}:{min_end}"
                no_sm += 1
        no += 1
for i in range(1, sheet_nn.max_column + 1):
    for j in range(1, sheet_nn.max_row + 1):
        sheet_nn.cell(row = j, column = i).border = Border(top=thin, left=thin, right=thin, bottom=thin) 

 
try:
    wb_grafik_per.save(f"перерывы_сборка.xlsx") 
    input("Таблицы для рассылки подготовлены. Открываем файл перерывы_сборка.xlsx и проверяем. Нажмите ENTER для продолжения...") 
except OSError:
    input("Невозможно сохранить данные. Закройте файл перерывы_сборка.xlsx и запустите программу заново. Нажмите ENTER для продолжения...")