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

def nsk(cell):
    time_regex = r"(\d{1,2})[\-\:]{1}(\d{1,2})"
    time_test = re.search(time_regex, cell)
    time_hour = ""
    time_min = ""
    if time_test is not None:
        time_hour = str(int(time_test.group(1))-4)
        time_min = time_test.group(2)
    return (f"{time_hour}:{time_min}")

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
    n1 = CC_name
    n2 = fio(oper)
    n3 = shift
    #print(type(sheet.cell(row = k, column = 5).value))
    #temp = sheet.cell(row = k, column = 5).value
    #print(shift(temp)[0])
    #n4 = shift(sheet.cell(row = k, column = 5).value)[0]
    #n5 = shift(sheet.cell(row = k, column = 5).value)[1]
    cell = sheet.cell(row = k, column = 9).value
    cell_plus_1 = sheet.cell(row = k+1, column = 9).value
    
    if (type(cell) == int or type(cell) == float) and (type(cell_plus_1) == int or type(cell_plus_1) == float):
        n4 = "с+доп"
    elif (type(cell) == int or type(cell) == float) and not (type(cell_plus_1) == int or type(cell_plus_1) == float):
        n4 = "c"
    elif (type(cell_plus_1) == int or type(cell_plus_1) == float) and not (type(cell) == int or type(cell) == float):
        n4 = "доп"
    
    if  (type(cell) == int or type(cell) == float) and (type(cell_plus_1) == int or type(cell_plus_1) == float):
        n6 = sheet.cell(row = k, column = 9).value + sheet.cell(row = k+1, column = 9).value
    elif (type(cell) == int or type(cell) == float) and not (type(cell_plus_1) == int or type(cell_plus_1) == float):
        n6 = sheet.cell(row = k, column = 9).value
    elif (type(cell_plus_1) == int or type(cell_plus_1) == float) and not (type(cell) == int or type(cell) == float):
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

def pereriv_CC(sheet_grafik1, i_op, sheet_rez, CC_name):
    redFill = PatternFill(start_color='FFEE1111', end_color='FFEE1111', fill_type='solid')
    greenFill = PatternFill(start_color='1FB714', end_color='1FB714', fill_type='solid')
    yellowFill = PatternFill(start_color='FCF305', end_color='FCF305', fill_type='solid')
    i_5_2 = 0
    i_2_2 = 0
    i_1_2 = 0
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

    oper = ""
    if i_op < 4:
        i_op = 4


    #for k in range (sting_no(sheet_grafik1)[0],sting_no(sheet_grafik1)[1]-3):
     #   if sheet_grafik1.cell(row = k, column = 5).value is not None and sheet_grafik1.cell(row = k, column = 6).value is not None and (type(sheet_grafik1.cell(row = k, column = 9).value) == int or type(sheet_grafik1.cell(row = k, column = 9).value) == float or type(sheet_grafik1.cell(row = k+1, column = 9).value) == int or type(sheet_grafik1.cell(row = k+1, column = 9).value) == float):
            #print(find_cell(k, sheet_grafik1, CC_name, "5/2")[6])
      #      for n in range (6):
       #         sheet_rez.cell(row = n+1, column = i_op).value = find_cell(k, sheet_grafik1, CC_name, "5/2")[n]
        #    i_op += 1
    if i_5_2 > 0:
        for k in range(i_5_2+1, i_2_2-3): 
            if sheet_grafik1.cell(row = k, column = 5).value is not None and sheet_grafik1.cell(row = k, column = 6).value is not None and (type(sheet_grafik1.cell(row = k, column = 9).value) == int or type(sheet_grafik1.cell(row = k, column = 9).value) == float or type(sheet_grafik1.cell(row = k+1, column = 9).value) == int or type(sheet_grafik1.cell(row = k+1, column = 9).value) == float):
                oper = fio_full(sheet_grafik1.cell(row = k, column = 6).value)
                print(oper)
                sheet_rez.cell(row = 1, column = i_op).value = CC_name
                sheet_rez.cell(row = 2, column = i_op).value = fio(oper)
                sheet_rez.cell(row = 3, column = i_op).value = "5/2"
                sheet_rez.cell(row = 4, column = i_op).value = shift(sheet_grafik1.cell(row = k, column = 5).value)[0]
                #print(shift(sheet_grafik1.cell(row = k, column = 5).value)[0])
                sheet_rez.cell(row = 5, column = i_op).value = shift(sheet_grafik1.cell(row = k, column = 5).value)[1]
                #print(shift(sheet_grafik1.cell(row = k, column = 5).value)[1])
                sheet_rez.cell(row = 6, column = i_op).value = find_cell(k, sheet_grafik1, CC_name, "5/2")[3]
                sheet_rez.cell(row = 7, column = i_op).value = find_cell(k, sheet_grafik1, CC_name, "5/2")[4]
                sheet_rez.cell(row = 8, column = i_op).value = find_cell(k, sheet_grafik1, CC_name, "5/2")[5]
                if find_cell(k, sheet_grafik1, CC_name, "5/2")[7] == 1:
                    sheet_rez.cell(row = 8, column = i_op).fill = redFill
                sheet_rez.cell(row = 9, column = i_op).value = f"=SUM({get_column_letter(i_op)}10:{get_column_letter(i_op)}273)"
                #if sheet_grafik1.cell(row = k, column = 9).value is not None and sheet_grafik1.cell(row = k+1, column = 9).value is not None:
                 #   sheet_rez.cell(row = 6, column = i_op).value = sheet_grafik1.cell(row = k, column = 9).value + sheet_grafik1.cell(row = k+1, column = 9).value
                #elif sheet_grafik1.cell(row = k, column = 9).value is not None and sheet_grafik1.cell(row = k+1, column = 9).value is None:
                 #   sheet_rez.cell(row = 6, column = i_op).value = sheet_grafik1.cell(row = k, column = 9).value
                #else:
                 #   sheet_rez.cell(row = 6, column = i_op).value = sheet_grafik1.cell(row = k+1, column = 9).value
                
                i_op += 1
    if i_2_2 > 0:           
        for k in range(i_2_2+1,  i_1_2-3): 
            if sheet_grafik1.cell(row = k, column = 4).value is not None and sheet_grafik1.cell(row = k, column = 6).value is not None and (type(sheet_grafik1.cell(row = k, column = 9).value) == int or type(sheet_grafik1.cell(row = k, column = 9).value) == float  or type(sheet_grafik1.cell(row = k+1, column = 9).value) == int or type(sheet_grafik1.cell(row = k+1, column = 9).value) == float):
                oper = fio_full(sheet_grafik1.cell(row = k, column = 6).value)
                print(oper)
                sheet_rez.cell(row = 1, column = i_op).value = CC_name
                sheet_rez.cell(row = 2, column = i_op).value = fio(oper)
                sheet_rez.cell(row = 3, column = i_op).value = "2/2"
                sheet_rez.cell(row = 4, column = i_op).value = shift(sheet_grafik1.cell(row = k, column = 5).value)[0]
                sheet_rez.cell(row = 5, column = i_op).value = shift(sheet_grafik1.cell(row = k, column = 5).value)[1]
                sheet_rez.cell(row = 6, column = i_op).value = find_cell(k, sheet_grafik1, CC_name, "2/2")[3]
                sheet_rez.cell(row = 7, column = i_op).value = find_cell(k, sheet_grafik1, CC_name, "2/2")[4]
                sheet_rez.cell(row = 8, column = i_op).value = find_cell(k, sheet_grafik1, CC_name, "2/2")[5]
                if find_cell(k, sheet_grafik1, CC_name, "5/2")[7] == 1:
                    sheet_rez.cell(row = 8, column = i_op).fill = redFill
                sheet_rez.cell(row = 9, column = i_op).value = f"=SUM({get_column_letter(i_op)}10:{get_column_letter(i_op)}273)"
                i_op += 1
    if i_1_2 > 0:
        for k in range(i_1_2+1,  sheet_grafik1.max_row + 1): 
            if sheet_grafik1.cell(row = k, column = 4).value is not None and sheet_grafik1.cell(row = k, column = 6).value is not None and (type(sheet_grafik1.cell(row = k, column = 9).value) == int or type(sheet_grafik1.cell(row = k, column = 9).value) == float  or type(sheet_grafik1.cell(row = k+1, column = 9).value) == int or type(sheet_grafik1.cell(row = k+1, column = 9).value) == float):
                oper = fio_full(sheet_grafik1.cell(row = k, column = 6).value)
                print(oper)
                sheet_rez.cell(row = 1, column = i_op).value = CC_name
                sheet_rez.cell(row = 2, column = i_op).value = fio(oper)
                sheet_rez.cell(row = 3, column = i_op).value = "1/2"
                sheet_rez.cell(row = 4, column = i_op).value = shift(sheet_grafik1.cell(row = k, column = 5).value)[0]
                sheet_rez.cell(row = 5, column = i_op).value = shift(sheet_grafik1.cell(row = k, column = 5).value)[1]
                sheet_rez.cell(row = 6, column = i_op).value = find_cell(k, sheet_grafik1, CC_name, "1/2")[3]
                sheet_rez.cell(row = 7, column = i_op).value = find_cell(k, sheet_grafik1, CC_name, "1/2")[4]
                sheet_rez.cell(row = 8, column = i_op).value = find_cell(k, sheet_grafik1, CC_name, "1/2")[5]
                if find_cell(k, sheet_grafik1, CC_name, "5/2")[7] == 1:
                    sheet_rez.cell(row = 8, column = i_op).fill = redFill
                sheet_rez.cell(row = 9, column = i_op).value = f"=SUM({get_column_letter(i_op)}10:{get_column_letter(i_op)}273)"
                i_op += 1
    return(sheet_rez, i_op) 