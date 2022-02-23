import openpyxl
import pereriv_lib as per
import datetime

wb_grafik_per = openpyxl.load_workbook(f"перерывы_сборка.xlsx")
for s in range(len(wb_grafik_per.sheetnames)):
    if wb_grafik_per.sheetnames[s] == 'перерывы':
        wb_grafik_per.active = s
        break

wb_oper = openpyxl.load_workbook(f"операторы.xlsx")
for s in range(len(wb_oper.sheetnames)):
    if wb_oper.sheetnames[s] == 'операторы':
        wb_oper.active = s
        break


print("Готовим таблицы для рассылки...")
sheet_per = wb_grafik_per.active
sheet_oper = wb_oper.active

sheet_sm = per.new_sheet_oktel(wb_grafik_per, "Лист1")

# формируем рассылку для Октел
per.get_grafik_oktel(sheet_per, sheet_sm)

today = datetime.date.today()
tomorrow = today + datetime.timedelta(days=1)
tomorrow_str = tomorrow.strftime('%Y%m%d')

i_range = []
for i in range(2, sheet_sm.max_row + 1):
    #print(sheet_sm.cell(row=i, column=1).value)
    temp_fio = sheet_sm.cell(row=i, column=1).value
    temp_fio = temp_fio.replace("ё","е")

    for j in range(2, sheet_oper.max_row + 1):
        #print(per.fio(sheet_oper.cell(row=j, column=1).value))

        temp_fio2 = per.fio(sheet_oper.cell(row=j, column=1).value)
        temp_fio2 = temp_fio2.replace("ё","е")
        #print(temp_fio2)
        if temp_fio == temp_fio2 or temp_fio == temp_fio2[:-3] or temp_fio[:-3] == temp_fio2:
            sheet_sm.cell(row=i, column=1).value = sheet_oper.cell(row=j, column=1).value
            #print(per.fio_full(sheet_oper.cell(row=j, column=1).value))
            break
    if temp_fio is None:
        next
    elif temp_fio == sheet_sm.cell(row=i, column=1).value:
        print(f"Ошибка: {temp_fio}")
        i_range.append(i)

for st in reversed(i_range):
    sheet_sm.delete_rows(st)


try:
    wb_grafik_per.save(f"user{tomorrow_str}.xlsx")
    input(f"Таблицы для рассылки подготовлены. Открываем файл user{tomorrow_str}.xlsx и "
          "проверяем. Нажмите ENTER для продолжения...")
except OSError:
    input(f"Невозможно сохранить данные. Закройте файл user{tomorrow_str}.xlsx и "
          "запустите программу заново. Нажмите ENTER для продолжения...")