import xlrd, xlwt
import copy

list_person = []
list_rule = []
dict_person = {}
ans_rule = []
ans_person = []
### готов список ролей

wb = xlrd.open_workbook('ipa_kcod_y6.xls')
sheet = wb.sheet_by_index(0)
x = 2
while (1):
    try:
        list_rule.append(sheet.row_values(x)[1])
        if (sheet.row_values(x)[1] not in ans_rule):
            ans_rule.append(sheet.row_values(x)[1])
        x += 1
    except IndexError:
        break
### готов список ролей

x = 2
while (1):
    try:
        list_person.append(sheet.row_values(1)[x])
        if (sheet.row_values(1)[x] not in ans_person):
            ans_person.append(sheet.row_values(1)[x])
        x += 1
    except IndexError:
        break

for x in list_person:
    dict_person[x] = []

for x in range(len(list_person)):
    for y in range(len(list_rule)):
        if (sheet.row_values(y + 2)[x + 2] != ""):
            dict_person[list_person[x]].append(sheet.row_values(y + 2)[1])

# в виде list_person[0][0] - имя, [0][1] и далее - роли
# список людей
#print(sheet.row_values(1)[3]) строка/столбец

wb = xlrd.open_workbook('ipa_opkc_y6.xls')
sheet = wb.sheet_by_index(0)

list_person = []
list_rule = []

x = 2
while (1):
    try:
        list_rule.append(sheet.row_values(x)[1])
        if (sheet.row_values(x)[1] not in ans_rule):
            ans_rule.append(sheet.row_values(x)[1])
        x += 1
    except IndexError:
        break
### готов список ролей

x = 2
while (1):
    try:
        list_person.append(sheet.row_values(1)[x])
        if (sheet.row_values(1)[x] not in ans_person):
            ans_person.append(sheet.row_values(1)[x])
        x += 1
    except IndexError:
        break


for x in range(len(list_person)):
    for y in range(len(list_rule)):
        if (sheet.row_values(y + 2)[x + 2] != ""):
            if list_person[x] in dict_person:
                dict_person[list_person[x]].append(sheet.row_values(y + 2)[1])
            else:
                dict_person[list_person[x]] = []
                dict_person[list_person[x]].append(sheet.row_values(y + 2)[1])



wb = xlrd.open_workbook('КЦОД_AD_УБ.xls')
sheet = wb.sheet_by_index(0)

list_person = []
list_rule = []

x = 2
while (1):
    try:
        list_rule.append(sheet.row_values(x)[1])
        if (sheet.row_values(x)[1] not in ans_rule):
            ans_rule.append(sheet.row_values(x)[1])
        x += 1
    except IndexError:
        break
### готов список ролей

x = 2
while (1):
    try:
        list_person.append(sheet.row_values(1)[x])
        if (sheet.row_values(1)[x] not in ans_person):
            ans_person.append(sheet.row_values(1)[x])
        x += 1
    except IndexError:
        break


for x in range(len(list_person)):
    for y in range(len(list_rule)):
        if (sheet.row_values(y + 2)[x + 2] != ""):
            if list_person[x] in dict_person:
                dict_person[list_person[x]].append(sheet.row_values(y + 2)[1])
            else:
                dict_person[list_person[x]] = []
                dict_person[list_person[x]].append(sheet.row_values(y + 2)[1])



wb = xlrd.open_workbook('ОПКЦ_AD_УБ.xls')
sheet = wb.sheet_by_index(0)

list_person = []
list_rule = []

x = 2
while (1):
    try:
        list_rule.append(sheet.row_values(x)[1])
        if (sheet.row_values(x)[1] not in ans_rule):
            ans_rule.append(sheet.row_values(x)[1])
        x += 1
    except IndexError:
        break
### готов список ролей

x = 2
while (1):
    try:
        list_person.append(sheet.row_values(1)[x])
        if (sheet.row_values(1)[x] not in ans_person):
            ans_person.append(sheet.row_values(1)[x])
        x += 1
    except IndexError:
        break


for x in range(len(list_person)):
    for y in range(len(list_rule)):
        if (sheet.row_values(y + 2)[x + 2] != ""):
            if list_person[x] in dict_person:
                dict_person[list_person[x]].append(sheet.row_values(y + 2)[1])
            else:
                dict_person[list_person[x]] = []
                dict_person[list_person[x]].append(sheet.row_values(y + 2)[1])

##########################################################3 готовы данные о персонале и ролях

wb = xlwt.Workbook()
ws = wb.add_sheet('1')
ws.write(0,0, "Группа прав")

for x in range(len(ans_rule)):
    ws.write(x+1,0,ans_rule[x])

for x in range(len(ans_person)):
    ws.write(0, x+1, ans_person[x])
    for y in range(len(ans_rule)):
        if (ans_rule[y] in dict_person[ans_person[x]]):
            ws.write(y+1, x+1, "Y")

# расставили роли и фио


wb.save("answer.xls")


print("aaa")