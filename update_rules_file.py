import xlrd, xlwt


import create_name_rules

create_name_rules.ws.write(0,len(create_name_rules.ans_person) + 1, "Подразделение владелец")
create_name_rules.ws.write(0,len(create_name_rules.ans_person) + 2, "Информационная система")
create_name_rules.ws.write(0,len(create_name_rules.ans_person) + 3, "Права доступа")
create_name_rules.ws.write(0,len(create_name_rules.ans_person) + 4, "Имя сервера")


temp = ""

dict_CN = {}
dict_cn = {}
dict_cn_server = {}

list_CN = []
list_cn = []

list_admin_CN = []
list_rdp_CN = []

########### работа с CN
for number in range(len(create_name_rules.ans_rule)):
    temp = create_name_rules.ans_rule[number]
    if (temp.startswith("CN=MSK-LAG-")):### случай с CN MSK-LAG
        temp = temp.split("CN=MSK-LAG-")[1]
        if ("_RDP" in temp):
            list_rdp_CN.append(create_name_rules.ans_rule[number])
        if ("_Local_Administrator" in temp):
            list_admin_CN.append(create_name_rules.ans_rule[number])
        if ("_" in temp):
            temp = temp.split("_")[0]
        temp = temp.split(",")[0]

        if (create_name_rules.ans_rule[number] not in dict_CN):
            dict_CN[create_name_rules.ans_rule[number]] = []
            dict_CN[create_name_rules.ans_rule[number]].append(temp)
            list_CN.append(create_name_rules.ans_rule[number])
            continue
        else:
            dict_CN[create_name_rules.ans_rule[number]].append(temp)
            list_CN.append(create_name_rules.ans_rule[number])
            continue


    elif (temp.startswith("CN=")):## случай с CN
        temp = temp.split("CN=")[1]

        if ("_RDP" in temp):
            list_rdp_CN.append(create_name_rules.ans_rule[number])
        if ("_Local_Administrator" in temp):
            list_admin_CN.append(create_name_rules.ans_rule[number])

        if ("_" in temp):
            temp = temp.split("_")[0]
        temp = temp.split(",")[0]
        if (create_name_rules.ans_rule[number] not in dict_CN):
            dict_CN[create_name_rules.ans_rule[number]] = []
            dict_CN[create_name_rules.ans_rule[number]].append(temp)
            list_CN.append(create_name_rules.ans_rule[number])
            continue
        else:
            dict_CN[create_name_rules.ans_rule[number]].append(temp)
            list_CN.append(create_name_rules.ans_rule[number])
            continue


wb_passport = xlrd.open_workbook('ПаспортаСерверов_21.03.2017.xls')
sheet_passport = wb_passport.sheet_by_index(0)
sheet_passport1 = wb_passport.sheet_by_index(1)

for x in list_CN:
    temp = dict_CN[x][0]
    row = 0
    col = 0
    while ((row < sheet_passport._cell_values.__len__())):
        if (temp.upper() == sheet_passport.row_values(row)[col].upper()):

            create_name_rules.ws.write(create_name_rules.ans_rule.index(x)+1, len(create_name_rules.ans_person) + 1,sheet_passport.row_values(row)[8])
            create_name_rules.ws.write(create_name_rules.ans_rule.index(x)+1, len(create_name_rules.ans_person) + 2,sheet_passport.row_values(row)[3])
            create_name_rules.ws.write(create_name_rules.ans_rule.index(x)+1, len(create_name_rules.ans_person) + 4,sheet_passport.row_values(row)[0])

        row += 1
    row = 0
    while ((row < sheet_passport1._cell_values.__len__())):
        if (temp.upper() == sheet_passport1.row_values(row)[col].upper()):
            create_name_rules.ws.write(create_name_rules.ans_rule.index(x) + 1, len(create_name_rules.ans_person) + 1,sheet_passport.row_values(row)[8])
        row += 1







#print(sheet_passport._cell_values.__len__())#количество строк


########### работа с cn




create_name_rules.wb.save("answer.xls")
print("asd")