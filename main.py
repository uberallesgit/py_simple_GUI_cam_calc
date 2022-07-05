import PySimpleGUI as sg
from PIL import Image, ImageDraw, ImageTk, ImageFont
from datetime import datetime as dt
from openpyexcel import load_workbook
import csv
import math
from time import sleep



ahd_2mp =  [[sg.Text(), sg.Column([[sg.Text('Клиент:'),sg.Input(key='-CLIENT-IN-', size=(30, 1),justification="l"),],

                                                      [sg.Text('Примерное количество кабеля,м      '),sg.Input(key='-CABLE-IN-', size=(5, 1),justification="r")],

                                                      [sg.Text('Количество камер:                           '),sg.Input(key='-CAM-IN-', size=(5, 1),justification="r")],
                                                      [sg.Text('Количество портов регистратора:')],
                                                      [sg.Radio('4', 'radio1', default=True, key='-4PORTS-', size=(2, 1)),
                                                      sg.Radio('8', 'radio1', key='-8PORTS-', size=(2, 1)),
                                                      sg.Radio('16', 'radio1', key='-16PORTS-', size=(2, 1)),
                                                      sg.Radio('32', 'radio1', key='-32PORTS-', size=(2, 1))],
        [sg.Text('                             '),sg.Button('OK', key='-CONFIRM_AHD_2MP-',size=(10,1))],
                                [sg.Text("CSV-Файл со сметой  сформирован! ", key="--FINAL_MESSAGE--",visible=False)],
                                   [sg.Image(filename="ahd.png")]
                                                      ], size=(350, 350), pad=(0, 0))]]

ahd_5mp = [[sg.Text(), sg.Column([[sg.Text('Клиент:'),sg.Input(key='-CLIENT-IN-', size=(30, 1),justification="l"),],

                                                      [sg.Text('Примерное количество кабеля,м      '),sg.Input(key='-CABLE-IN-', size=(5, 1),justification="r")],

                                                      [sg.Text('Количество камер:                           '),sg.Input(key='-CAM-IN-', size=(5, 1),justification="r")],
                                                      [sg.Text('Количество портов регистратора:')],
                                                      [sg.Radio('4', 'radio2', default=True, key='-4PORTS-', size=(2, 1)),
                                                      sg.Radio('8', 'radio2', key='-8PORTS-', size=(2, 1)),
                                                      sg.Radio('16', 'radio2', key='-16PORTS-', size=(2, 1)),
                                                      sg.Radio('32', 'radio2', key='-32PORTS-', size=(2, 1))],
        [sg.Text('                             '),sg.Button('OK', key='-CONFIRM_AHD_5MP-',size=(10,1))],
                                [sg.Text("CSV-Файл со сметой  сформирован! ", key="--FINAL_MESSAGE--",visible=False)],
                                  [sg.Image(filename="ahd.png")]
                                                      ], size=(350, 350), pad=(0, 0))]]
ip_2mp = [[sg.Text(), sg.Column([[sg.Text('Клиент:'),sg.Input(key='-CLIENT-IN-', size=(30, 1),justification="l"),],

                                                      [sg.Text('Примерное количество кабеля,м      '),sg.Input(key='-CABLE-IN-', size=(5, 1),justification="r")],

                                                      [sg.Text('Количество камер:                           '),sg.Input(key='-CAM-IN-', size=(5, 1),justification="r")],
                                                      [sg.Text('Количество портов регистратора:')],
                                                      [sg.Radio('4', 'radio3', default=True, key='-4PORTS-', size=(2, 1)),
                                                      sg.Radio('8', 'radio3', key='-8PORTS-', size=(2, 1)),
                                                      sg.Radio('16', 'radio3', key='-16PORTS-', size=(2, 1)),
                                                      sg.Radio('32', 'radio3', key='-32PORTS-', size=(2, 1))],
        [sg.Text('                             '),sg.Button('OK', key='-CONFIRM_IP_2MP-',size=(10,1))],
                                [sg.Text("CSV-Файл со сметой  сформирован! ", key="--FINAL_MESSAGE--",visible=False)],
                                 [sg.Image(filename="IP.png")]
                                                      ], size=(350, 350), pad=(0, 0))]]

ip_5mp = [[sg.Text(), sg.Column([[sg.Text('Клиент:'),sg.Input(key='-CLIENT-IN-', size=(30, 1),justification="l"),],

                                                      [sg.Text('Примерное количество кабеля,м      '),sg.Input(key='-CABLE-IN-', size=(5, 1),justification="r")],

                                                      [sg.Text('Количество камер:                           '),sg.Input(key='-CAM-IN-', size=(5, 1),justification="r")],
                                                      [sg.Text('Количество портов регистратора:')],
                                                      [sg.Radio('4', 'radio4', default=True, key='-4PORTS-', size=(2, 1)),
                                                      sg.Radio('8', 'radio4', key='-8PORTS-', size=(2, 1)),
                                                      sg.Radio('16', 'radio4', key='-16PORTS-', size=(2, 1)),
                                                      sg.Radio('32', 'radio4', key='-32PORTS-', size=(2, 1))],
        [sg.Text('                             '),sg.Button('OK', key='-CONFIRM_IP_5MP-',size=(10,1))],
                                [sg.Text("CSV-Файл со сметой  сформирован! ", key="--FINAL_MESSAGE--",visible=False)],
                                 [sg.Image(filename="IP.png")]
                                                      ], size=(350, 350), pad=(0, 0))]]

compact_ip =  [[sg.Text(), sg.Column([[sg.Text('Клиент:'),sg.Input(key='-CLIENT-IN-', size=(30, 1),justification="l"),],

                                                      [sg.Text('Примерное количество кабеля,м      '),sg.Input(key='-CABLE-IN-', size=(5, 1),justification="r")],

                                                      [sg.Text('Количество камер:                           '),sg.Input(key='-CAM-IN-', size=(5, 1),justification="r")],

        [sg.Text('                             '),sg.Button('OK', key='-CONFIRM_COMPAC-',size=(10,1))],
                                [sg.Text("CSV-Файл со сметой  сформирован! ", key="--FINAL_MESSAGE--",visible=False)],
                                 [sg.Image(filename="compac.png")]

                                                      ], size=(350, 350), pad=(0, 0))]]

tab_group = [[sg.TabGroup(
                  [[
                    sg.Tab("2MP",ahd_2mp,key='-2MP-',expand_x=True,background_color="Brown",p=(6,6)),
                    sg.Tab("5MP",ahd_5mp,key='-5MP-',expand_x=True,background_color="Blue"),
                    sg.Tab("IP-2MP",ip_2mp,key='-IP2MP-',expand_x=True,background_color="Orange"),
                    sg.Tab("IP-5MP",ip_5mp,key='-IP5MP-',expand_x=True,background_color="Purple"),
                    sg.Tab("Compact-IP",compact_ip, key='-COMPAC-',expand_x=True,background_color="Magenta"),
                  ]],pad=(10,10))]]

window = sg.Window('Cameras calculator', tab_group)

def today_is():
    return dt.now().strftime("%d.%m.%Y")


reg_channels = 0
counter = 1
quantity = 0
client = ""
reg_count = 0
hdd_count = 0



##################### Объявляем эксэль книгу  и лист ######################################################

book = load_workbook(filename="Equipment.xlsx")
sheet = book["Equip"]
############################################################################################################

def cam_calc_1(counter,quantity):
    print("Вариант 2MP...")
    list = [2,15,16,18]
    for i in list:
         with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",newline="") as file:
             writer = csv.writer(file)
             writer.writerow((counter,sheet[f"a{i}"].value,sheet[f"B{i}"].value,quantity,int(sheet[f"B{i}"].value)*int(quantity)))
             counter+=1

    #Единичные товары(жесткий, БП)
    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",newline="") as file:
        writer = csv.writer(file)
        #Считаем кабель
        writer.writerow((counter, sheet[f"a{17}"].value, sheet[f"B{17}"].value, cable, int(sheet[f"B{17}"].value) * int(cable)))
        counter+=1

        #Жесткий диск
        writer.writerow((counter, sheet[f"a{11}"].value, sheet[f"B{11}"].value, 1, int(sheet[f"B{11}"].value) * 1))
        counter+=1

    # Условие  выбора регистратора:  ############################################################################
    if values["-4PORTS-"]:
        reg_count = 4
        ports = 4
        if int(quantity) > ports:
            sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")
    if values["-8PORTS-"]:
        reg_count = 5
        ports = 8
        if int(quantity) > ports:
            sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")
    if values["-16PORTS-"]:
        reg_count = 6
        ports = 16
        if int(quantity) > ports:
            sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")
    if values["-32PORTS-"]:
        reg_count = 22
        ports = 32
        if int(quantity) > ports:
            sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")

    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)
        writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                         int(sheet[f"B{reg_count}"].value) * 1))
        counter += 1

        # подсчет суммы :  ####################################################################################

    list2 = [reg_count, 11]
    sum1 = 0
    sum2 = 0

    for i in list:
        sum1 = int(sheet[f"B{i}"].value)*int(quantity)+sum1

    for i in list2:
        sum2 = int(sheet[f"B{i}"].value) + sum2

    # расчет мощности блоков питания(2А или 3А)
    power_supply = 0

    if int(quantity) < 2:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{14}"].value, sheet[f"B{14}"].value, 1,
                             int(sheet[f"B{14}"].value) * 1))
            power_supply = int(sheet[f"B{14}"].value)
    elif 2 <= int(quantity) < 4:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, 1,
                             int(sheet[f"B{13}"].value) * 1))
        power_supply = int(sheet[f"B{13}"].value)
    else:
        #расчет количества блоков питания (power suply quantity(PSQ))    ########################################

        psq = math.ceil(int(quantity)/7)
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, psq,
                             int(sheet[f"B{13}"].value) * psq))
        power_supply = int(sheet[f"B{13}"].value) * psq

    #Всего кабеля

    sum_cab = int(cable)*int(sheet[f"B{17}"].value)

    #Итого
    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(("****","**************Итого****************","*******","********", sum1+sum2+sum_cab+power_supply))
########################################################################################################################

def cam_calc_2(counter,quantity):
    print("Вариант 5MP...")


    list = [3, 15, 16, 18]

    for i in list:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{i}"].value, sheet[f"B{i}"].value, quantity,
                             int(sheet[f"B{i}"].value) * int(quantity)))
            counter += 1

    # Единичные товары(жесткий, БП)
    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)
        # Считаем кабель
        writer.writerow(
            (counter, sheet[f"a{17}"].value, sheet[f"B{17}"].value, cable, int(sheet[f"B{17}"].value) * int(cable)))
        counter += 1

        # Жесткий диск
        writer.writerow((counter, sheet[f"a{11}"].value, sheet[f"B{11}"].value, 1, int(sheet[f"B{11}"].value) * 1))
        counter += 1

        # Условие  выбора регистратора:  ############################################################################
        if values["-4PORTS-"]:
            reg_count = 4
            ports = 4
            if int(quantity) > ports:
                sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")
        if values["-8PORTS-"]:
            reg_count = 5
            ports = 8
            if int(quantity) > ports:
                sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")
        if values["-16PORTS-"]:
            reg_count = 6
            ports = 16
            if int(quantity) > ports:
                sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")
        if values["-32PORTS-"]:
            reg_count = 22
            ports = 32
            if int(quantity) > ports:
                sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")

    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)
        writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                         int(sheet[f"B{reg_count}"].value) * 1))
        counter += 1

        # подсчет суммы :  ####################################################################################

    list2 = [reg_count, 11]
    sum1 = 0
    sum2 = 0

    for i in list:
        sum1 = int(sheet[f"B{i}"].value) * int(quantity) + sum1

    for i in list2:
        sum2 = int(sheet[f"B{i}"].value) + sum2

    # расчет мощности блоков питания(2А или 3А)
    power_supply = 0

    if int(quantity) < 2:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{14}"].value, sheet[f"B{14}"].value, 1,
                             int(sheet[f"B{14}"].value) * 1))
            power_supply = int(sheet[f"B{14}"].value)
    elif 2 <= int(quantity) < 4:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, 1,
                             int(sheet[f"B{13}"].value) * 1))
        power_supply = int(sheet[f"B{13}"].value)
    else:
        # расчет количества блоков питания (power suply quantity(PSQ))    ########################################

        psq = math.ceil(int(quantity) / 7)
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, psq,
                             int(sheet[f"B{13}"].value) * psq))
        power_supply = int(sheet[f"B{13}"].value) * psq

    # Всего кабеля

    sum_cab = int(cable) * int(sheet[f"B{17}"].value)

    # Итого
    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(("", "Итого", "", "", sum1 + sum2 + sum_cab + power_supply))
########################################################################################################################

def ip_cam_calc_3(counter,quantity):
    print("Вариант P 2MP...")
    counter = 1
    list = [9, 21, 16, 18]
    for i in list:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{i}"].value, sheet[f"B{i}"].value, quantity,
                             int(sheet[f"B{i}"].value) * int(quantity)))
            counter += 1
    # Единичные товары(жесткий, БП)
    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)

        # Считаем кабель
        writer.writerow(
            (counter, sheet[f"a{24}"].value, sheet[f"B{24}"].value, cable, int(sheet[f"B{24}"].value) * int(cable)))
        counter += 1
        # Жесткий диск
        writer.writerow((counter, sheet[f"a{11}"].value, sheet[f"B{11}"].value, 1, int(sheet[f"B{11}"].value) * 1))
        counter += 1
        # Условие  выбора регистратора:  ############################################################################
        if values["-4PORTS-"]:
            reg_count = 4
            ports = 4
            if int(quantity) > ports:
                sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")
        if values["-8PORTS-"]:
            reg_count = 5
            ports = 8
            if int(quantity) > ports:
                sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")
        if values["-16PORTS-"]:
            reg_count = 6
            ports = 16
            if int(quantity) > ports:
                sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")
        if values["-32PORTS-"]:
            reg_count = 22
            ports = 32
            if int(quantity) > ports:
                sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")

    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)
        writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                         int(sheet[f"B{reg_count}"].value) * 1))
        counter += 1

        ############## подсчет суммы

    list2 = [reg_count, 11]



    sum1 = 0
    sum2 = 0

    for i in list:
        sum1 = int(sheet[f"B{i}"].value) * int(quantity) + sum1

    for i in list2:
        sum2 = int(sheet[f"B{i}"].value) + sum2

    # Блок питания :условие для подсчета количества БП
    power_supply = 0

    if int(quantity) < 2:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{14}"].value, sheet[f"B{14}"].value, 1,
                             int(sheet[f"B{14}"].value) * 1))
            power_supply = int(sheet[f"B{14}"].value)
    elif int(quantity) >= 2 and int(quantity) < 4:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, 1,
                             int(sheet[f"B{13}"].value) * 1))
            power_supply = int(sheet[f"B{13}"].value)
    else:
        # расчет количества блоков питания (power suply quantity), PSQ
        psq = math.ceil(int(quantity) / 7)
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, psq,
                             int(sheet[f"B{13}"].value) * psq))
            power_supply = int(sheet[f"B{13}"].value) * psq

    # Всего кабеля
    sum_cab = int(cable) * int(sheet[f"B{24}"].value)

    # Итого
    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(("", "Итого", "", "", sum1 + sum2 + sum_cab + power_supply))

########################################################################################################################

def ip_cam_calc_4(counter,quantity):
    print("Вариант IP 5MP...")

    counter = 1
    list = [10, 21, 16, 18]
    for i in list:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{i}"].value, sheet[f"B{i}"].value, quantity,
                             int(sheet[f"B{i}"].value) * int(quantity)))
            counter += 1

    # Единичные товары(жесткий, БП)
    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)

        # Считаем кабель
        writer.writerow(
            (counter, sheet[f"a{24}"].value, sheet[f"B{24}"].value, cable, int(sheet[f"B{24}"].value) * int(cable)))
        counter += 1
        # Жесткий диск
        writer.writerow((counter, sheet[f"a{11}"].value, sheet[f"B{11}"].value, 1, int(sheet[f"B{11}"].value) * 1))
        counter += 1

        # Условие  выбора регистратора:  ############################################################################
        if values["-4PORTS-"]:
            reg_count = 4
            ports = 4
            if int(quantity) > ports:
                sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")
        if values["-8PORTS-"]:
            reg_count = 5
            ports = 8
            if int(quantity) > ports:
                sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")
        if values["-16PORTS-"]:
            reg_count = 6
            ports = 16
            if int(quantity) > ports:
                sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")
        if values["-32PORTS-"]:
            reg_count = 22
            ports = 32
            if int(quantity) > ports:
                sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")

        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1
    list2 = [reg_count, 11]

    # подсчет суммы

    sum1 = 0
    sum2 = 0

    for i in list:
        sum1 = int(sheet[f"B{i}"].value) * int(quantity) + sum1

    for i in list2:
        sum2 = int(sheet[f"B{i}"].value) + sum2

    # Блок питания :условие для подсчета количества БП
    power_supply = 0

    if int(quantity) < 2:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{14}"].value, sheet[f"B{14}"].value, 1,
                             int(sheet[f"B{14}"].value) * 1))
            power_supply = int(sheet[f"B{14}"].value)
    elif int(quantity) >= 2 and int(quantity) < 4:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, 1,
                             int(sheet[f"B{13}"].value) * 1))
            power_supply = int(sheet[f"B{13}"].value)
    else:
        # расчет количества блоков питания (power suply quantity), PSQ
        psq = math.ceil(int(quantity) / 7)
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, psq,
                             int(sheet[f"B{13}"].value) * psq))
            power_supply = int(sheet[f"B{13}"].value) * psq

    # Всего кабеля
    sum_cab = int(cable) * int(sheet[f"B{24}"].value)

    # Итого
    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(("", "Итого", "", "", sum1 + sum2 + sum_cab + power_supply))
        print("sum1", sum1)
        print("sum2", sum2)
        print("sum_cab", sum_cab)
        print("power_supply", power_supply)
########################################################################################################################

def ip_sd_calc_5(counter,quantity):
    print("Вариант IP-Compac ...")
    counter = 1
    list = [7,12,16]
    cam_sum = 0
    for i in list:
        with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{i}"].value, sheet[f"B{i}"].value, quantity,
                             int(sheet[f"B{i}"].value) * int(quantity)))
            cam_sum = (int(sheet[f"B{i}"].value) * int(quantity)) + cam_sum
            counter += 1

    with open(f"Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(("", "Итого", "", "", cam_sum))
########################################################################################################################

#MAIN LOOP

while True:
    event, values = window.read()
    print(event, values)
    if event == sg.WIN_CLOSED:
        break


    if event == '-CONFIRM_5MP-' or "-CONFIRM_AHD_2MP-" or "-CONFIRM_IP_2MP-" or "-CONFIRM_IP_5MP-" or "-COMPAC-":
        client = values["-CLIENT-IN-"]
        if client == "":
            client = "somebody_someone"
        print(client)
        cable = values["-CABLE-IN-"]
        if not cable.isnumeric():
            cable = "0"
        else:
            print("cable = ", cable)
        quantity = values["-CAM-IN-"]

        if quantity == "":
            quantity = 1
            sg.popup("Количество камер не указано!\nБудет сформирована смета для ОДНОЙ камеры!")
        elif int(quantity) <= 0:
            sg.popup("Пардон, сударь, но, похоже, вы гоните ! Количество камер меньше 1")


    print(f"РАсчеты")
    with open(f"Видеонаблюдение для {client.title()} на {today_is()}.csv", "w", encoding="utf-8", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(("№", "Наименование", "Цена", "количество", "Сумма!"))




    if event == "-CONFIRM_AHD_2MP-":
        cam_calc_1(counter,quantity)

    elif event == "-CONFIRM_AHD_5MP-":
        cam_calc_2(counter,quantity)

    elif event == "-CONFIRM_IP_2MP-":
        ip_cam_calc_3(counter,quantity)

    elif event == "-CONFIRM_IP_5MP-":
        ip_cam_calc_4(counter,quantity)

    elif event == '-CONFIRM_COMPAC-':
        ip_sd_calc_5(counter,quantity)

    window["--FINAL_MESSAGE--"].update(visible=True)



window.close()

