import PySimpleGUI as sg
from PIL import Image, ImageDraw, ImageTk, ImageFont
from datetime import datetime as dt
from openpyexcel import load_workbook
import csv
import math
import os

#Переменные


reg_channels = 0
counter = 1
quantity = 0
client = ""
reg_count = 0
hdd_count = 0
hdd_size = None
cable = 0
power_backup = 0
power_backup_index = 32


#Интерфейс


ahd_2mp =  [[sg.Text(), sg.Column([[sg.Text('Клиент:'),sg.Input(key='-CLIENT-IN-2MP-', size=(30, 1),justification="l"),],

                                                      [sg.Text('Примерное количество кабеля,м      '),sg.Input(key='-CABLE-IN-2MP-', size=(5, 1),justification="r")],

                                                      [sg.Text('Количество камер:                           '),sg.Input(key='-CAM-IN-2MP-', size=(5, 1),justification="r")],
[sg.Text('_'*45,text_color="Darkgray")],
                                                      [sg.Text('Количество портов регистратора:')],
                                                      [sg.Radio('4', 'radio1', default=True, key='-4PORTS1-', size=(2, 1)),
                                                      sg.Radio('8', 'radio1', key='-8PORTS1-', size=(2, 1)),
                                                      sg.Radio('16', 'radio1', key='-16PORTS1-', size=(2, 1)),
                                                      sg.Radio('32', 'radio1', key='-32PORTS1-', size=(2, 1))],
[sg.Text('_'*45,text_color="Darkgray")],


[sg.Text('Размер жесткого диска:')],             [sg.Radio('1 ТБ', 'radio_HDD_SIZE_2MP', default=True, key='-1TB-HDD-2MP-', size=(5, 1)),
                                                  sg.Radio('2 ТБ', 'radio_HDD_SIZE_2MP', key='-2TB-HDD-2MP-', size=(5, 1)),
                                                  sg.Radio('3 ТБ ', 'radio_HDD_SIZE_2MP', key='-3TB-HDD-2MP-', size=(5, 1)),
                                                  sg.Radio('4 ТБ ', 'radio_HDD_SIZE_2MP', key='-4TB-HDD-2MP-', size=(5, 1))],
                                                  [sg.Radio('5 ТБ ', 'radio_HDD_SIZE_2MP', key='-5TB-HDD-2MP-', size=(5, 1)),
                                                  sg.Radio('6 ТБ ', 'radio_HDD_SIZE_2MP', key='-6TB-HDD-2MP-', size=(5, 1)),
                                                  sg.Radio('7 ТБ ', 'radio_HDD_SIZE_2MP', key='-7TB-HDD-2MP-', size=(5, 1)),
                                                  sg.Radio('8 ТБ', 'radio_HDD_SIZE_2MP', key='-8TB-HDD-2MP-', size=(5, 1))],
[sg.Text('_'*45,text_color="Darkgray")],

[sg.Text('Количество жестких дисков:')],              [sg.Radio('1 HDD', 'radio_HDD_2MP', default=True, key='-1ITEMS-HDD-2MP-', size=(5, 1)),
                                                      sg.Radio('2 HDD', 'radio_HDD_2MP', key='-2ITEMS-HDD-2MP-', size=(5, 1))],
[sg.Text('_'*45,text_color="Darkgray")],

[sg.Text('Блок бесперебойного питания:')],              [sg.Radio('Нет', 'power_backup', default=True, key='-PB-NO-2MP-', size=(5, 1)),
                                                      sg.Radio('Да', 'power_backup', key='-PB-YES-2MP-', size=(5, 1))],
[sg.Text('_'*45,text_color="Darkgray")],


        [sg.Text('                             '),sg.Button('OK', key='-CONFIRM_AHD_2MP-',size=(10,1))],
                                # [sg.Text("CSV-Файл со сметой  сформирован! ", key="--FINAL_MESSAGE--",visible=False)],
                                   [sg.Text('      '), sg.Image(filename="img/ahd.png")]
                                                      ],  pad=(0, 0))]]

ahd_5mp = [[sg.Text(), sg.Column([[sg.Text('Клиент:'),sg.Input(key='-CLIENT-IN-5MP-', size=(30, 1),justification="l"),],

                                                      [sg.Text('Примерное количество кабеля,м      '),sg.Input(key='-CABLE-IN-5MP-', size=(5, 1),justification="r")],

                                                      [sg.Text('Количество камер:                           '),sg.Input(key='-CAM-IN-5MP-', size=(5, 1),justification="r")],
[sg.Text('_'*45,text_color="Darkgray")],
                                                      [sg.Text('Количество портов регистратора:')],
                                                      [sg.Radio('4', 'radio2', default=True, key='-4PORTS2-', size=(2, 1)),
                                                      sg.Radio('8', 'radio2', key='-8PORTS2-', size=(2, 1)),
                                                      sg.Radio('16', 'radio2', key='-16PORTS2-', size=(2, 1)),
                                                      sg.Radio('32', 'radio2', key='-32PORTS2-', size=(2, 1))],
[sg.Text('_'*45,text_color="Darkgray")],

[sg.Text('Размер жесткого диска:')],             [sg.Radio('1 ТБ', 'radio_HDD_SIZE_5MP', default=True, key='-1TB-HDD-5MP-', size=(5, 1)),
                                                  sg.Radio('2 ТБ', 'radio_HDD_SIZE_5MP', key='-2TB-HDD-5MP-', size=(5, 1)),
                                                  sg.Radio('3 ТБ ', 'radio_HDD_SIZE_5MP', key='-3TB-HDD-5MP-', size=(5, 1)),
                                                  sg.Radio('4 ТБ ', 'radio_HDD_SIZE_5MP', key='-4TB-HDD-5MP-', size=(5, 1))],
                                                  [sg.Radio('5 ТБ ', 'radio_HDD_SIZE_5MP', key='-5TB-HDD-5MP-', size=(5, 1)),
                                                  sg.Radio('6 ТБ ', 'radio_HDD_SIZE_5MP', key='-6TB-HDD-5MP-', size=(5, 1)),
                                                  sg.Radio('7 ТБ ', 'radio_HDD_SIZE_5MP', key='-7TB-HDD-5MP-', size=(5, 1)),
                                                  sg.Radio('8 ТБ', 'radio_HDD_SIZE_5MP', key='-8TB-HDD-5MP-', size=(5, 1))],
[sg.Text('_'*45,text_color="Darkgray")],

[sg.Text('Количество жестких дисков:')],[sg.Radio('1 HDD', 'radio_HDD_5MP', default=True, key='-1ITEMS-HDD-5MP-', size=(5, 1)),
                                                      sg.Radio('2 HDD', 'radio_HDD_5MP', key='-2ITEMS-HDD-5MP-', size=(5, 1))],
[sg.Text('_'*45,text_color="Darkgray")],

[sg.Text('Блок бесперебойного питания:')],              [sg.Radio('Нет', 'power_backup_ahd_5', default=True, key='-PB-NO-5MP-', size=(5, 1)),
                                                      sg.Radio('Да', 'power_backup_ahd_5', key='-PB-YES-5MP-', size=(5, 1))],
[sg.Text('_'*45,text_color="Darkgray")],

        [sg.Text('                             '),sg.Button('OK', key='-CONFIRM_AHD_5MP-',size=(10,1))],
                                # [sg.Text("CSV-Файл со сметой  сформирован! ", key="--FINAL_MESSAGE--",visible=False)],
                                  [sg.Text('      '), sg.Image(filename="img/ahd.png")]
                                                      ], pad=(0, 0))]]
ip_2mp = [[sg.Text(), sg.Column([[sg.Text('Клиент:'),sg.Input(key='-CLIENT-IN-IP2MP-', size=(30, 1),justification="l"),],

                                                      [sg.Text('Примерное количество кабеля,м      '),sg.Input(key='-CABLE-IN-IP2MP-', size=(5, 1),justification="r")],

                                                      [sg.Text('Количество камер:                           '),sg.Input(key='-CAM-IN-IP2MP-', size=(5, 1),justification="r")],

[sg.Text('_'*45,text_color="Darkgray")],

                                                      [sg.Text('Количество портов регистратора:')],
                                                      [sg.Radio('4', 'radio3', default=True, key='-4PORTS3-', size=(2, 1)),
                                                      sg.Radio('8', 'radio3', key='-8PORTS3-', size=(2, 1)),
                                                      sg.Radio('9', 'radio3', key='-9PORTS3-', size=(2, 1)),
                                                      sg.Radio('16', 'radio3', key='-16PORTS3-', size=(2, 1)),
                                                      sg.Radio('32', 'radio3', key='-32PORTS3-', size=(2, 1))],
[sg.Text('_'*45,text_color="Darkgray")],

[sg.Text('Размер жесткого диска:')],             [sg.Radio('1 ТБ', 'radio_HDD_SIZE_IP2MP', default=True, key='-1TB-HDD-IP2MP-', size=(5, 1)),
                                                  sg.Radio('2 ТБ', 'radio_HDD_SIZE_IP2MP', key='-2TB-HDD-IP2MP-', size=(5, 1)),
                                                  sg.Radio('3 ТБ ', 'radio_HDD_SIZE_IP2MP', key='-3TB-HDD-IP2MP-', size=(5, 1)),
                                                  sg.Radio('4 ТБ ', 'radio_HDD_SIZE_IP2MP', key='-4TB-HDD-IP2MP-', size=(5, 1))],
                                                  [sg.Radio('5 ТБ ', 'radio_HDD_SIZE_IP2MP', key='-5TB-HDD-IP2MP-', size=(5, 1)),
                                                  sg.Radio('6 ТБ ', 'radio_HDD_SIZE_IP2MP', key='-6TB-HDD-IP2MP-', size=(5, 1)),
                                                  sg.Radio('7 ТБ ', 'radio_HDD_SIZE_IP2MP', key='-7TB-HDD-IP2MP-', size=(5, 1)),
                                                  sg.Radio('8 ТБ', 'radio_HDD_SIZE_IP2MP', key='-8TB-HDD-IP2MP-', size=(5, 1))],

[sg.Text('_'*45,text_color="Darkgray")],

[sg.Text('Количество жестких дисков:')],[sg.Radio('1 HDD', 'radio_HDD_ip2mp', default=True, key='-1ITEMS-HDD-IP2MP-', size=(5, 1)),
                                                      sg.Radio('2 HDD', 'radio_HDD_ip2mp', key='-2ITEMS-HDD-IP2MP-', size=(5, 1))],
[sg.Text('_'*45 ,text_color="Darkgray")],

[sg.Text('Блок бесперебойного питания:')],              [sg.Radio('Нет', 'power_backup_ip_2mp', default=True, key='-PB-NO-IP2MP-', size=(5, 1)),
                                                      sg.Radio('Да', 'power_backup_ip_2mp', key='-PB-YES-IP2MP-', size=(5, 1))],
[sg.Text('_'*45,text_color="Darkgray")],

        [sg.Text('                             '),sg.Button('OK', key='-CONFIRM_IP_2MP-',size=(10,1))],
                                 [sg.Text('      '), sg.Image(filename="img/IP.png")]
                                                      ], pad=(0, 0))]]

ip_5mp = [[sg.Text(), sg.Column([[sg.Text('Клиент:'),sg.Input(key='-CLIENT-IN-IP5MP-', size=(30, 1),justification="l"),],

                                                      [sg.Text('Примерное количество кабеля,м      '),sg.Input(key='-CABLE-IN-IP5MP-', size=(5, 1),justification="r")],

                                                      [sg.Text('Количество камер:                           '),sg.Input(key='-CAM-IN-IP5MP-', size=(5, 1),justification="r")],

[sg.Text('_'*45,text_color="Darkgray")],
                                                      [sg.Text('Количество портов регистратора:')],
                                                      [sg.Radio('4', 'radio4', default=True, key='-4PORTS4-', size=(2, 1)),
                                                      sg.Radio('8', 'radio4', key='-8PORTS4-', size=(2, 1)),
                                                      sg.Radio('9', 'radio4', key='-9PORTS4-', size=(2, 1)),
                                                      sg.Radio('16', 'radio4', key='-16PORTS4-', size=(2, 1)),
                                                      sg.Radio('32', 'radio4', key='-32PORTS4-', size=(2, 1))],

[sg.Text('_'*45,text_color="Darkgray")],

[sg.Text('Размер жесткого диска:')],             [sg.Radio('1 ТБ', 'radio_HDD_SIZE_IP5MP', default=True, key='-1TB-HDD-IP5MP-', size=(5, 1)),
                                                  sg.Radio('2 ТБ', 'radio_HDD_SIZE_IP5MP', key='-2TB-HDD-IP5MP-', size=(5, 1)),
                                                  sg.Radio('3 ТБ ', 'radio_HDD_SIZE_IP5MP', key='-3TB-HDD-IP5MP-', size=(5, 1)),
                                                  sg.Radio('4 ТБ ', 'radio_HDD_SIZE_IP5MP', key='-4TB-HDD-IP5MP-', size=(5, 1))],
                                                  [sg.Radio('5 ТБ ', 'radio_HDD_SIZE_IP5MP', key='-5TB-HDD-IP5MP-', size=(5, 1)),
                                                  sg.Radio('6 ТБ ', 'radio_HDD_SIZE_IP5MP', key='-6TB-HDD-IP5MP-', size=(5, 1)),
                                                  sg.Radio('7 ТБ ', 'radio_HDD_SIZE_IP5MP', key='-7TB-HDD-IP5MP-', size=(5, 1)),
                                                  sg.Radio('8 ТБ', 'radio_HDD_SIZE_IP5MP', key='-8TB-HDD-IP5MP-', size=(5, 1))],

[sg.Text('_'*45,text_color="Darkgray")],

[sg.Text('Количество жестких дисков:')],[sg.Radio('1 HDD', 'radio_HDD_ip5mp', default=True, key='-1ITEMS-HDD-IP5MP-', size=(5, 1)),
                                                      sg.Radio('2 HDD', 'radio_HDD_ip5mp', key='-2ITEMS-HDD-IP5MP-', size=(5, 1))],

[sg.Text('_'*45,text_color="Darkgray")],

[sg.Text('Блок бесперебойного питания:')],              [sg.Radio('Нет', 'power_backup_ip_5mp', default=True, key='-PB-NO-IP5MP-', size=(5, 1)),
                                                      sg.Radio('Да', 'power_backup_ip_5mp', key='-PB-YES-IP5MP-', size=(5, 1))],
[sg.Text('_'*45,text_color="Darkgray")],

        [sg.Text('                             '),sg.Button('OK', key='-CONFIRM_IP_5MP-',size=(10,1))],
                                # [sg.Text("CSV-Файл со сметой  сформирован! ", key="--FINAL_MESSAGE--",visible=False)],
                                 [sg.Text('      '), sg.Image(filename="img/IP.png")]
                                                      ], pad=(0, 0))]]

compact_ip =  [[sg.Text(), sg.Column([[sg.Text('Клиент:'),sg.Input(key='-CLIENT-IN-COMP-', size=(30, 1),justification="l"),],

                                                      [sg.Text('Примерное количество кабеля,м      '),sg.Input(key='-CABLE-IN-COMP-', size=(5, 1),justification="r")],

                                                      [sg.Text('Количество камер:                           '),sg.Input(key='-CAM-IN-COMP-', size=(5, 1),justification="r")],

        [sg.Text('                         '),sg.Button('OK', key='-CONFIRM_COMPAC-',size=(10,1))],
                                # [sg.Text("CSV-Файл со сметой  сформирован! ", key="--FINAL_MESSAGE--",visible=False)],
                                 [sg.Text('     '), sg.Image(filename="img/compac.png")]

                                                      ], pad=(0, 0))]]

tab_group = [[sg.TabGroup(
                  [[
                    sg.Tab("2MP",ahd_2mp,key='-2MP-',expand_x=True,background_color="Green"),
                    sg.Tab("5MP",ahd_5mp,key='-5MP-',expand_x=True,background_color="Blue"),
                    sg.Tab("IP-2MP",ip_2mp,key='-IP2MP-',expand_x=True,background_color="Orange"),
                    sg.Tab("IP-5MP",ip_5mp,key='-IP5MP-',expand_x=True,background_color="Purple"),
                    sg.Tab("Compact-IP",compact_ip, key='-COMPAC-',expand_x=True),
                  ]],pad= 0)]]

window = sg.Window('MatrixCam', tab_group)





def today_is():
    return dt.now().strftime("%d.%m.%Y")







##################### Объявляем эксэль книгу  и лист ######################################################

book = load_workbook(filename="Equipment.xlsx")
sheet = book["Equip"]
############################################################################################################

def parse_prices():
    pass








def cam_calc_1(counter,quantity,hdd_count,hdd_size,power_backup):
    print("Вариант 2MP...")
    list = [2,15,16,18]
    for i in list:
         with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",newline="") as file:
             writer = csv.writer(file)
             writer.writerow((counter,sheet[f"a{i}"].value,sheet[f"B{i}"].value,quantity,int(sheet[f"B{i}"].value)*int(quantity)))
             counter+=1

    #Единичные товары(жесткий, БП)
    with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",newline="") as file:
        writer = csv.writer(file)
        #Считаем кабель
        writer.writerow((counter, sheet[f"a{17}"].value, sheet[f"B{17}"].value, cable, int(sheet[f"B{17}"].value) * int(cable)))
        counter+=1

        #Жесткий диск

        #Размер жесткого диска
    if values['-1TB-HDD-2MP-']:
        hdd_size = 11

    elif  values['-2TB-HDD-2MP-']:
        hdd_size = 25

    elif values['-3TB-HDD-2MP-']:
        hdd_size = 26

    elif values['-4TB-HDD-2MP-']:
        hdd_size = 27

    elif values['-5TB-HDD-2MP-']:
        hdd_size = 28

    elif values['-6TB-HDD-2MP-']:
        hdd_size = 29

    elif values['-7TB-HDD-2MP-']:
        hdd_size = 30

    elif values['-8TB-HDD-2MP-']:
        hdd_size = 31


        # Количество жестких дисков

    if values['-1ITEMS-HDD-2MP-']:
        hdd_count = 1
    elif values['-2ITEMS-HDD-2MP-']:
        hdd_count = 2


    with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",newline="") as file:
        writer = csv.writer(file)
        writer.writerow((counter, sheet[f"a{hdd_size}"].value, sheet[f"B{hdd_size}"].value, hdd_count, int(sheet[f"B{hdd_size}"].value) * hdd_count))
        counter += 1

    # Условие  выбора регистратора:  ############################################################################
    if values["-4PORTS1-"]:
        reg_count = 4
        ports = 4

    if values["-8PORTS1-"]:
        reg_count = 5
        ports = 8

    if values["-16PORTS1-"]:
        reg_count = 6
        ports = 16

    if values["-32PORTS1-"]:
        reg_count = 22
        ports = 32
    if int(quantity) > ports:
        sg.popup("Внимание! Портов регистратора меньше, чем камер")




    with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)
        writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                         int(sheet[f"B{reg_count}"].value) * 1))
        counter += 1

        ##Бесперебойник###########

    if values['-PB-YES-2MP-']:
        power_backup = int(sheet[f"B{power_backup_index}"].value)
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{power_backup_index}"].value, sheet[f"B{power_backup_index}"].value, 1,
                             int(sheet[f"B{power_backup_index}"].value) * 1))
            counter += 1




        # подсчет суммы :  ####################################################################################
    hdd_sum = int(sheet[f"B{hdd_size}"].value) * hdd_count

    list2 = [reg_count]
    sum1 = 0
    sum2 = 0

    for i in list:
        sum1 = int(sheet[f"B{i}"].value)*int(quantity)+sum1


    for i in list2:
        sum2 = int(sheet[f"B{i}"].value) + sum2

    # расчет мощности блоков питания(2А или 3А)
    power_supply = 0

    if int(quantity) < 2:
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{14}"].value, sheet[f"B{14}"].value, 1,
                             int(sheet[f"B{14}"].value) * 1))
            power_supply = int(sheet[f"B{14}"].value)
    elif 2 <= int(quantity) < 4:
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, 1,
                             int(sheet[f"B{13}"].value) * 1))
        power_supply = int(sheet[f"B{13}"].value)
    else:
        #расчет количества блоков питания (power suply quantity(PSQ))    ########################################

        psq = math.ceil(int(quantity)/7)
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, psq,
                             int(sheet[f"B{13}"].value) * psq))
        power_supply = int(sheet[f"B{13}"].value) * psq

    #Всего кабеля

    sum_cab = int(cable)*int(sheet[f"B{17}"].value)

    #Итого
    with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(("","Итого","","", sum1 + sum2 + sum_cab + power_supply + hdd_sum + power_backup))
########################################################################################################################

def cam_calc_2(counter,quantity,hdd_count,hdd_size,power_backup):
    print("Вариант 5MP...")


    list = [3, 15, 16, 18]

    for i in list:
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{i}"].value, sheet[f"B{i}"].value, quantity,
                             int(sheet[f"B{i}"].value) * int(quantity)))
            counter += 1

    # Единичные товары(жесткий, БП)
    with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)
        # Считаем кабель
        writer.writerow(
            (counter, sheet[f"a{17}"].value, sheet[f"B{17}"].value, cable, int(sheet[f"B{17}"].value) * int(cable)))
        counter += 1

        # Жесткий диск
        # Размер жесткого диска

    if values['-1TB-HDD-5MP-']:
        hdd_size = 11

    elif values['-2TB-HDD-5MP-']:
        hdd_size = 25

    elif values['-3TB-HDD-5MP-']:
        hdd_size = 26

    elif values['-4TB-HDD-5MP-']:
        hdd_size = 27

    elif values['-5TB-HDD-5MP-']:
        hdd_size = 28

    elif values['-6TB-HDD-5MP-']:
        hdd_size = 29

    elif values['-7TB-HDD-5MP-']:
        hdd_size = 30

    elif values['-8TB-HDD-5MP-']:
        hdd_size = 31

        # количество жестких дисков


    if values['-1ITEMS-HDD-5MP-']:
        hdd_count = 1
    elif values['-2ITEMS-HDD-5MP-']:
        hdd_count = 2

    with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)
        writer.writerow((counter, sheet[f"a{hdd_size}"].value, sheet[f"B{hdd_size}"].value, hdd_count,
                         int(sheet[f"B{hdd_size}"].value) * hdd_count))
        counter += 1


        # Условие  выбора регистратора:  ############################################################################
        if values["-4PORTS2-"]:
            reg_count = 4
            ports = 4

        if values["-8PORTS2-"]:
            reg_count = 5
            ports = 8

        if values["-16PORTS2-"]:
            reg_count = 6
            ports = 16

        if values["-32PORTS2-"]:
            reg_count = 22
            ports = 32
        if int(quantity) > ports:
            # sg.popup_yes_no("Внимание! количество камер больше , чем выбрано портов  регистратора")
            sg.popup("Внимание! Портов регистратора меньше, чем камер")

    with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)
        writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                         int(sheet[f"B{reg_count}"].value) * 1))
        counter += 1

        ##Бесперебойник###########

    if values['-PB-YES-5MP-']:
        power_backup = int(sheet[f"B{power_backup_index}"].value)
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow(
                (counter, sheet[f"a{power_backup_index}"].value, sheet[f"B{power_backup_index}"].value, 1,
                 int(sheet[f"B{power_backup_index}"].value) * 1))
            counter += 1

        # подсчет суммы :  ####################################################################################

    hdd_sum = int(sheet[f"B{hdd_size}"].value) * hdd_count
    list2 = [reg_count]
    sum1 = 0
    sum2 = 0

    for i in list:
        sum1 = int(sheet[f"B{i}"].value) * int(quantity) + sum1

    for i in list2:
        sum2 = int(sheet[f"B{i}"].value) + sum2

    # расчет мощности блоков питания(2А или 3А)
    power_supply = 0

    if int(quantity) < 2:
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{14}"].value, sheet[f"B{14}"].value, 1,
                             int(sheet[f"B{14}"].value) * 1))
            power_supply = int(sheet[f"B{14}"].value)
    elif 2 <= int(quantity) < 4:
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, 1,
                             int(sheet[f"B{13}"].value) * 1))
        power_supply = int(sheet[f"B{13}"].value)
    else:
        # расчет количества блоков питания (power suply quantity(PSQ))    ########################################

        psq = math.ceil(int(quantity) / 7)
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, psq,
                             int(sheet[f"B{13}"].value) * psq))
        power_supply = int(sheet[f"B{13}"].value) * psq

    # Всего кабеля

    sum_cab = int(cable) * int(sheet[f"B{17}"].value)

    # Итого
    with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(("", "Итого", "", "", sum1 + sum2 + sum_cab + power_supply + hdd_sum + power_backup))
########################################################################################################################

def ip_cam_calc_3(counter,quantity,hdd_count,hdd_size,power_backup):
    print("Вариант IP 2MP...")
    counter = 1
    list = [9, 21, 16, 18]
    for i in list:
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{i}"].value, sheet[f"B{i}"].value, quantity,
                             int(sheet[f"B{i}"].value) * int(quantity)))
            counter += 1
    # Единичные товары(жесткий, БП)
    with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)

        # Считаем кабель
        writer.writerow(
            (counter, sheet[f"a{24}"].value, sheet[f"B{24}"].value, cable, int(sheet[f"B{24}"].value) * int(cable)))
        counter += 1
        # Жесткий диск

        # Размер жесткого диска
    if values['-1TB-HDD-IP2MP-']:
        hdd_size = 11

    elif values['-2TB-HDD-IP2MP-']:
        hdd_size = 25

    elif values['-3TB-HDD-IP2MP-']:
        hdd_size = 26

    elif values['-4TB-HDD-IP2MP-']:
        hdd_size = 27

    elif values['-5TB-HDD-IP2MP-']:
        hdd_size = 28

    elif values['-6TB-HDD-IP2MP-']:
        hdd_size = 29

    elif values['-7TB-HDD-IP2MP-']:
        hdd_size = 30

    elif values['-8TB-HDD-IP2MP-']:
        hdd_size = 31

    if values['-1ITEMS-HDD-IP2MP-']:
        hdd_count = 1
    elif values['-2ITEMS-HDD-IP2MP-']:
        hdd_count = 2

    with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",newline="") as file:
        writer = csv.writer(file)
        writer.writerow((counter, sheet[f"a{hdd_size}"].value, sheet[f"B{hdd_size}"].value, hdd_count,int(sheet[f"B{hdd_size}"].value) * hdd_count))
        counter += 1

        # Условие  выбора регистратора:  ############################################################################
        if values["-4PORTS3-"]:
            reg_count = 4
            ports = 4

        if values["-8PORTS3-"]:
            reg_count = 5
            ports = 8


        if values["-9PORTS3-"]:
            reg_count = 23
            ports = 9

        if values["-16PORTS3-"]:
            reg_count = 6
            ports = 16

        if values["-32PORTS3-"]:
            reg_count = 22
            ports = 32
        if int(quantity) > ports:
            sg.popup("Внимание! количество камер больше , чем выбрано портов  регистратора")

    with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)
        writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                         int(sheet[f"B{reg_count}"].value) * 1))
        counter += 1

        ##Бесперебойник###########

    if values['-PB-YES-IP2MP-']:
        power_backup = int(sheet[f"B{power_backup_index}"].value)
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow(
                (counter, sheet[f"a{power_backup_index}"].value, sheet[f"B{power_backup_index}"].value, 1,
                 int(sheet[f"B{power_backup_index}"].value) * 1))
            counter += 1

        ############## подсчет суммы

    hdd_sum = int(sheet[f"B{hdd_size}"].value) * hdd_count
    list2 = [reg_count]
    sum1 = 0
    sum2 = 0

    for i in list:
        sum1 = int(sheet[f"B{i}"].value) * int(quantity) + sum1

    for i in list2:
        sum2 = int(sheet[f"B{i}"].value) + sum2

    # расчет мощности блоков питания(2А или 3А)
    power_supply = 0

    if int(quantity) < 2:
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{14}"].value, sheet[f"B{14}"].value, 1,
                             int(sheet[f"B{14}"].value) * 1))
            power_supply = int(sheet[f"B{14}"].value)
    elif 2 <= int(quantity) < 4:
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, 1,
                             int(sheet[f"B{13}"].value) * 1))
        power_supply = int(sheet[f"B{13}"].value)
    else:
        # расчет количества блоков питания (power suply quantity(PSQ))    ########################################

        psq = math.ceil(int(quantity) / 7)
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, psq,
                             int(sheet[f"B{13}"].value) * psq))
        power_supply = int(sheet[f"B{13}"].value) * psq

    # Всего кабеля

    sum_cab = int(cable) * int(sheet[f"B{17}"].value)

    # Итого
    with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(("", "Итого", "", "", sum1 + sum2 + sum_cab + power_supply + hdd_sum + power_backup))

########################################################################################################################

def ip_cam_calc_4(counter,quantity,hdd_count,hdd_size,power_backup):
    print("Вариант IP 5MP...")

    counter = 1
    list = [10, 21, 16, 18]
    for i in list:
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{i}"].value, sheet[f"B{i}"].value, quantity,
                             int(sheet[f"B{i}"].value) * int(quantity)))
            counter += 1

    # Единичные товары(жесткий, БП)
    with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)

        # Считаем кабель
        writer.writerow(
            (counter, sheet[f"a{24}"].value, sheet[f"B{24}"].value, cable, int(sheet[f"B{24}"].value) * int(cable)))
        counter += 1
        # Жесткий диск
        # Размер жесткого диска
    if values['-1TB-HDD-IP5MP-']:
        hdd_size = 11

    elif values['-2TB-HDD-IP5MP-']:
        hdd_size = 25

    elif values['-3TB-HDD-IP5MP-']:
        hdd_size = 26

    elif values['-4TB-HDD-IP5MP-']:
        hdd_size = 27

    elif values['-5TB-HDD-IP5MP-']:
        hdd_size = 28

    elif values['-6TB-HDD-IP5MP-']:
        hdd_size = 29

    elif values['-7TB-HDD-IP5MP-']:
        hdd_size = 30

    elif values['-8TB-HDD-IP5MP-']:
        hdd_size = 31


    if values['-1ITEMS-HDD-IP5MP-']:
        hdd_count = 1
    elif values['-2ITEMS-HDD-IP5MP-']:
        hdd_count = 2

    with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
              newline="") as file:
        writer = csv.writer(file)
        writer.writerow((counter, sheet[f"a{hdd_size}"].value, sheet[f"B{hdd_size}"].value, hdd_count,
                         int(sheet[f"B{hdd_size}"].value) * hdd_count))
        counter += 1

        # Условие  выбора регистратора:  ############################################################################
        if values["-4PORTS4-"]:
            reg_count = 4
            ports = 4

        if values["-8PORTS4-"]:
            reg_count = 5
            ports = 8

        if values["-9PORTS4-"]:
            reg_count = 23
            ports = 9

        if values["-16PORTS4-"]:
            reg_count = 6
            ports = 16

        if values["-32PORTS4-"]:
            reg_count = 22
            ports = 32
        if int(quantity) > ports:
            # sg.popup_yes_no("Внимание! количество камер больше , чем выбрано портов  регистратора")
            sg.popup("Внимание! Портов регистратора меньше, чем камер")

        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{reg_count}"].value, sheet[f"B{reg_count}"].value, 1,
                             int(sheet[f"B{reg_count}"].value) * 1))
            counter += 1

            ##Бесперебойник###########

            if values['-PB-YES-IP5MP-']:
                power_backup = int(sheet[f"B{power_backup_index}"].value)

                with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                          newline="") as file:
                    writer = csv.writer(file)
                    writer.writerow(
                        (counter, sheet[f"a{power_backup_index}"].value, sheet[f"B{power_backup_index}"].value, 1,
                         int(sheet[f"B{power_backup_index}"].value) * 1))
                    counter += 1


    # подсчет суммы

    hdd_sum = int(sheet[f"B{hdd_size}"].value) * hdd_count
    list2 = [reg_count]
    sum1 = 0
    sum2 = 0

    for i in list:
        sum1 = int(sheet[f"B{i}"].value) * int(quantity) + sum1

    for i in list2:
        sum2 = int(sheet[f"B{i}"].value) + sum2

    # расчет мощности блоков питания(2А или 3А)
    power_supply = 0

    if int(quantity) < 2:
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{14}"].value, sheet[f"B{14}"].value, 1,
                             int(sheet[f"B{14}"].value) * 1))
            power_supply = int(sheet[f"B{14}"].value)
    elif 2 <= int(quantity) < 4:
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, 1,
                             int(sheet[f"B{13}"].value) * 1))
        power_supply = int(sheet[f"B{13}"].value)
    else:
        # расчет количества блоков питания (power suply quantity(PSQ))    ########################################

        psq = math.ceil(int(quantity) / 7)
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{13}"].value, sheet[f"B{13}"].value, psq,
                             int(sheet[f"B{13}"].value) * psq))
        power_supply = int(sheet[f"B{13}"].value) * psq

    # Всего кабеля

    sum_cab = int(cable) * int(sheet[f"B{17}"].value)

    # Итого
    with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(("", "Итого", "", "", sum1 + sum2 + sum_cab + power_supply + hdd_sum+power_backup))

########################################################################################################################

def ip_sd_calc_5(counter,quantity):
    print("Вариант IP-Compac ...")
    counter = 1
    list = [7,12,16]
    cam_sum = 0
    for i in list:
        with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8",
                  newline="") as file:
            writer = csv.writer(file)
            writer.writerow((counter, sheet[f"a{i}"].value, sheet[f"B{i}"].value, quantity,
                             int(sheet[f"B{i}"].value) * int(quantity)))
            cam_sum = (int(sheet[f"B{i}"].value) * int(quantity)) + cam_sum
            counter += 1

    with open(f"ready/Видеонаблюдение для {client.capitalize()} на {today_is()}.csv", "a", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(("", "Итого", "", "", cam_sum))
########################################################################################################################





#MAIN LOOP

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break

    if event == '-CONFIRM_AHD_2MP-':

        client = values["-CLIENT-IN-2MP-"]
        if client == "":
            client = "somebody_someone"
        cable = values["-CABLE-IN-2MP-"]
        if not cable.isnumeric():
            cable = "0"

        quantity = values["-CAM-IN-2MP-"]

        if quantity == "":
            quantity = 1
            sg.popup("Количество камер не указано!\nБудет сформирована смета для ОДНОЙ камеры!")
            # break
        elif int(quantity) <= 0:
            sg.popup("Пардон, сударь, но, похоже, вы гоните ! Количество камер меньше 1")
            ########################################################################################

    if event == '-CONFIRM_AHD_5MP-':

        client = values["-CLIENT-IN-5MP-"]
        if client == "":
            client = "somebody_someone"
        cable = values["-CABLE-IN-5MP-"]
        if not cable.isnumeric():
            cable = "0"
        quantity = values["-CAM-IN-5MP-"]

        if quantity == "":
            quantity = 1
            sg.popup("Количество камер не указано!\nБудет сформирована смета для ОДНОЙ камеры!")
        elif int(quantity) <= 0:
            sg.popup("Пардон, сударь, но, похоже, вы гоните ! Количество камер меньше 1")
###############################################################################################

    if event == '-CONFIRM_IP_2MP-':

        client = values["-CLIENT-IN-IP2MP-"]
        if client == "":
            client = "somebody_someone"

        cable = values["-CABLE-IN-IP2MP-"]
        if not cable.isnumeric():
            cable = "0"

        quantity = values["-CAM-IN-IP2MP-"]

        if quantity == "":
            quantity = 1
            sg.popup("Количество камер не указано!\nБудет сформирована смета для ОДНОЙ камеры!",auto_close=True)
        elif int(quantity) <= 0:
            sg.popup("Пардон, сударь, но, похоже, вы гоните ! Количество камер меньше 1")
  ###########################################################################################

    if event == '-CONFIRM_IP_5MP-':

        client = values["-CLIENT-IN-IP5MP-"]
        if client == "":
            client = "somebody_someone"

        cable = values["-CABLE-IN-IP5MP-"]
        if not cable.isnumeric():
            cable = "0"

        quantity = values["-CAM-IN-IP5MP-"]

        if quantity == "":
            quantity = 1
            sg.popup("Количество камер не указано!\nБудет сформирована смета для ОДНОЙ камеры!")
        elif int(quantity) <= 0:
            sg.popup("Пардон, сударь, но, похоже, вы гоните ! Количество камер меньше 1")
            #################################################################################

    if event == '-CONFIRM_COMPAC-':
        client = values["-CLIENT-IN-COMP-"]
        if client == "":
            client = "somebody_someone"

        cable = values["-CABLE-IN-COMP-"]
        if not cable.isnumeric():
            cable = "0"

        quantity = values["-CAM-IN-COMP-"]

        if quantity == "":
            quantity = 1
            sg.popup("Количество камер не указано!\nБудет сформирована смета для ОДНОЙ камеры!",auto_close=True)
        elif int(quantity) <= 0:
            sg.popup("Пардон, сударь, но, похоже, вы гоните ! Количество камер меньше 1")


    # with open(f"ready/Видеонаблюдение для {client.title()} на {today_is()}.csv", "w", encoding="utf-8", newline="") as file:
    #     writer = csv.writer(file)
    #     writer.writerow(("№", "Наименование", "Цена", "количество", "Сумма!"))

    if event == "-CONFIRM_AHD_2MP-":
        cam_calc_1(counter,quantity,hdd_count,hdd_size,power_backup)
        os.startfile(os.getcwd()+"/ready")

    elif event == "-CONFIRM_AHD_5MP-":
        cam_calc_2(counter,quantity,hdd_count,hdd_size,power_backup)
        os.startfile(os.getcwd() + "/ready")

    elif event == "-CONFIRM_IP_2MP-":
        ip_cam_calc_3(counter,quantity,hdd_count,hdd_size,power_backup)
        os.startfile(os.getcwd() + "/ready")

    elif event == "-CONFIRM_IP_5MP-":
        ip_cam_calc_4(counter,quantity,hdd_count,hdd_size,power_backup)
        os.startfile(os.getcwd() + "/ready")

    elif event == '-CONFIRM_COMPAC-':
        ip_sd_calc_5(counter,quantity)
        os.startfile(os.getcwd() + "/ready")

    # window["--FINAL_MESSAGE--"].update(visible=True)
    # sg.popup(f"Файл Видеонаблюдение для {client.title()} на {today_is()}.csv сформирован!")




window.close()



