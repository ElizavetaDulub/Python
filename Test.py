from xml.dom import minidom
import pandas as pd
import datetime as dt
import os

import xml.etree.ElementTree as ET

directory = r'D:\\Python\\CreditReports'
first_level_node = ['credittransaction', 'LeasingTransaction']
second_level_node = ['lastpresentation','latesum', 'latepercent', 'LateLeasingSum']


def parse_sum(node_2):
    for node_3 in node_2.childNodes:
        if node_3.nodeName == "rest":
            # запись суммы просрочки по основному долгу
            print("Late sum: " + str(node_3.childNodes[0].nodeValue))
            return float(node_3.childNodes[0].nodeValue)


def parse_node_1(node_1, df):
    latesum, latepercent, date, client_number, client_name, sign_date = None, None, None, None, None, None
    if node_1.nodeName == 'contractnumber':
        latesum = 'ДОГОВОР ' + node_1.childNodes[0].nodeValue
        df.loc[df.shape[0]] = [latesum, latepercent, date, client_name, client_number, sign_date]
    for i, item in enumerate(first_level_node):
        if node_1.nodeName == item:
            for node_2 in node_1.childNodes:
                if node_2.nodeName == "lastpresentation":
                        # запись даты
                        print(str(node_2.childNodes[0].nodeValue))
                        date = dt.datetime.strptime(node_2.childNodes[0].nodeValue, "%d.%m.%Y").date()
                elif node_2.nodeName == "latesum":
                        latesum = parse_sum(node_2)
                elif node_2.nodeName == "latepercent":
                        latepercent = parse_sum(node_2)
                elif node_2.nodeName == "LateLeasingSum":
                        latesum = parse_sum(node_2)

            df.loc[df.shape[0]] = [latesum, latepercent, date, client_name, client_number, sign_date]


def parse_client_info(doc, df):
    client_info = doc.getElementsByTagName('Response')
    latesum, latepercent, date, client_number, client_name, sign_date = None, None, None, None, None, None
    for client in client_info:
        if client.getAttribute('type') == '11012' and client.getAttribute('name') == 'getfullhistoryfiz':
            print('ФИЗЛИЦО')
            for number in doc.getElementsByTagName('IDNumber'):
                client_number = number.childNodes[0].nodeValue
                print(client_number)
            for name in doc.getElementsByTagName('FIO'):
                client_name =str()
                for i in name.childNodes:
                    client_name += i.childNodes[0].nodeValue + ' '
                    print(client_name)
        elif client.getAttribute('type') == '11022' and client.getAttribute('name') == 'getfullhistoryjur':
            print('ЮРЛИЦО')
            for number in doc.getElementsByTagName('UNP'):
                client_number = number.childNodes[0].nodeValue
                print(client_number)
            for name in doc.getElementsByTagName('name'):
                client_name = name.childNodes[0].nodeValue
                print(client_name)
    for s_date in doc.getElementsByTagName('sign_time'):
        sign_date = s_date.childNodes[0].nodeValue
        print(sign_date)
    df.loc[df.shape[0]] = [latesum, latepercent, date, client_name, client_number, sign_date]
    return client_number


def parse_reports(file, df):
    doc = minidom.parse(file)
    reports_tag = doc.getElementsByTagName("contract")
    print(type(reports_tag))

    client_number = parse_client_info(doc, df)

    for contract in reports_tag:
        for node_1 in contract.childNodes:
            parse_node_1(node_1, df)

    writer = pd.ExcelWriter('dataframe.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name=client_number, startrow=1, header=False, index=False, freeze_panes=(2, 0))


    # writer = pd.ExcelWriter('dataframe.xlsx', engine='openpyxl', mode='a', if_sheet_exists='new')
    # df.to_excel(writer, sheet_name=client_number, startrow=1, header=False, index=False, freeze_panes=(2, 0))

    workbook = writer.book
    worksheet = writer.sheets[client_number]

    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'vcenter',
        'align' : 'center',
        'fg_color': '#D7E4BC',
        'border': 1
    })

    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    writer.close()
    # df.style.set_table_styles([headers]).to_excel("dataframe.xlsx", index=False, freeze_panes=(1, 0))


for file in os.scandir(directory):
    if file.name.endswith('.xml'):
        print(file.path)
        df = pd.DataFrame(columns=['Сумма основного долга', 'Сумма процентов', 'Дата', 'УНП/Индентиф. номер',
                                   'Наименование клиента', 'Дата формирования отчёта'])
        parse_reports(file.path, df)



# flag_dolg, flag_proc = False, False
# count_dolg, count_proc, amount_days = 0, 0, 0
# date_dolg_in = None
#
#
# for index, row in data.iterrows():
#     # dolg
#     if row['dolg'] == 0.0 and flag_dolg:
#         count_dolg += 1
#         flag_dolg = False
#         diff_date = index - date_dolg_in  # amount of days
#         if diff_date > 30:
#             amount_days += 1
#
#     if row['dolg'] != None and not flag_dolg and row['dolg'] != 0.0:
#         flag_dolg = True
#         date_dolg_in = index
#
#     # procent
#     if row['procent'] == 0.0 and flag_proc:
#         count_proc += 1
#         flag_proc = False
#
#     if row['procent'] != None and not flag_proc and row['procent'] != 0.0:
#         flag_proc = True