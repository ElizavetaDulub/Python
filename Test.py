from xml.dom import minidom
import pandas as pd
import datetime as dt
import os
import numpy as np
from lxml import html, etree
from bs4 import BeautifulSoup


# with open('file_legal_entity.html', 'r', encoding='latin-1') as inp:
#     soup = BeautifulSoup(inp, 'html.parser')
#
# # Split the document by lines and join the lines
# # from index 1 to remove the doctype Html as it is
# # present in index 0 from the parsed document.
# lines = soup.prettify().splitlines()
# content = "\n".join(lines[1:])
#
# # Open a output.xml file and write the modified content.
# with open("output.xml", 'w', encoding='latin-1') as out:
#     out.write(content)
# with open('file_legal_entity.html', 'r', encoding='utf-8') as inp:
#     htmldoc = html.fromstring(inp.read())
#
# with open("output.xml", 'wb') as out:
#     out.write(etree.tostring(htmldoc))

directory = r'D:\\Python\\CreditReports'
first_level_node = ['credittransaction', 'LeasingTransaction']


def parse_sum(transaction_child):
    for rest in transaction_child.childNodes:
        if rest.nodeName == "rest":
            # запись суммы просрочки по основному долгу
            # print("Late sum: " + str(rest.childNodes[0].nodeValue))
            return float(rest.childNodes[0].nodeValue)


def parse_node_1(contract_child, df):
    latesum, latepercent, date, client_number, client_name, sign_date = None, None, None, None, None, None
    if contract_child.nodeName == 'contractnumber':
        client_name = 'ДОГОВОР ' + contract_child.childNodes[0].nodeValue
        print(client_name)
        df.loc[df.shape[0]] = [latesum, latepercent, date, client_name, client_number, sign_date]
    for i, item in enumerate(first_level_node):
        if contract_child.nodeName == item:
            for transaction_child in contract_child.childNodes:
                if transaction_child.nodeName == "lastpresentation":
                        # запись даты
                        # print(str(transaction_child.childNodes[0].nodeValue))
                    date = dt.datetime.strptime(transaction_child.childNodes[0].nodeValue, "%d.%m.%Y").date()
                elif transaction_child.nodeName == "latesum":
                    latesum = parse_sum(transaction_child)
                elif transaction_child.nodeName == "latepercent":
                    latepercent = parse_sum(transaction_child)
                elif transaction_child.nodeName == "LateLeasingSum":
                    latesum = parse_sum(transaction_child)
            df.loc[df.shape[0]] = [latesum, latepercent, date, client_name, client_number, sign_date]
    return df


def delay(temp_df, df, delay_df):
    flag_latesum, flag_latepercent = False, False
    delay_latesum, delay_perc, amount_days, diff_date = 0, 0, 0, 0
    date_dolg_in = None
    last_date = df.loc[df.index[0], 'Дата формирования отчёта'] - dt.timedelta(weeks=52)
    print('last date', last_date)
    temp_df = temp_df.sort_values(by=['Дата'], ascending=True).fillna(np.nan).replace([np.nan], [None])
    print(temp_df.loc[:, 'Сумма основного долга':'Дата'])

    for ind in range(temp_df.shape[0]):
        row = temp_df.iloc[ind]
        if row['Дата'] is None:
            continue
        if row['Дата'] < last_date:
            continue
        if row['Сумма основного долга'] == 0.0 and flag_latesum:
            delay_latesum += 1
            flag_latesum = False
            diff_date = row['Дата'] - date_dolg_in
            if diff_date > dt.timedelta(days=30):
                amount_days += 1
            print('Флаг опустился: просрочка погашена')
            print(row['Сумма основного долга'])
        elif row['Сумма основного долга'] is not None and not flag_latesum and \
                row['Сумма основного долга'] != 0.0:
            flag_latesum = True
            date_dolg_in = row['Дата']
            print('Флаг поднялся: клиент вышел на просрочку')
            print(row['Сумма основного долга'])
        if row['Сумма процентов'] == 0.0 and flag_latepercent:
            delay_perc += 1
            flag_latepercent = False
            diff_date = row['Дата'] - date_dolg_in
            if diff_date > dt.timedelta(days=30):
                amount_days += 1
            print('Флаг опустился: просрочка погашена')
            print(row['Сумма процентов'])
        elif row['Сумма процентов'] is not None and not flag_latepercent and \
                row['Сумма процентов'] != 0.0:
            flag_latepercent = True
            date_dolg_in = row['Дата']
            print('Флаг поднялся: клиент вышел на просрочку')
            print(row['Сумма процентов'])
    if flag_latesum:
        delay_latesum += 1
        diff_date = temp_df.iloc[-2, 2] - date_dolg_in
        if diff_date > dt.timedelta(days=30):
            amount_days += 1
    if flag_latepercent:
        delay_perc += 1
        diff_date = temp_df.iloc[-2, 2] - date_dolg_in
        if diff_date > dt.timedelta(days=30):
            amount_days += 1
    delay_sum = delay_latesum + delay_perc
    if delay_sum > 4:
        print('ПРОСРОЧКА БОЛЕЕ 4 РАЗ!!!!!!!!!')
    delay_df.loc[delay_df.shape[0]] = [delay_sum, amount_days]
    return print('Количество задолженностей', delay_latesum, delay_perc, '  ', 'Больше 30 дней', amount_days)


def parse_client_info(doc, df, df_final):
    client_info = doc.getElementsByTagName('Response')
    latesum, latepercent, date, client_number, client_name, sign_date = None, None, None, None, None, None
    delay_sum, count_delay_sum, delay_30days, count_delay_30days, credit_history, conclusion = \
        None, None, None, None, None, None
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
        sign_date = dt.datetime.fromisoformat(s_date.childNodes[0].nodeValue).date()

    df.loc[df.shape[0]] = [latesum, latepercent, date,  client_number, client_name, sign_date]

    return sign_date, client_number, client_name


def parse_reports(file):
    client_number, client_name, sign_date = None, None, None
    index, delay_sum, count_delay_sum, delay_30days, count_delay_30days, credit_history, conclusion = \
        None, None, None, None, None, None, None
    df = pd.DataFrame(columns=['Сумма основного долга', 'Сумма процентов', 'Дата', 'УНП/Индентиф. номер',
                               'Наименование клиента', 'Дата формирования отчёта'])
    df_final = pd.DataFrame(columns=['№ п/п', 'Дата формирования отчета', 'УНП/ личный номер', 'Наименование клиента',
                                     'Имелись факты наличия просрочек за последние 12 месяцев по погашению '
                                     'основного долга и/или процентов (лизинговых платежей) более 4 раз',
                                     'Имелись факты наличия просрочек за последние 12 месяцев по погашению основного '
                                     'долга и/или процентов (лизинговых платежей) продолжительностью более 30 '
                                     'календарных дней по любой из просрочек', 'На текущую дату у Клиента имеется '
                                     'просроченная задолженность по любому из договоров кредитного характера*',
                                     'Отсутствие кредитной истории', 'Вывод об оценке кредитной истории'])
    doc = minidom.parse(file)
    contract_tags = doc.getElementsByTagName("contract")
    print(type(contract_tags))

    sign_date, client_number, client_name = parse_client_info(doc, df, df_final)
    delay_df = pd.DataFrame(columns=['Delays count', 'Delays count 30days'])

    for contract in contract_tags:
        temp_df = pd.DataFrame(columns=['Сумма основного долга', 'Сумма процентов', 'Дата', 'УНП/Индентиф. номер',
                               'Наименование клиента', 'Дата формирования отчёта'])
        if contract.parentNode.nodeName == 'CollateralContractList' \
                or contract.parentNode.nodeName == 'SuretyContractList':
            continue
        else:
            for contact_child in contract.childNodes:
                temp_df = parse_node_1(contact_child, temp_df)
        df = pd.concat([df, temp_df], ignore_index=True)
        delay(temp_df, df, delay_df)
    count_delay_sum = sum(delay_df['Delays count'].to_list())
    if count_delay_sum > 4:
        delay_sum = 'ДА'
    else:
        delay_sum = 'НЕТ'
    count_delay_30days = sum(delay_df['Delays count 30days'].to_list())
    if count_delay_30days > 0:
        delay_30days = 'ДА'
    else:
        delay_30days = 'НЕТ'
    df_final.loc[df_final.shape[0]] = [sign_date, client_number, client_name, delay_sum, count_delay_sum,
                                       delay_30days, count_delay_30days, credit_history, conclusion]

    return df, df_final, client_number

def df_writer(df, df_final, client_number):
    writer = pd.ExcelWriter('dataframe.xlsx', engine='xlsxwriter')
    file1 = 'dataframe.xlsx'
    print(os.path.abspath(file1))
    if os.path.exists(r'D:\Python\dataframe.xlsx'):
        df = pd.read_excel(r'D:\Python\dataframe.xlsx')
        df.to_excel(writer, sheet_name=client_number, startrow=1, header=False, index=False, freeze_panes=(2, 0))
    writer_final = pd.ExcelWriter('dataframe_final.xlsx', engine='xlsxwriter')
    df_final.to_excel(writer_final, startrow=1, header=False, freeze_panes=(1, 0))

    # writer = pd.ExcelWriter('dataframe.xlsx', engine='openpyxl', mode='a', if_sheet_exists='new')
    # df.to_excel(writer, sheet_name=client_number, startrow=1, header=False, index=False, freeze_panes=(2, 0))

    workbook = writer.book
    workbook_final = writer_final.book
    worksheet = writer.sheets[client_number]
    worksheet_final = writer_final.sheets['Sheet1']

    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'vcenter',
        'align': 'center',
        'fg_color': '#D7E4BC',
        'border': 1
    })
    header_format_final = workbook_final.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'vcenter',
        'align': 'center',
        'fg_color': '#D7E4BC',
        'border': 1
    })

    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
    for col_num, value in enumerate(df_final.columns.values):
        worksheet_final.write(0, col_num, value, header_format_final)

    writer.close()
    writer_final.close()


for file in os.scandir(directory):
    if file.name.endswith('.xml'):
        print(file.path)
        df, df_final, client_number = parse_reports(file.path)
        df_writer(df, df_final, client_number)


# flag_dolg, flag_proc = False, False
#     count_dolg, count_proc, amount_days = 0, 0, 0
#     date_dolg_in = None
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
