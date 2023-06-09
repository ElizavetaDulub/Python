import datetime as dt
import os
from xml.dom import minidom
import numpy as np
import pandas as pd

directory = r'D:\\Python\\CreditReports'
first_level_node = ['credittransaction', 'LeasingTransaction']


class ParseDocs:

    def __init__(self, file_path):
        self.final_index = 0
        # self.writer = pd.ExcelWriter(file1, engine='xlsxwriter')
        self.file1 = file_path

    def parse_sum(self, transaction_child):
        for rest in transaction_child.childNodes:
            if rest.nodeName == "rest":
                # запись суммы просрочки по основному долгу
                return float(rest.childNodes[0].nodeValue)

    def df_format(self, writer, client_number, df):
        workbook = writer.book
        worksheet = writer.sheets[client_number]
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

    def df_final_format(self, writer_final, df_final):
        workbook_final = writer_final.book
        worksheet_final = writer_final.sheets['Sheet1']

        header_format_final = workbook_final.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'fg_color': '#D7E4BC',
            'border': 1
        })

        worksheet_final.merge_range(0, 4, 0, 5, 'Имелись факты наличия просрочек за последние 12 месяцев по погашению '
                                                'основного долга и/или процентов (лизинговых платежей)'
                                                ' более 4 раз', header_format_final)
        worksheet_final.merge_range(0, 6, 0, 7, 'Имелись факты наличия просрочек за последние 12 месяцев по погашению '
                                                'основного долга и/или процентов (лизинговых платежей) продолжительностью '
                                                'более 30 календарных дней по любой из просрочек', header_format_final)
        worksheet_final.set_row(0, 110)
        worksheet_final.set_column('A:K', 20)

        for col_num, value in enumerate(df_final.columns.values):
            worksheet_final.write(0, col_num, value, header_format_final)

    def parse_node_1(self, contract_child, df):
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
                        latesum = self.parse_sum(transaction_child)
                    elif transaction_child.nodeName == "latepercent":
                        latepercent = self.parse_sum(transaction_child)
                    elif transaction_child.nodeName == "LateLeasingSum":
                        latesum = self.parse_sum(transaction_child)
                df.loc[df.shape[0]] = [latesum, latepercent, date, client_name, client_number, sign_date]
        return df

    def debt_calculation(self, temp_df, df, delay_df):
        flag_latesum, flag_latepercent = False, False
        delay_latesum, delay_perc, amount_days, diff_date, current_delay = 0, 0, 0, 0, 0
        date_dolg_in = None
        last_date = df.loc[df.index[0], 'Дата формирования отчёта'] - dt.timedelta(weeks=52)
        print('last date', last_date)
        temp_df = temp_df.sort_values(by=['Дата'], ascending=True).fillna(np.nan).replace([np.nan], [None])
        # print(temp_df.loc[:, 'Сумма основного долга':'Дата'])

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
            current_delay += 1
            diff_date = temp_df.iloc[-2, 2] - date_dolg_in
            if diff_date > dt.timedelta(days=30):
                amount_days += 1
        if flag_latepercent:
            delay_perc += 1
            current_delay += 1
            diff_date = temp_df.iloc[-2, 2] - date_dolg_in
            if diff_date > dt.timedelta(days=30):
                amount_days += 1
        delay_sum = delay_latesum + delay_perc
        if delay_sum > 4:
            print('ПРОСРОЧКА БОЛЕЕ 4 РАЗ!!!!!!!!!')
        delay_df.loc[delay_df.shape[0]] = [delay_sum, amount_days]
        return delay_latesum, delay_perc, amount_days, current_delay

    def parse_client_info(self, doc, df, df_final):
        client_info = doc.getElementsByTagName('Response')
        latesum, latepercent, date, client_number, client_name, sign_date = None, None, None, None, None, None
        for client in client_info:
            if client.getAttribute('type') == '11012' and client.getAttribute('name') == 'getfullhistoryfiz':
                print('ФИЗЛИЦО')
                for number in doc.getElementsByTagName('IDNumber'):
                    client_number = number.childNodes[0].nodeValue
                    print(client_number)
                for name in doc.getElementsByTagName('FIO'):
                    client_name = str()
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
        df.loc[df.shape[0]] = [latesum, latepercent, date, client_number, client_name, sign_date]
        return sign_date, client_number, client_name

    def parse_reports(self, file):
        client_number, client_name, sign_date = None, None, None
        index, delay_sum, count_delay_sum, delay_30days, count_delay_30days, current_delay, credit_history, \
        conclusion = None, None, None, None, None, None, None, None
        df = pd.DataFrame(columns=['Сумма основного долга', 'Сумма процентов', 'Дата', 'УНП/Индентиф. номер',
                                   'Наименование клиента', 'Дата формирования отчёта'])
        df_final = pd.DataFrame(columns=['№ п/п',
                                         'Дата формирования отчета',
                                         'УНП/ личный номер',
                                         'Наименование клиента',
                                         'Имелись факты наличия просрочек за последние 12 месяцев по погашению основного '
                                         'долга и/или процентов (лизинговых платежей) более 4 раз',
                                         'Количество задолженностей',
                                         'Имелись факты наличия просрочек за последние 12 месяцев по погашению основного '
                                         'долга и/или процентов (лизинговых платежей) продолжительностью более 30 '
                                         'календарных дней по любой из просрочек', 'Количество задолженностей 30 дней',
                                         'На текущую дату у Клиента имеется просроченная задолженность по любому из '
                                         'договоров кредитного характера*',
                                         'Отсутствие кредитной истории', 'Вывод об оценке кредитной истории'])
        doc = minidom.parse(file)
        contract_tags = doc.getElementsByTagName("contract")

        sign_date, client_number, client_name = self.parse_client_info(doc, df, df_final)
        delay_df = pd.DataFrame(columns=['Delays count', 'Delays count 30days'])

        for contract in contract_tags:
            temp_df = pd.DataFrame(columns=['Сумма основного долга', 'Сумма процентов', 'Дата', 'УНП/Индентиф. номер',
                                            'Наименование клиента', 'Дата формирования отчёта'])
            if contract.parentNode.nodeName == 'CollateralContractList' \
                    or contract.parentNode.nodeName == 'SuretyContractList':
                continue
            else:
                for contact_child in contract.childNodes:
                    temp_df = self.parse_node_1(contact_child, temp_df)
            df = pd.concat([df, temp_df], ignore_index=True)
            delay_latesum, delay_perc, amount_days, current_delay = self.debt_calculation(temp_df, df, delay_df)

        count_delay_sum = sum(delay_df['Delays count'].to_list())
        count_delay_30days = sum(delay_df['Delays count 30days'].to_list())
        delay_sum = 'ДА' if count_delay_sum > 4 else 'НЕТ'
        current_delay = 'ДА' if current_delay > 0 else 'НЕТ'
        delay_30days = 'ДА' if count_delay_30days > 0 else 'НЕТ'

        df_final.loc[df_final.shape[0]] = [self.final_index, sign_date, client_number, client_name, delay_sum,
                                           count_delay_sum, delay_30days, count_delay_30days, current_delay,
                                           credit_history, conclusion]
        self.final_index += 1

        return df, df_final, client_number

    def df_writer(self, df, df_final, client_number):
        # df.to_excel(self.writer, sheet_name=client_number, startrow=1, header=False, index=False, freeze_panes=(2, 0))
        file2 = 'dataframe_final.xlsx'
        print(os.path.abspath(self.file1))
        if os.path.exists(r'D:\Python\dataframe.xlsx'):
            # df_old = pd.read_excel(r'D:\Python\dataframe.xlsx')
            # df = pd.concat([df, df_old], ignore_index=True)  # smth like that
            with pd.ExcelWriter(path=self.file1, mode='a', engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=client_number, startrow=1, header=False, index=False,
                            freeze_panes=(2, 0))
                # self.df_format(writer, client_number, df)

        else:
            with pd.ExcelWriter(path=self.file1, mode='wb') as writer:
                df.to_excel(writer, sheet_name=client_number, startrow=1, header=False, index=False,
                            freeze_panes=(2, 0))
                self.df_format(writer, client_number, df)

        if os.path.exists(r'D:\Python\dataframe_final.xlsx'):
            df_final_old = pd.read_excel(r'D:\Python\dataframe_final.xlsx')
            # writer_final = pd.ExcelWriter(file2, engine='xlsxwriter')
            df_final = pd.concat([df_final_old, df_final], ignore_index=True)  # smth like that
            with pd.ExcelWriter(path=file2, mode='wb') as writer_final:
                df_final.to_excel(writer_final, startrow=1, header=False, freeze_panes=(1, 0), index=False)
                self.df_final_format(writer_final, df_final)
        else:
            # writer_final = pd.ExcelWriter(file2, engine='xlsxwriter')
            with pd.ExcelWriter(path=file2, mode='wb') as writer_final:
                df_final.to_excel(writer_final, startrow=1, header=False, freeze_panes=(1, 0), index=False)
                self.df_final_format(writer_final, df_final)

    def run_parse(self):
        for file in os.scandir(directory):
            if file.name.endswith('.xml'):
                print(file.path)
                df, df_final, client_number = self.parse_reports(file.path)
                self.df_writer(df, df_final, client_number)


file1 = 'dataframe.xlsx'
# writer = pd.ExcelWriter(file1, engine='xlsxwriter')
if os.path.exists(r'D:\Python\dataframe_final.xlsx'):
    os.remove(path=r'D:\Python\dataframe_final.xlsx')
if os.path.exists(r'D:\Python\dataframe.xlsx'):
    os.remove(path=r'D:\Python\dataframe.xlsx')

cl = ParseDocs(file_path=file1)
cl.run_parse()

# for file in os.scandir(directory):
#     if file.name.endswith('.xml'):
#         print(file.path)
#         df, df_final, client_number = cl.parse_reports(file.path)
#         cl.df_writer(df, df_final, client_number, writer)
# 
# writer.close()
