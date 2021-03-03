#!/usr/bin/python3
# -*- coding: utf-8 -*-

import xlwings as xw
import numpy as np
import os
import pandas as pd
import datetime as dt
import time
import win32com.client as win32
import pythoncom
from dateutil import rrule
import datetime
import easygui as eg

def get_file_name(dirpath=''):
    file_name = []
    list1 = os.listdir(dirpath)
    for i in range(0, len(list1)):
        #if os.path.isfile(list1[i]):
        file_name.append(os.path.join(dirpath, list1[i]))
    return file_name

def find_invoice_e9p(dirpath):
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    filepath = os.path.join(dirpath, 'zf31.xls')
    wb = app.books.open(filepath)
    E9P_finding_invoice = wb.sheets[0].range('D5:D2000').value
    E9P_finding_date = wb.sheets[0].range('A5:A2000').value
    E9P_finding_method = wb.sheets[0].range('B5:B2000').value
    df = pd.DataFrame({"Invoice": E9P_finding_invoice,
                       'Date': E9P_finding_date,
                       'Method': E9P_finding_method})
    df = df.dropna()
    wb.save()
    wb.close()
    time.sleep(3)
    return df

def find_invoice_inbound(filepath):
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    wb = app.books.open(filepath)
    inbound_finding = list(filter(None,wb.sheets[0].range('B2:B2000').value))+list(filter(\
        None,wb.sheets[1].range('A2:A2000').value))
    wb.save()
    wb.close()
    time.sleep(3)
    return inbound_finding


def workdays(start, end, holidays=0, days_off=None):
    if days_off is None:
        days_off = 5, 6
    workdays = [x for x in range(7) if x not in days_off]
    days = rrule.rrule(rrule.DAILY, dtstart=start, until=end, byweekday=workdays)
    return days.count() - holidays

def send_mail(invoice=[], reciList=[],subject=[],name=[],**kwargs):
    outlook = win32.Dispatch('Outlook.Application')
    for i in range(len(reciList)):
        mail_item = outlook.CreateItem(0)  # 0: olMailItem
        mail_item.Recipients.Add(reciList[i])
        mail_item.Subject = subject
        mail_item.BodyFormat = 2  # 2: Html format
        mail_item.HTMLBody = '''
                            Dear {0},\n\n

                            Can you send me the pre-alert corresponding to following invoices. Thanks.\n\n

                            {1}

                            '''.format(name,invoice)

        # mail_item.Attachments.Add('d:\doc\IT Info.xlsx')
        mail_item.Send()

if __name__ == '__main__':
    pythoncom.CoInitialize()
    sea_mail = []
    air_mail = []
    sea_mail.append(eg.enterbox(msg='请输入海运预报预警的邮箱（多个邮箱请用空格分开）： ', title="Pre-alert tracking V1.0\t   作者：\
    Henry Xue ",default=['heng.xue@volkswagen.com.cn'])) # yasin.ask@volkswagen.de yvonne.turovsky@volkswagen.de
    air_mail.append(eg.enterbox(msg='请输入空运预报预警的邮箱（多个邮箱请用空格分开）： ', title="Pre-alert tracking V1.0\t   作者：\
    Henry Xue ",default=['heng.xue@volkswagen.com.cn']))  # Kay.arrano-gonzalez@dhl.com','Chih-yun.wang@dhl.com
    dirpath = eg.diropenbox(msg='请指定发票文件夹的位置： ', title="Pre-alert tracking V1.0\t   作者：Henry Xue ")
    path_filename = get_file_name(dirpath)
    inbound_result = []
    today = time.time()
    sea_invoice = []
    air_invoice = []
    sea_limit = 7
    air_limit = 3
    sea_subject = 'Sea Per-alert notice'
    air_subject = 'Air Per-alert notice'
    sea_name = 'All'
    air_name = 'All'
    inbound_result_int = []
    for j in path_filename:
        if 'zf31' in j:
            path_filename.remove(j)
    for i in path_filename:
        inbound_result.extend(find_invoice_inbound(i))      #把inbound表中的发票号导出到列表中
    E9P_df = find_invoice_e9p(dirpath)      #把E9P里的发票数据导出
    result_df = pd.DataFrame(columns=['Invoice', 'Date', 'Method'])
    t = 0
    for i, j, k in zip(E9P_df['Invoice'], E9P_df['Date'], E9P_df['Method']): #比对E9P与inbound之间的差异，把他存入result_df
        if i not in inbound_result:
            result_df.loc[t] = [i, j, k]
            t = t + 1
    Sea_df = (result_df[result_df['Method'] == 'ZFVA'])     #把海运和空运分开建DF
    Air_df = (result_df[result_df['Method'] == 'ZFSO'])
    for i, j, k in zip(Sea_df['Invoice'], Sea_df['Date'], Sea_df['Method']):        #海运超过时限，则存到一个列表中
        Sea_date = dt.datetime.strptime(j, '%d.%m.%Y').date()
        i = int(i)
        gap = workdays(Sea_date, dt.datetime.today())
        if gap > sea_limit:
            sea_invoice.append(i)

    for i, j, k in zip(Air_df['Invoice'], Air_df['Date'], Air_df['Method']):        #空运超过时限，则存到一个列表中
        Air_date = dt.datetime.strptime(j, '%d.%m.%Y').date()
        i = int(i)
        gap = workdays(Air_date, dt.datetime.today())
        if gap > air_limit:
            air_invoice.append(i)
    send_mail(sea_invoice,sea_mail,sea_subject,sea_name)        #发海运邮件
    send_mail(air_invoice,air_mail,air_subject,air_name)        #发空运邮件
