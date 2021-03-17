#!/usr/bin/python3
# -*- coding: utf-8 -*-

import xlwings as xw
import os
import pandas as pd
import datetime as dt
import time
import win32com.client as win32
import pythoncom
from dateutil import rrule
import easygui as eg

def get_file_name(dirpath=''):
    file_name = []
    list1 = os.listdir(dirpath)
    for i in range(len(list1)):
        if os.path.isfile(os.path.join(dirpath,list1[i])):
            if '~' not in list1[i]:
                file_name.append(os.path.join(dirpath, list1[i]))
    return file_name

def find_invoice_e9p(dirpath):
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    filepath = os.path.join(dirpath, 'zf31.xls')
    wb = app.books.open(filepath)
    E9P_finding_invoice = wb.sheets['zf31'].range('D5:D2000').value
    E9P_finding_date = wb.sheets['zf31'].range('A5:A2000').value
    E9P_finding_method = wb.sheets['zf31'].range('B5:B2000').value
    df = pd.DataFrame({"Invoice": E9P_finding_invoice,
                       'Date': E9P_finding_date,
                       'Method': E9P_finding_method})
    df = df.dropna()
    wb.save()
    wb.close()
    return df

def find_invoice_inbound(filepath):
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    wb = app.books.open(filepath)
    inbound_finding = list(filter(None,wb.sheets['IN PROCESS'].range('B2:B2000').value))+list(filter(\
        None,wb.sheets['VW SP'].range('A2:A2000').value))
    wb.save()
    wb.close()
    return inbound_finding

def send_mail(invoice=[], addressee=[],cc=[],subject=[],name=[]):
    olook = win32.Dispatch("outlook.Application")  # 固定写法
    mail = olook.CreateItem(0)  # 固定写法
    mail.To = addressee  # 收件人
    mail.CC = cc  # 抄送人
    mail.Subject = subject
    #mail_item.BodyFormat = 2  # 2: Html format
    mail.Body = '''Dear {0},\n\nCan you send me the pre-alert corresponding to following invoices. Thanks.\n\n{1}'''\
        .format(name,invoice)
    mail.Send()  # 发送

if __name__ == '__main__':
    pythoncom.CoInitialize()
    sea_mail_re = []
    sea_mail_cc = []
    air_mail_re = []
    air_mail_cc = []
    Collect_fields = ['发送邮箱','抄送邮箱']
    temp_mail_sea = eg.multenterbox(msg='请输入海运预警的邮箱（多个邮箱请用分号隔开）：', title="Pre-alert tracking V1.0 beta 作者：\
    Henry Xue ",fields= Collect_fields,values=['heng.xue@volkswagen.com.cn','heng.xue@volkswagen.com.cn']) # yasin.ask@volkswagen.de yvonne.turovsky@volkswagen.de
    sea_mail_re = temp_mail_sea[0]
    sea_mail_cc = temp_mail_sea[1]
    temp_mail_air = eg.multenterbox(msg='请输入空运预警的邮箱（多个邮箱请用分号隔开）：', title="Pre-alert tracking V1.0 beta 作者：\
    Henry Xue ",fields= Collect_fields,values=['heng.xue@volkswagen.com.cn','heng.xue@volkswagen.com.cn'])  # Kay.arrano-gonzalez@dhl.com','Chih-yun.wang@dhl.com
    air_mail_re = temp_mail_air[0]
    air_mail_cc = temp_mail_air[1]
    dirpath = eg.diropenbox(msg='请指定发票文件夹的位置： ', title="Pre-alert tracking V1.0\t   作者：Henry Xue ")
    sea_limit = int(eg.enterbox(msg ='请输入海运预警天数（自然日）：',default=7))
    air_limit = int(eg.enterbox(msg='请输入空运预警天数（自然日）：',default=5))
    path_filename = get_file_name(dirpath)
    inbound_result = []
    today = time.time()
    sea_invoice = []
    air_invoice = []
    sea_subject = 'Sea Per-alert notice  ' + str(dt.datetime.now())[0:10]
    air_subject = 'Air Per-alert notice  ' + str(dt.datetime.now())[0:10]
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
        Sea_gap = rrule.rrule(rrule.DAILY, dtstart = Sea_date, until = dt.datetime.today()).count()     #计算自然日差
        if Sea_gap > sea_limit:
            sea_invoice.append(i)

    for i, j, k in zip(Air_df['Invoice'], Air_df['Date'], Air_df['Method']):        #空运超过时限，则存到一个列表中
        Air_date = dt.datetime.strptime(j, '%d.%m.%Y').date()
        i = int(i)
        Air_gap = rrule.rrule(rrule.DAILY, dtstart = Air_date, until = dt.datetime.today()).count()  #计算自然日差
        if Air_gap > air_limit:
            air_invoice.append(i)

    if eg.ynbox(msg='''海运有如下发票未能及时收到预报，是否需要发送邮件？\n\n{}'''.format(sea_invoice)):
        send_mail(sea_invoice,sea_mail_re,sea_mail_cc,sea_subject,sea_name)        #发海运邮件

    if eg.ynbox(msg='''空运有如下发票未能及时收到预报，是否需要发送邮件？\n\n{}'''.format(air_invoice)):
        send_mail(air_invoice,air_mail_re,air_mail_cc,air_subject,air_name)        #发空运邮件