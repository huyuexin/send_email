#coding:utf-8   #强制使用utf-8编码格式
# -*- coding:utf-8 -*-
import smtplib
#加载smtplib模块
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.utils import formataddr

import xlrd

#发件人邮箱账号q
my_sender=input('input your qq email:')
my_password=input('input your password:')
smtp_var1=my_sender.split('@')
my_smtp='smtp.'+smtp_var1[1]
#收件人称呼、邮箱地址、邮箱内容、附件名字
file_name = 'mail.xls'

def readFile(filename):
    rowdates = []
    workbook = xlrd.open_workbook(filename)
    sheet2_name = workbook.sheet_names()
    #从第一个表开始 或指定名字
    sheet2 = workbook.sheet_by_index(0)
    #sheet2 = workbook.sheet_by_name('sheet 1')
    for row_i in range(sheet2.nrows):
        rowdate= sheet2.row_values(row_i)
        rowdates.append(rowdate)
    return rowdates



def mail(my_sender,my_password,my_smtp,maildate_name,maildate_email,maildate_contents,maildate_append,maildate_appendname):
    ret=True
    try:
        msg = MIMEMultipart()
        msg["Subject"] = "党员信息"
        msg["From"] = formataddr(["揭阳市揭东第一中学",my_sender])
        msg["To"] = formataddr([maildate_name,maildate_email])

        # ---这是文字部分---
        part = MIMEText(maildate_contents,'plain','utf-8')
        msg.attach(part)


        # ---这是附件部分---
        # xls类型附件
        part = MIMEApplication(open(maildate_append, 'rb').read())
        part.add_header('Content-Disposition', 'attachment', filename=maildate_appendname)
        msg.attach(part)


        server=smtplib.SMTP(my_smtp,25)
        #发件人邮箱中的SMTP服务器，端口是25

        server.login(my_sender,my_password)
        #括号中对应的是发件人邮箱账号、邮箱密码

        server.sendmail(my_sender,[maildate_email,],msg.as_string())
        #括号中对应的是发件人邮箱账号、收件人邮箱账号、发送邮件

        server.quit()
        #这句是关闭连接的意思
        print('发送成功')

    except Exception:
        #如果try中的语句没有执行，则会执行下面的ret=False
        print('发送失败')


maildate = readFile(file_name)

for i in range(1,len(maildate)):
    name = maildate[i][1]
    email = maildate[i][2]
    contents =maildate[i][3]
    appends =maildate[i][4]

    appendname =maildate[i][5]
    print(appendname)
    mail(my_sender,my_password,my_smtp,name,email,contents,appends,appendname)
#ret=mail()
#if ret:
#    print("ok")
#    #如果发送成功则会返回ok，稍等20秒左右就可以收到邮件
#else:
#    print("filed")
#    #如果发送失败则会返回filed
