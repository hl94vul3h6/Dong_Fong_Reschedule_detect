#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText #專門傳送正文
import smtplib
import sys
import datetime
import os
from pandas.core.frame import DataFrame


# In[2]:


#讀excel
def load_data(file_name):
    df_data = pd.read_excel(file_name)
    return df_data


# In[3]:


#資料處理跟抓取要放入郵件內容
def data_process(df, df1):
    
    id2_count = 0
    mail_list = []
    for id2 in df.iteritems(): 
        id2_count += 1
        id1_count = 0
        for id1 in df1.iteritems():
            id1_count += 1 
            if(id2[1][1] == id1[1][1]) and (id2_count != id1_count):
                print("id1:",id1[1][1],"id2:",id2[1][1],"id1_count:",id1_count,"id2_count:",id2_count)
                mail_list.append(id1[1][1])
                
 

        
    if df.equals(df1):
        return None
    
               
    c = {"mail_list" : mail_list} #list轉df
    mail_list_df = DataFrame(c)
               
    #dataframe轉html
    df_html = mail_list_df.to_html(escape=False,index=False, justify = 'center')
    
    # 回傳final的html格式
    return df_html


# In[4]:


def read_file(file):
    with open(file,'r',encoding="utf-8") as f:
            content_list = f.read().split('\n')
    return content_list


# In[5]:


def SendMail(msg_html):
    #print('start to send mail...')
    emails = read_file('email.txt') #讀email資料
    sender = emails[0] #step1:setup sender gmail,ex:"Fene1977@superrito.com"
    password = emails[1] #step2:setup sender gmail password
    recipients = emails[2] #step3:setup recipients mail
    today_date = datetime.date.today() 
    sub = today_date.strftime("%m/%d") + "訂單交期更改通知" #step4:setup your subject
    
    outer = MIMEMultipart()
    outer['From'] = sender #step:setup sender gmail
    outer['To'] = recipients #step:setup recipient mail
    #outer["Cc"] = cc_mail #step:setup cc mail
    outer['Subject'] = sub #step:setup your subject

    #設定純文字資訊
    plainText = "偵測到訂單排定交貨日更改，請確認以下訂單："
    msgText = MIMEText(plainText, 'plain', 'utf-8')
    outer.attach(msgText)
    
    #設定HTML資訊
    htmlText = msg_html #step7:edit your mail content
    msgText = MIMEText(htmlText, 'html', 'utf-8')
    outer.attach(msgText)

    mailBody = outer.as_string()
    #-----------------------------------------------------------------------
    # 寄送EMAIL
    try:
        with smtplib.SMTP(host="smtp.gmail.com", port="587") as s: #send webservice to gmail smtp socket
            s.ehlo()  # 驗證SMTP伺服器
            s.starttls()  # 建立加密傳輸
            s.login(sender, password)  # 登入寄件者gmail
            s.sendmail(sender, recipients,mailBody)  # 寄送郵件
            s.close()
        print("Email sent!")
    except:
        print("Unable to send the email. Error: ", sys.exc_info()[0])
        raise


# In[6]:


def removeoldfile(olderfile):
    olderfile = r"05191.xlsx"

    try:
        os.remove(olderfile)
    except OSError as e:
        print(e)
    else:
        print("File is deleted successfully")


# In[7]:


import shutil
def file_copy_change_name(todayfile):
    #src = "生產進度110.1-12.xlsx"
    #dst = "生產進度110.1-12_old.xlsx"
    src = "05192.xlsx"
    dst = "05191.xlsx"
    shutil.copyfile(src, dst)
    return dst


# In[8]:


def main():
    #yesterday = load_data('生產進度110.1-12_old.xlsx', sheet_name = '109.01-')
    #today = load_data('生產進度110.1-12.xlsx', sheet_name = '109.01-')
    yesterday = load_data('05191.xlsx')
    today = load_data('05192.xlsx')
    mail_content = data_process(yesterday, today)

    if mail_content == None:
        print("交期未更改")
    else:
        SendMail(mail_content)
        print("交期已被更改")
        removeoldfile(yesterday)
        file_copy_change_name(today)
main()

