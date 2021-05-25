# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import numpy as np
#import openpyxl

def email():
    content = MIMEMultipart()  #建立MIMEMultipart物件
    content["subject"] = "交期更改"  #郵件標題
    content["from"] = "hl94vul3h6@gmail.com"  #寄件者
    content["to"] = "gary6780gary6780@gmail.com" #收件者
    content.attach(MIMEText("您負責的客戶交期已被更改"))  #郵件內容

    with smtplib.SMTP(host="smtp.gmail.com", port="587") as smtp:  # 設定SMTP伺服器
        try:
            smtp.ehlo()  # 驗證SMTP伺服器
            smtp.starttls()  # 建立加密傳輸
            smtp.login("hl94vul3h6@gmail.com", "eeecnzvlhavhksyb")  # 登入寄件者gmail
            smtp.send_message(content)  # 寄送郵件
            print("Complete!")
        except Exception as e:
            print("Error message: ", e)

df = pd.read_excel("C:/Users/gary6/OneDrive/桌面/reschedule/05191.xlsx", usecols ="H:M") 
df1 = pd.read_excel("C:/Users/gary6/OneDrive/桌面/reschedule/05192.xlsx", usecols ="H:M")
new_df = df[1:2]
new_df1 = df1[1:2] 
#d2 = df.iat[2,7] #客人
print(new_df)
print(new_df1)
#if new_df != new_df1:
    #msg1 = '{}的訂單已被更改'
    #msg2 = msg1.format(d2)
    #print(msg2)

for id1 in new_df1.iteritems():
    
    flag = 0
    if(id1[1][1] is np.nan):
        continue
    for id2 in new_df.iteritems():
        if(id1[1][1] == id2[1][1]):
            flag = 1
            break
    if(flag == 0):
        #為舊檔案賦予新值再去比較
        id2[1][1] = id1[1][1]
        print(id2)
        print(id1)
if new_df.equals(new_df1):
    print('未更改')    
else:
    email()        
    print("您負責的客戶交期已被更改")        
            
        
        

           
            
            
            
            


    
#wb=openpyxl.load_workbook("C:/Users/gary6/OneDrive/桌面/reschedule/05191.xlsx")
#ws=wb.get_sheet_by_name('109')
#for row in ws.rows:
    #for cell in row:
        
        #print(cell.value,end='\t')
        #print()

#rows=tuple(ws.rows) #ws.rows會形成一個叫做generator的東東，稱之為生成器，它是所有列的集合，必須先轉成tuple元組或串列之後方可讀取。
#row=rows[2]  #選擇欲讀取的是第幾列
#for r in row:
    #依序取出每個cell的資料
    #print(r.value,end='\t')


#wb1=openpyxl.load_workbook("C:/Users/gary6/OneDrive/桌面/reschedule/05192.xlsx")
#ws1=wb1.get_sheet_by_name('109')
#for row in ws1.rows:
    #for cell in row:
        
        #print(cell.value,end='\t')
        #print()

#rows1=tuple(ws1.rows) #ws.rows會形成一個叫做generator的東東，稱之為生成器，它是所有列的集合，必須先轉成tuple元組或串列之後方可讀取。
#row2=rows1[2]  #選擇欲讀取的是第幾列
#for r in row2:
    #依序取出每個cell的資料
    #print(r.value,end='\t')

#for r in row:
    #for r1 in row2:
        #if r.value != r1.value:
            #print("交期已被更改")
    
