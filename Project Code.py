# -*- coding: utf-8 -*-
"""
Created on Tue Nov 30 13:04:41 2021

@author: Admin
"""

#To read the googlesheet in python
import gspread
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from os.path import basename
import time

#To Access all the data from google spreadsheet
    
sa= gspread.service_account(filename = "mypyproject.json")
sh= sa.open("Datasheet")
wks = sh.worksheet("Sheet1")
data = wks.get_all_records()

#print(data)

#To filter the data according to supplier
#created loop here 

import pandas as pd
sheet_id = "1jlt4P483NulkUrvRSvtXB5vyadOWOJaCyXfrxnb5wBk"
sheet_name = "Sheet1"
url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"

df= pd.read_csv(url)

suppliers = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
dictionary = {
                'A' : 'suman.salunkhe555@gmail.com',
                'B' : 'pjsalunk@ncsu.edu',
                'C' :'rushikeshattarde@gmail.com',
                'D' :'salunkhe.pratiksha555@gmail.com',
                'E' :'foodyteddy10@gmail.com',
                'F' :'foodyteddy10@gmail.com',
                'G' :'pjsalunk@ncsu.edu',
                'H' :'suman.salunkhe555@gmail.com',
                'I' :'salunkhe.pratiksha555@gmail.com',
                'J' :'salunkhe.pratiksha555@gmail.com',
      
                }

for i in suppliers:
    
    newdf= df[df.Supplier==i]
    
    data = newdf.filter(['Product ID',"Name of the Product", "Qty to be ordered", "Retail price per unit", "Total"])
    
    print(data)
    
    #To save the filterd data as an excel
    data.to_excel(r"C:\Users\Admin\Desktop\Purchase Order" + i +".xlsx")
    time.sleep(2)
    
    #To send excel file as an attachment via Mail

    
    def send_mail(send_from: str, subject: str, text: str, 
    send_to: list, files= None):
    
        send_to= default_address if not send_to else send_to
    
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = ', '.join(send_to)  
        msg['Subject'] = subject
    
        msg.attach(MIMEText(text))
    
        for f in files or []:
            with open(f, "rb") as fil: 
                ext = f.split('.')[-1:]
                attachedfile = MIMEApplication(fil.read(), _subtype = ext)
                attachedfile.add_header(
                    'content-disposition', 'attachment', filename=basename(f) )
            msg.attach(attachedfile)
    
    
        smtp = smtplib.SMTP(host="smtp.gmail.com", port= 587) 
        smtp.starttls()
        smtp.login(username,password)
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.close()
          
    #its not recommended to directly use the mailID and Password in program  
    #instead generating environment variables in control panel is more secure
    #this is demo mailID we are going to use
    
    username = 'foodyteddy10@gmail.com'
    password = 'Foody@10'
    default_address = ['foodyteddy10@gmail.com'] 
    
    send_mail(send_from= username,
    subject="Purchase Order for supply of product",
    text="Hello,\nGreetings of the Day!\nWe are pleased to place our order on you for the items attached to this mail. \nThank you.",
    
    send_to= [dictionary[i]],
    files= [r"C:\Users\Admin\Desktop\Purchase Order" + i +".xlsx"])
    


#filtering data for pie chart
import pandas as pd
from matplotlib import pyplot as plt

sheet_id = "1jlt4P483NulkUrvRSvtXB5vyadOWOJaCyXfrxnb5wBk"
sheet_name = "Sheet1"
url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"

df= pd.read_csv(url)


M_avg = df['Exp1Wk'].mean()
N_avg = df['Exp2Wk'].mean()
O_avg = df['Exp3Wk'].mean()
P_avg = df['Remaining'].mean()

labels = ['Exp 1 week', 'Exp 2 weeks', 'Exp 3 weeks', 'Remaining Goods']

data = [M_avg, N_avg, O_avg, P_avg]
exp = [0.2, 0, 0, 0]
cl = ['red', 'green', 'blue', 'purple']

fig = plt.figure(figsize =(10, 7))
plt.pie(data, labels = labels, explode= exp, colors=cl)

