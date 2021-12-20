
import smtplib
import os
import pandas as pd
from datetime import date
import datetime
import numpy as np
from datetime import timedelta
import ssl
#import xlwings as xw
import time
import threading
import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import sys
from datetime import datetime
threads = []

EMAIL_ADDRESS = "mailsenderforschools@gmail.com"
EMAIL_PASS = "kndtmsjatpmqrxzw"



print("Hello, welcome to safe guarding Mail auto-sender.")


def import_spreadsheet(fn, ln, d):
    dt = pd.read_excel('/home/pi/Downloads/mailsender_safeGuarding/Safeguarding.xlsx')

    # Print all values of the Product column
    # .values return only the data in a list without data type
    Fnames_list = dt['Fname']
    Sname_list = dt['Sname']
    date_list = dt['Date_Due']
    if fn == 1 and d == 1 and ln :
        return dt
    elif d == 1:

        return date_list
    elif fn == 1:
        return Fnames_list
    elif ln == 1:
        return Sname_list

    else:
        print("No value printed")


# def start(date_list, name_list):
#    data = date_list.values.tolist()
#    while True:
#        for x, email_date in enumerate(data):
#            if data[x] == date:
#                names = name_list.values.tolist()
#                print(names[x]+"'s", "email has been sent")
#
#            else:
#                continue
def start():

    num = 0
    
    while True:
        
        for x, data in enumerate(import_spreadsheet(1, 1, 1).values):
            t = threading.Thread(target=targ, args=[x, data[0], data[1], data[2]])
            t.start()
            threads.append(t)
        for thread in threads:
            thread.join()



def targ(id, Sname, Lname, date_due):
    

    #dtObj = datetime.strptime(str(date_due), "%d.%m.%Y")
    #dtObj = dtObj.strftime("%d.%m.%Y")
    dtObj = datetime.strptime(str(date_due))
    email_date = dtObj - pd.DateOffset(months=1)
    if email_date.strftime("%d.%m.%Y") == date.today().strftime("%d.%m.%Y"):
    
        print(Sname+"'s", "email has been sent")
        context = ssl.create_default_context()
        try:
            with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                print(1)
                subject = "Safe Guarding refresher course for: {Sname} {Lname}".format(Sname=Sname, Lname=Lname)

                text = """
                
                Dear {Sname} {Lname}
                
                Dear 

Your level 1 safeguarding is due to expire in 4 weeks, please log on to Educare to complete the training 

You should have an Educare account but if you require password reminders please email the below link -
http://www.myeducare.com/login/forgot_password.php
 
The course that needs to be completed is Child Protection in Education (secondary) & contains 5 units.
 
Please liaise with your individual Line Managers if this is necessary regarding the completion of this.

-- 
Kind Regards 

Linda Iveson
Child protection & pastoral manager

Newland school for girls
Cottingham road
Hull
Hu6 7Ru

Tel; 01482 343098 ext 213
Email; ivesonl@thrivetrust.uk
W: https://www.newland.hull.sch.uk

                
                
                """.format(Sname=Sname, Lname=Lname)
                msg = 'Subject: {}\n\n{}'.format(subject, text)

                smtp.ehlo()  # Can be omitted
                smtp.starttls(context=context)  # Secure the connection
                smtp.ehlo()
                smtp.login(EMAIL_ADDRESS, EMAIL_PASS)
                smtp.sendmail(EMAIL_ADDRESS, 'IvesonL@thrivetrust.uk', msg)
                smtp.quit()

                while True:
                    try:
                        wb = openpyxl.load_workbook('/home/pi/Downloads/mailsender_safeGuarding/Safeguarding.xlsx')
                        sheet = wb.active


                        date_next_due = dtObj + pd.DateOffset(months=24)


                        i = id+2
                        sheet["C"+str(i)] = date_next_due.strftime("%d.%m.%Y")
                        wb.save("/home/pi/Downloads/mailsender_safeGuarding/Safeguarding.xlsx")
                        print("Excel Updated")
                        break
                        
                        
                        
                    except PermissionError as e:
                        print("Waiting for excel to close")
                        time.sleep(5)
        except smtplib.SMTPServerDisconnected as e:
            print(e)
    
        
        
        
        
        










start()