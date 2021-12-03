# -*- coding: utf-8 -*-
"""
Created on Wed Dec  1 10:57:00 2021

@author: anton
"""


import socket
import smtplib
from email.mime.multipart import MIMEMultipart


def connect_check(address):
    global s
    s = socket.socket()       # Create a socket object
    port = 445                # Reserve a port for your service.
    return s.connect_ex((address, port))


sender = 'AUTO@mail2.sumika.com.tw'  
recipients = ['anton@sumika.com.tw','kyle.hu@sumika.com.tw'] 
COMMASPACE = ','
ip = "172.21.0.249"

if connect_check(ip) == 0:
    print("有連線")
    s.close()
else:
   outer = MIMEMultipart()
   outer['Subject'] = 'RFID遠端主機未偵測到連線,請確認有無關機'
   outer['To'] = COMMASPACE.join(recipients)
   outer['From'] = sender
   composed = outer.as_string()  
   smtp = smtplib.SMTP()
   smtp.connect('172.16.16.17')
   smtp.sendmail(sender, recipients, composed)
   smtp.quit()
    


