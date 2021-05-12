#!/usr/bin/env python
# coding: utf-8

# In[2]:


#!pip install numpy
#!pip install pandas

import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
import ssl 
from datetime import datetime as dt
import zipfile
import os
import random
import string 
import numpy as np
import pandas as pd
import sqlite3 as slt
import jinja2


class NotificationEmail:
    """
        Class used to send notification by:
        - e-mail: complete
    """
    
    def __init__(self, subject:str='', message:str='', email:str='cob.corporativo@equatorialenergia.com.br', 
                 password:str='', path_file:str='', email_list:list=[], 
                 host:str='webmail.equatorialenergia.com.br', port:int=25):
        """
            Arguments:
            - subject: title of the e-mail (string format)
            - message: content of the message (string format)
            - email: e-mail of the sender (string format)
            - password: password of the e-mail from the sender (string format)
            - path_file: path of the file, if have a attachment (string format)
            - email_list: list of e-mails to send the e-mail (list of e-mails in string format)
            - host: address of the host from e-mail server (string format)
            - port: port of the host (int format)
        """
        self.subject = subject
        self.message = message
        self.email = email
        self.password = password
        self.path_file = path_file
        self.email_list = email_list
        self.host = host
        self.port = port 

    def sendEmail(self):
        """
            used to send the message by e-mail after the arguments have been supplied to the constructor when instantiating the class 
        """
        address_book = self.email_list
        msg = MIMEMultipart()
        sender = self.email
        msg['From'] = sender
        msg['To'] = ','.join(address_book)
        msg['Subject'] = self.subject
        msg.attach(MIMEText(self.message, 'html'))
        if(self.path_file != ''):
            part = self.__prepareAttachment()
            if part:
                msg.attach(part)
        text=msg.as_string()
        #context = ssl.create_default_context()
        with smtplib.SMTP(host=self.host , port=self.port) as server:
            print('sending...')
            server.ehlo()
            server.starttls()
            server.login(self.email,self.password)
            server.sendmail(sender,address_book, text)
            print('sended!!!')

    def __compressFiles(self)-> str:
        if (self.path_file != ''):
            print(f"Compressing file {self.path_file} ...")
            name = self.__randomString()
            if os.path.isdir(self.path_file):
                output_file = os.path.join(self.path_file , f'{name}.zip')
                zip_file = zipfile.ZipFile(output_file, 'w')
                for dirpath, _ , fnames in os.walk(self.path_file):
                    for f in fnames:
                        if f.endswith(".db") or f.endswith(".txt") or f.endswith(".htm") or f.endswith(".xlsx"):
                            fl = os.path.join(dirpath, f)
                            zip_file.write(fl, arcname=f , compress_type=zipfile.ZIP_DEFLATED)
                zip_file.close()
            elif os.path.isfile(self.path_file):
                name = self.path_file.split(os.sep)[-1]
                opath = os.path.abspath(os.path.join(self.path_file , '..'))
                output_file = os.path.join(opath , f'{name}.zip')
                zip_file = zipfile.ZipFile(output_file, 'w')
                zip_file.write(self.path_file, arcname=name , compress_type=zipfile.ZIP_DEFLATED)
                zip_file.close()
            print(f"File {self.path_file} compressed!!!")
        else:
            print('path is empty!')
        return output_file

    def __prepareAttachment(self):    
        zipedfile = self.__compressFiles()
        print(f"attaching file {self.path_file} ...")
        part = MIMEBase('application', "zip")
        part.set_payload( open(zipedfile,'rb').read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="arquivos.zip"')
        print(f"File {self.path_file} attached!!!")
        return part

    def __prepareContent(self, filename:str , dateref:dt.date , content:list , company:str):
        templateLoader = jinja2.FileSystemLoader(searchpath=self.path_file)
        templateEnv = jinja2.Environment(loader=templateLoader)
        templ = templateEnv.get_template(filename)

        html = templ.render(
                                company=company ,
                                data=dateref.strftime('%d/%m/%Y'),
                                content = content
                           )
        return html

    def __randomString(self, stringLength=10):
        letters = string.ascii_lowercase
        return ''.join(random.choice(letters) for i in range(stringLength))

# In[ ]:




