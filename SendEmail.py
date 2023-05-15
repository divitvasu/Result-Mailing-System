
#############################################
#############IMPORTING MODULES###############
#############################################

import smtplib
import os.path as op
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders
import xlrd
import easygui

import xlrd
try:
    b=xlrd.open_workbook('EmailList.xls')
except OSError as e:
    print(e)
    import SendEmail

sheet = b.sheet_by_name('Sheet1')

#############################################
###############SENDING EMAIL#################
#############################################

def send_mail(send_from="", send_to="", subject="Exam Result", message="Please Find your result Attached.", file1='',
              server="smtp.gmail.com", port=587, username="", password="",
              use_tls=True):
    """Compose and send email with provided info and attachments.

    Args:
        send_from (str): from name
        send_to (str): to name
        subject (str): message title
        message (str): message body
        files (list[str]): list of file paths to be attached to email
        server (str): mail server host name
        port (int): port number
        username (str): server auth username
        password (str): server auth password
        use_tls (bool): use TLS mode
    """
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(message))

    path=file1
    part = MIMEBase('application', "octet-stream")
    with open(path, 'rb') as file:
        part.set_payload(file.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition',
                        'attachment; filename="{}"'.format(op.basename(path)))
    msg.attach(part)
    
    smtp = smtplib.SMTP(server, port)
    if use_tls:
        smtp.starttls()
    smtp.login(username, password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    print("Mail Sent to "+send_to+" with "+file1+" as attachment ")
    smtp.quit()
    
"""to=input('To: ')
f=input('From: ')
password=input('Password: ')
subject=input('Subject: ')
Message=input('Message: ')
attachment=input('Filename: ')
"""

#############################################
##############SENDER'S EMAIL#################
#############################################

send_from=input("Email: ")
password=easygui.passwordbox("Password: ")

#############################################
############FETCHING EMAIL LIST##############
#############################################

for r in range(sheet.nrows-1):
    #print(sheet.cell(r+1,0).value)
    attachment=str(int(sheet.cell(r+1,0).value))+'.pdf'
    mailto=sheet.cell(r+1,1).value
    try:
        send_mail(send_from=send_from,send_to=mailto,file1=attachment,username=send_from,password=password)
    except Exception:
        print("Unable to send Mail to "+send_to)

b.release_resources()
del b

