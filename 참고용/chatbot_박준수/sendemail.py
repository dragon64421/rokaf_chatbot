import smtplib
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
import os

def send_email(destin, subject, body, file):
    sender = 'afakakaobot@gmail.com'
    email_password = 'Qq28095774!'
    filename = file
    
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = destin
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    
    try:
        filename = filename+'.xlsx'
        attachment = open(filename, 'rb')
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment", filename=os.path.basename(filename))
        msg.attach(part)
    except:
        print('no file')
    
    text = msg.as_string()
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(sender, email_password)
    server.sendmail(sender, destin, text)
    server.quit()