#### All functions relating to Emails ####

import mimetypes
import os
import smtplib
import string
import sys
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from robo_configs import (
    robo_user,
    robo_passwd,
)

def sendemail_attach(
        fromaddr=robo_user,
        pswd=robo_passwd,
        message=None,
        subject=None,
        to=None,
        cc=[],
        attachment=0,
        attachmentname=0,
        html=0
    ):
    """Function to send an email with an attachment.

    kwargs:

    fromaddr -- sender email address
    pswd -- sender password
    message -- text message contents
    subject -- email subject
    to -- list of emails to send email to ie. to=['you@me.com', 'her@me.com',
        'him@me.com']
    cc -- works the same as `to`
    attachment -- path to the attachment
    attachmentname -- name you want to show up for the attachment in the email
    html -- html message, defaults to false

    **written with A LOT of help from google**

    """
    emailMsg = MIMEMultipart('alternative')
    emailMsg['Subject'] = subject
    emailMsg['From'] = fromaddr
    emailMsg['To'] = ', '.join(to)
    emailMsg.attach(MIMEText(message, 'plain', _charset='UTF-8'))
    if html:
        emailMsg.attach(MIMEText(html, 'html'))
    if attachment != 0:
        ctype,encoding = mimetypes.guess_type(attachment)
        if ctype is None or encoding is not None:
            ctype = 'application/octet-stream'
        maintype, subtype = ctype.split('/', 1)
        if maintype == 'text':
            fp = open(attachment)
            msg = MIMEText(fp.read(), _subtype=subtype)
            fp.close()
        elif maintype == 'image':
            fp = open(attachment, 'rb')
            msg = MIMEImage(fp.read(), _subtype=subtype)
            fp.close()
        else:
            fp = open(attachment,'rb')
            msg = MIMEBase(maintype, subtype)
            msg.set_payload(fp.read())
            fp.close()
            encoders.encode_base64(msg)
        msg.add_header('Content-Disposition',
            'attachment', filename=attachmentname)
        emailMsg.attach(msg)
    composed = emailMsg.as_string()
    username = fromaddr
    password = pswd
    server = smtplib.SMTP('smtp.gmail.com:587')
    server.starttls()
    server.login(username, password)
    server.sendmail(fromaddr, to, composed)
    server.quit()

def send_email(*args, **kwargs):
    """Copy of sendemail_attach.

    sendemail_attach should just be named send_email, but a lot of older
    scripts refrence the old name.

    """
    sendemail_attach(*args, **kwargs)
