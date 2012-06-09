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
    robo_password,
)

def sendemail_attach(
        fromaddr=robo_user,
        password=robo_password,
        message=None,
        subject=None,
        to=None,
        cc=[],
        attachment=False,
        attachmentname=None,
        html=None,
    ):
    """Function to send an email with an attachment.

    kwargs:

    fromaddr -- sender email address
    password -- sender password
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
    if attachment:
        ctype,encoding = mimetypes.guess_type(attachment)
        if ctype is None or encoding is not None:
            ctype = 'application/octet-stream'
        maintype, subtype = ctype.split('/', 1)
        if maintype == 'text':
            with open(attachment, 'r') as attachment_file:
                msg = MIMEText(attachment_file.read(), _subtype=subtype)
        elif maintype == 'image':
            with open(attachment, 'r') as attachment_file:
                msg = MIMEImage(attachment_file.read(), _subtype=subtype)
        else:
            msg = MIMEBase(maintype, subtype)
            with open(attachment, 'r') as attachment_file:
                msg.set_payload(attachment_file.read())
            encoders.encode_base64(msg)
        msg.add_header('Content-Disposition',
            'attachment', filename=attachmentname)
        emailMsg.attach(msg)
    composed = emailMsg.as_string()
    server = smtplib.SMTP('smtp.gmail.com:587')
    server.starttls()
    server.login(fromaddr, password)
    server.sendmail(fromaddr, to, composed)
    server.quit()

def send_email(*args, **kwargs):
    """Copy of sendemail_attach.

    sendemail_attach should just be named send_email, but a lot of older
    scripts refrence the old name.

    """
    sendemail_attach(*args, **kwargs)
