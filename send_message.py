import settings
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders


def say_string(num):
    string = 'строк'
    if num == 21:
        return f'{string}а'
    elif num in [22, 23, 24]:
        return f'{string}и'
    return string

def send_mail(subject, text, mail_to, files, isTls=True):
    msg = MIMEMultipart()
    msg['From'] = settings.mail_from
    msg['To'] = settings.mail_to
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = subject
    msg.attach(MIMEText(text))

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(files, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="WorkBook3.xlsx"')
    msg.attach(part)

    smtp = smtplib.SMTP(settings.server, settings.port)
    if isTls:
        smtp.starttls()
    smtp.login(settings.mail_from, settings.password)
    smtp.sendmail(settings.mail_from, mail_to, msg.as_string())
    smtp.quit()
