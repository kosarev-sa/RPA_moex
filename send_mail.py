"""
Программа отправки писем:
12. Направить итоговый файл отчета себе на почту;
13. В письме указать количество строк в Excel в правильном склонении.
"""
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from platform import python_version

from moex_to_xlsx import PATH
from moex_to_xlsx import Moex
from psw_conf import PASSWORD

server = 'smtp.yandex.ru'
user = 'mrsupersirius@yandex.ru'
password = PASSWORD

recipients = ['kosarev-sa@mail.ru', 'mrsupersirius@yandex.ru']
sender = 'mrsupersirius@yandex.ru'
subject = 'Тестовое задание - Гринатом'
text = f'Файл <b>{PATH}</b> содержит <h1>{Moex.num_rows(PATH)}</h1> c данными.<br><br>' \
       f'<i>Сергей Косарев</i>'
html = '<html><head></head><body><p>' + text + '</p></body></html>'

filepath = PATH
basename = os.path.basename(filepath)
filesize = os.path.getsize(filepath)

msg = MIMEMultipart('alternative')
msg['Subject'] = subject
msg['From'] = 'Python script for Greenatom<' + sender + '>'
msg['To'] = ', '.join(recipients)
msg['Reply-To'] = sender
msg['Return-Path'] = sender
msg['X-Mailer'] = 'Python/' + (python_version())

part_text = MIMEText(text, 'plain')
part_html = MIMEText(html, 'html')
part_file = MIMEBase('application', 'octet-stream; name="{}"'.format(basename))
part_file.set_payload(open(filepath, "rb").read())
part_file.add_header('Content-Description', basename)
part_file.add_header('Content-Disposition', 'attachment; filename="{}"; size={}'
                     .format(basename, filesize))
encoders.encode_base64(part_file)

msg.attach(part_text)
msg.attach(part_html)
msg.attach(part_file)

if __name__ == '__main__':
    mail = smtplib.SMTP_SSL(server)
    mail.login(user, password)
    mail.sendmail(sender, recipients, msg.as_string())
    mail.quit()
