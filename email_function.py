# -*- coding: utf-8 -*-
import smtplib
from email.mime.text import MIMEText


def emails_send_funct(to_, subj_, msg_, from_, pass_):
    # Create session for mail.ru
    s = smtplib.SMTP('smtp.yandex.ru', 587)
    # Start TLS for security
    s.starttls()
    # Authentication
    s.login(from_, pass_)
    # Message to be sent
    msg = MIMEText(msg_, "plain", "utf-8")
    msg['Subject'] = subj_
    msg['From'] = from_
    msg['To'] = to_
    # Sending the mail
    s.sendmail(from_, to_, msg.as_string())
    x = s.ehlo()
    if x[0] == 250:
        return 's'
    else:
        return 'f'
    # Terminating the session
    s.quit()
