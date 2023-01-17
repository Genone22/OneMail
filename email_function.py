# -*- coding: utf-8 -*-
import smtplib
# from email.mime.text import MIMEText


def emails_send_funct(to_, subj_, msg_, from_, pass_):
    # print(to_, subj_, msg_, from_, pass_)

    # Create session for mail.ru
    s = smtplib.SMTP('smtp.yandex.ru', 587)  # 587
    # Start TLS for security
    s.starttls()
    # Authentication
    s.login(from_, pass_)
    # Message to be sent
    msg = MIMEText(msg_)
    msg = 'Subject: {}\n\n{}'.format(subj_, msg_)  # , encoding='utf-8'
    # Sending the mail
    s.sendmail(from_, to_, msg_)  # , msg_.as_string()
    x = s.ehlo()
    if x[0] == 250:
        return 's'
    else:
        return 'f'
    # Terminating the session
    s.quit()
