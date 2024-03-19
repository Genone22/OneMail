import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def emails_send_funct(to_, subj_, msg_, from_, name):
    try:
        s = smtplib.SMTP("connect.smtp.bz", 587)
        s.command_encoding = "utf-8"
        s.starttls()
        s.login("pismo@otkaznoe-pismo.ru", "*MvA0S4v")

        msg = MIMEMultipart("related")
        msg["Subject"] = subj_
        msg["From"] = "{} <{}>".format(name, from_)
        msg["To"] = to_

        # Вкладываем обычный и HTML-версии тела сообщения в раздел 'alternative',
        # чтобы агенты сообщений могли решить, что отображать.
        msg_alternative = MIMEMultipart("alternative")
        msg.attach(msg_alternative)

        msg_text = MIMEText(msg_, "plain", "utf-8")
        msg_alternative.attach(msg_text)

        msg_html = MIMEText(msg_, "html", "utf-8")
        msg_alternative.attach(msg_html)

        # Отправка почты
        encoded_msg_string = msg.as_string().encode("utf-8")
        try:
            s.sendmail(from_, to_, encoded_msg_string)
        except UnicodeEncodeError:
            print(
                f"Невозможно отправить письмо на адрес: {to_} (содержит неподдерживаемые символы)"
            )
            return "f"

        x = s.ehlo()
        if x[0] == 250:
            return "s"
        else:
            return "f"

    except smtplib.SMTPRecipientsRefused as e:
        print("Не удалось отправить письмо: Неверный адрес получателя")
        print("Подробности ошибки:", e.recipients)

    except smtplib.SMTPException as e:
        print("Не удалось отправить письмо:", e)

    finally:
        s.quit()
