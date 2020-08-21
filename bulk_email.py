import random
import smtplib
import ssl
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import xlrd


def configure(sender_name, sender_email, sender_password):
    port = 465
    sender_email = sender_email
    password = sender_password
    context = ssl.create_default_context()
    return sender_email, sender_name, password, context, port


def sender_details():
    loc = ("smtp.xlsx")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(1, 1)

    sender_name = []
    sender_email = []
    sender_password = []
    for i in range(sheet.nrows):
        if i == 0:
            continue
        sender_name.append(sheet.cell_value(i, 0))
        sender_email.append(sheet.cell_value(i, 1))
        sender_password.append(sheet.cell_value(i, 2))

    return sender_name, sender_email, sender_password


def read_files():
    name_list = []
    email_list = []
    with open('body.txt', 'r') as body:
        body_list = [line.rstrip() for line in body]

    loc = ("email.xlsx")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)

    for i in range(sheet.nrows):
        name_list.append(sheet.cell_value(i, 0))
        email_list.append(sheet.cell_value(i, 1))

    return name_list, email_list, body_list


def send_email(sender_email, sender_name, password, context, port, email_list, name_list, body_list):
    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        try:
            server.login(sender_email, password)

            print(f'Email will be sent shortly by {sender_email}')

            for receiver_email, receiver_name in zip(email_list, name_list):
                secure_random = random.SystemRandom()
                html = secure_random.choice(body_list)

                html = f"""\
                    <html>
                     <body>
                       <p>Hello {receiver_name}</p><br>
                       {html}
                     </body>
                    </html>
                """

                message = MIMEMultipart("alternative")
                message["Subject"] = "Re: Facebook"
                message["From"] = f"{sender_name} <{sender_email}>"
                message["To"] = receiver_email

                part1 = MIMEText(html, "html")

                filename = "attachment.jpg"

                with open(filename, "rb") as attachment:
                    part2 = MIMEBase("application", "octet-stream")
                    part2.set_payload(attachment.read())

                encoders.encode_base64(part2)

                part2.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {filename}",
                )

                message.attach(part1)
                message.attach(part2)

                print(f"Mail Send To {receiver_email}")
                server.sendmail(sender_email, receiver_email,
                                message.as_string())
        except Exception as e:
            print(f"We Faced Some Problem With This Sender {sender_email}.Don't Worry We Will Try To Send These Email Using Another Sender")
            return False


def split(arr, count):
    return [arr[i::count] for i in range(count)]


if __name__ == '__main__':
    sender_name, sender_email, sender_password = sender_details()
    name_list, email_list, body_list = read_files()

    limit = len(sender_email)
    email_list_chunk = split(email_list, limit)
    name_list_chunk = split(name_list, limit)

    failed_email_list = []
    failed_name_list = []
    success_sender_index = []
    for i in range(limit):
        from_email, from_name, from_password, context, port = configure(
            sender_name[i], sender_email[i], sender_password[i])

        result = send_email(from_email, from_name, from_password, context,
                            port, email_list_chunk[i], name_list_chunk[i], body_list)
        if result == False:
            failed_email_list.extend(email_list_chunk[i])
            failed_name_list.extend(name_list_chunk[i])
        else:
            success_sender_index.append(i)
            continue

    if len(failed_email_list) == len(email_list):
        print(
            "Your All Sender Email Is Not Correct.Please Fix The Problem And Try It Again")
        sys.exit()
    else:
        limit = len(success_sender_index)
        email_list_chunk = split(failed_email_list, limit)
        name_list_chunk = split(failed_name_list, limit)
        for i in success_sender_index:
            from_email, from_name, from_password, context, port = configure(
                sender_name[i], sender_email[i], sender_password[i])

            result = send_email(from_email, from_name, from_password, context,
                                port, email_list_chunk[i], name_list_chunk[i], body_list)

    print('Complete')
