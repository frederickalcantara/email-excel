import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd

def sendemail(customer, subject, email):
    port = smtp_port_to_use
    password = password_to_use

    context = ssl.create_default_context()

    sender_email = email_address_to_send_it_under
    receiver_email = customer
    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = sender_email
    message["To"] = receiver_email

    text_parse = MIMEText(email, "plain")

    message.attach(text_parse)

    with smtplib.SMTP_SSL(domain_to_put_for_smtp, port, context=context) as server:
        server.login(email_to_send_under, password)
        server.sendmail(
            sender_email, 
            receiver_email, 
            message.as_string()
        )

df = pd.read_excel('file_path_to_excel_doc')

print(df)

email = df['Email']

subject = df['Subject']

message = df['Email Message']

if email.isnull().values.any():
    print("There is an Empty Value in the Email Column")
elif subject.isnull().values.any():
    print("There is an Empty Value in the Subject Message Column")
elif message.isnull().values.any():
    print("There is an Empty Value in the Email Message Column")
else: 
    for i in df.index:
        email = df['Email'][i]
        subject = df['Subject'][i]
        message = df['Email Message'][i]
        sendemail(email, subject, message)
