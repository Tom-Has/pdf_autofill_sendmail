import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import time

# SMTP-Server Einstellungen
smtp_server = 'smtp server'
smtp_port = 587
smtp_username = "mail address"
smtp_password = 'password'

# E-Mail Vorlage
email_subject = f"Subject"
# E-Mail-Body aus Textdatei lesen
with open("pathtofile/checkout.html", "r", encoding='utf-8') as file:
    email_body_template = file.read()

# Lese die Empfänger aus der Excel-Datei
df = pd.read_excel("pathtofile/checkout.xlsx")

# Standard BCC-Adresse
main_recipient = "recipient address"

# Schleife durch die Empfängerdaten
for index, row in df.iterrows():
    email_address = row['Email']

    # E-Mail zusammenstellen
    email_body = email_body_template

    # Erstelle die E-Mail
    msg = MIMEMultipart()
    msg['From'] = smtp_username
    msg['To'] = main_recipient
    msg['Subject'] = email_subject
    msg.attach(MIMEText(email_body, 'plain'))

    # Füge BCC-Empfänger hinzu
    bcc_list = [email_address]  # Füge die aktuelle E-Mail-Adresse zu BCC hinzu
    for i in range(100 - 1):  # Füge 29 weitere BCC-Empfänger hinzu (insgesamt 30)
        if index + i + 1 < len(df):
            bcc_list.append(df.iloc[index + i + 1]['Email'])  # Nächste E-Mail-Adresse aus der Liste nehmen

    msg['Bcc'] = ', '.join(bcc_list)

    # Sende die E-Mail
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.sendmail(smtp_username, main_recipient, msg.as_string())

        print(f"E-Mail an {email_address} mit BCC an {', '.join(bcc_list)} gesendet.")
    except Exception as e:
        print(f"Fehler beim Senden der E-Mail an {email_address}: {e}")

    # Verzögerung einfügen, um den Server nicht zu überlasten
    time.sleep(1)
