import pandas as pd
import PyPDF2
from PyPDF2.generic import NameObject, TextStringObject, BooleanObject, NumberObject
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

data = pd.read_excel("path_to_file/filename.xlsx")
counter = 0

from_email = "sender address"
from_password = "password"

server = smtplib.SMTP('SMTP server', 587)
server.starttls()
server.login(from_email, from_password)


def fill_pdf_template(pdf_template_path, output_pdf_path, row):
    with open(pdf_template_path, "rb") as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        pdf_writer = PyPDF2.PdfWriter()

        for page_num in range(len(pdf_reader.pages)):
            pdf_writer.add_page(pdf_reader.pages[page_num])

        pdf_writer._root_object.update({
            NameObject("/NeedAppearances"): BooleanObject(True)
        })

        for page_num in range(len(pdf_reader.pages)):
            page = pdf_writer.pages[page_num]
            if '/Annots' in page:
                for annot in page['/Annots']:
                    annot_obj = annot.get_object()
                    if annot_obj.get('/Subtype') == '/Widget' and annot_obj.get('/T'):
                        field_name = annot_obj.get('/T')
                        if field_name in row:
                            field_value = str(row[field_name])
                            annot_obj.update({
                                NameObject('/V'): TextStringObject(field_value),
                                #NameObject('/Ff'): BooleanObject(False),  # Set read-only to false if needed
                                NameObject('/Ff'): NumberObject(4096)  # Set multi-line text
                            })

        with open(output_pdf_path, "wb") as output_pdf_file:
            pdf_writer.write(output_pdf_file)

def send_email(subject, body, to_email, attachment_path):
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'html'))

    with open(attachment_path, "rb") as attachment:
        part = MIMEApplication(attachment.read(), Name=attachment_path)
        part['Content-Disposition'] = f'attachment; filename="{attachment_path}"'
        msg.attach(part)

        server.sendmail(from_email, to_email, msg.as_string())

# Mapping erstellen (falls notwendig)
"""
mapping = {
    "CSV_Spaltenname1": "PDF_Feldname1",
    "CSV_Spaltenname2": "PDF_Feldname2",
    # Fügen Sie alle notwendigen Zuordnungen hinzu
}
"""

# E-Mail-Body aus Textdatei lesen
with open("path_to_file/filename.html", "r", encoding='utf-8') as file:
    email_body_template = file.read()

for index, row in data.iterrows():

    #data_dict = {mapping.get(key, key): value for key, value in row.items()}
    pdf_file_name = f"{row['Feld2']}.pdf"
    pdf_output_path = f"Unterstuetzungserklaerung_ausgefuellt_{pdf_file_name}"

    birthdatemissing = ("Wir haben für diese Anmeldung kein Geburtsdatum erhalten. "
                        + "Bitte nicht vergessen, dieses in die Unterstützungserklärung "
                        + "statt des Platzhalters einzutragen!") if row['Feld4'] == 0 and row['Feld5'] == 0 else ""
    if birthdatemissing:
        print(f"{row['Feld2']}" + ": " + birthdatemissing)

    # PDF ausfüllen und speichern
    fill_pdf_template("path_to_file/UE_blank.pdf", pdf_output_path, row)

    # E-Mail versenden
    email_subject = f"Erinnerung Unterstuetzungserklaerung fuer {row['Feld2']}"
    email_body = email_body_template.format(name=row['Feld2'], gebdat=birthdatemissing)
    send_email(email_subject, email_body, row['Email'], pdf_output_path)
    counter += 1
    print("Zähler: " + str(counter) + " - Versand erledigt an: " + row['Email'])
