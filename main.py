'''
A Python script to onboard Sunbelt Rentals' top 100 Customers
for Invoice Integration by sending mail individually.
'''

import smtplib
import xlrd
from password import mailId, password
from email.message import EmailMessage


def send_mail(to, cc, bcc, name, company, portal):
    '''
    Function to send email when called

    :param to: 'to' email address
    :param cc: 'cc' email address
    :param bcc: 'bcc' email address
    :param name: customer's name
    :param company: company's name
    :return: None

    '''

    # Setting up the SMTP server through secure SSL encryption
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(mailId, password)

        # Subject of the mail
        subject = f"Integration acknowledgement for {company}"
        # Body of the mail
        body = f"{name}\n\n" \
               f"The body of the mail."

        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = mailId
        msg['To'] = to
        msg['CC'] = ','.join(cc)

        msg.set_content(body)

        with open("Sample_attachment.pdf", "rb") as f:
            file_data = f.read()
            file_name = f.name
            msg.add_attachment(file_data, maintype='application', subtype="pdf", filename=file_name)

        # Sending the mail to the 'To' address
        # msg = f"Subject:{subject}\nTO:{to}\nCC:{cc}\n\n{body}"
        smtp.sendmail(mailId, to, msg.as_string())

        # Sending the mail to the 'CC' address
        if len(cc) != 1 or (len(cc) == 1 and cc[0] != ''):  # To check if CC is empty
            for i in cc:
                # msg = f"Subject:{subject}\nTO:{to}\nCC:{','.join(cc)}\n\n{body}"
                smtp.sendmail(mailId, i, msg.as_string())

        # Sending the mail to the 'BCC' address
        if len(bcc) != 1 or (len(bcc) == 1 and bcc[0] != ''):  # To check if BCC is empty
            for i in bcc:
                msg = EmailMessage()
                msg['Subject'] = subject
                msg['From'] = mailId
                msg['To'] = i
                # msg = f"Subject:{subject}\nTO:{i}\n\n{body}"
                smtp.sendmail(mailId, i, msg.as_string())


if __name__ == "__main__":

    # Opening the file
    doc = xlrd.open_workbook("Dummy_onboarding.xls")
    # Choosing the sheet
    sheet = doc.sheet_by_index(0)

    # Extracting the row count
    num_of_rows = sheet.nrows

    # Looping through each row
    for i in range(1, num_of_rows):
        # Extracting the data from the sheet
        company = sheet.cell_value(i, 0)
        name = sheet.cell_value(i, 1)
        to = sheet.cell_value(i, 2)
        cc = sheet.cell_value(i, 3).split(',')
        bcc = sheet.cell_value(i, 4).split(',')
        portal = sheet.cell_value(i, 5)

        # Calling the send_mail function to send the mail
        send_mail(to, cc, bcc, name, company, portal)
        print(i, "mail sent!")
