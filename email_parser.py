# pylint: disable=redefined-outer-name
import imaplib
import email
from email.policy import default
import openpyxl

def connect_to_email(username, password, server='imap.gmail.com'):
    mail = imaplib.IMAP4_SSL(server)
    mail.login(username, password)
    return mail

def fetch_emails(mail, folder='inbox'):
    mail.select(folder)
    status, data = mail.search(None, 'ALL')
    email_ids = data[0].split()
    return email_ids

def get_email_content(mail, email_id):
    status, data = mail.fetch(email_id, '(RFC822)')
    raw_email = data[0][1]
    msg = email.message_from_bytes(raw_email, policy=default)
    return msg

def parse_email(msg):
    subject = msg['subject']
    from_ = msg['from']
    date = msg['date']
    
    body = None
    html_body = None
    
    if msg.is_multipart():
        for part in msg.iter_parts():
            if part.get_content_type() == 'text/plain':
                body = part.get_payload(decode=True).decode()
            elif part.get_content_type() == 'text/html':
                html_body = part.get_payload(decode=True).decode()
    else:
        if msg.get_content_type() == 'text/plain':
            body = msg.get_payload(decode=True).decode()
        elif msg.get_content_type() == 'text/html':
            html_body = msg.get_payload(decode=True).decode()
    
    return {'subject': subject, 'from': from_, 'date': date, 'body': body, 'html_body': html_body}

def store_data_in_excel(parsed_emails, filename='emails.xlsx'):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Emails'

    # Write headers
    headers = ['Subject', 'From', 'Date', 'Body', 'HTML Body']
    sheet.append(headers)

    # Write email data
    for email in parsed_emails:
        sheet.append([
            email['subject'],
            email['from'],
            email['date'],
            email['body'],
            email['html_body']
        ])

    workbook.save(filename)

def main(username, password):
    mail = connect_to_email(username, password)
    email_ids = fetch_emails(mail)
    
    parsed_emails = []
    for email_id in email_ids:
        msg = get_email_content(mail, email_id)
        parsed_email = parse_email(msg)
        parsed_emails.append(parsed_email)
    
    store_data_in_excel(parsed_emails)

if __name__ == '__main__':
    username = input('Email: ')
    password = input('Password: ')
    main(username, password)
