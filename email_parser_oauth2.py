import os.path
import base64
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import openpyxl

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def authenticate_gmail():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('gmail', 'v1', credentials=creds)
    return service

def fetch_emails(service):
    results = service.users().messages().list(userId='me', labelIds=['INBOX']).execute()
    messages = results.get('messages', [])

    email_data = []
    for message in messages:
        msg = service.users().messages().get(userId='me', id=message['id']).execute()
        msg_data = msg.get('payload', {})
        headers = msg_data.get('headers', [])

        email_dict = {'subject': None, 'from': None, 'date': None, 'body': None}
        for header in headers:
            if header['name'] == 'Subject':
                email_dict['subject'] = header['value']
            elif header['name'] == 'From':
                email_dict['from'] = header['value']
            elif header['name'] == 'Date':
                email_dict['date'] = header['value']

        if 'parts' in msg_data:
            for part in msg_data['parts']:
                if part['mimeType'] == 'text/plain':
                    body_data = base64.urlsafe_b64decode(part['body']['data'].encode('UTF-8')).decode('UTF-8')
                    email_dict['body'] = body_data

        email_data.append(email_dict)

    return email_data

def store_data_in_excel(parsed_emails, filename='emails.xlsx'):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Emails'

    headers = ['Subject', 'From', 'Date', 'Body']
    sheet.append(headers)

    for email in parsed_emails:
        sheet.append([
            email['subject'],
            email['from'],
            email['date'],
            email['body']
        ])

    workbook.save(filename)

def main():
    service = authenticate_gmail()
    parsed_emails = fetch_emails(service)
    store_data_in_excel(parsed_emails)

if __name__ == '__main__':
    main()
