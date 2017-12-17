#!/usr/bin/env python3

from __future__ import print_function
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate
from gmail_secret import gmail_user, gmail_pswd
import httplib2
import os
import re
import smtplib

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-cotus-checker.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API for Python'

EMAIL_REGEX = re.compile(r"[^@]+@[^@]+\.[^@]+")


def get_credentials(args, my_dirname):
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """

    credential_dir = os.path.join(my_dirname, '.credentials')
    if not os.path.exists(credential_dir):
        os.mkdir(credential_dir)
    credential_path = os.path.join(credential_dir, 'sheets.googleapis.com-cotus-checker.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store, args)
        print('Storing credentials to ' + credential_path)
    return credentials


def get_data_from_sheet(args, my_dirname):
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object
    """
    row_num = 2
    file_name = os.path.join(my_dirname, 'google_sheet.log')
    if os.path.isfile(file_name):
        row_num = int(open(file_name, 'r').read())

    credentials = get_credentials(args, my_dirname)
    http = credentials.authorize(httplib2.Http())
    discovery_url = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
    service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=discovery_url)

    spreadsheet_id = '1FWYQBZLjvVLFrp88BbPmPJXE0wDZDY_L73y7VQIeRFI'
    range_name = 'Form Responses 1!B{0}:F'.format(row_num)
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])

    if values:
        orders = []
        for row in values:
            row = list(map(str.strip, row))
            if row[1] == 'VIN':
                if len(row[4]) != 17 or not row[4].isalnum():
                    info = ', '.join(['VIN', row[4].upper().strip(), row[0].lower().strip()])
                    print(info)
                    print('Invalid Order.\n')
                    # send_email_invalid_order(info, row[0])
                    continue
                orders.append(','.join(['vin', row[4].upper().strip(), row[0].lower().strip()]))
            else:
                if len(row[2]) != 4 or len(row[3]) != 6 or not row[2].isalnum() or not row[3].isalnum():
                    info = ', '.join(['Order Number & Dealer Code', row[2].upper().strip(), row[3].upper().strip(), row[0].lower().strip()])
                    print(info)
                    print('Invalid Order.\n')
                    # send_email_invalid_order(info, row[0])
                    continue
                orders.append(','.join(['num', row[2].upper().strip(), row[3].upper().strip(), row[0].lower().strip()]))

        row_num += len(values)
        open(file_name, 'w').write(str(row_num))
        return orders

    return []


def send_email_invalid_order(info, email):
    if not EMAIL_REGEX.match(email):
        return -1, 'Invalid email address.'
    print('send email for invalid order')
    print(info)
    if not gmail_user or not gmail_pswd:
        return -1, 'Empty Gmail Username or Password'
    else:
        email_from = gmail_user
        email_body = 'The information you entered is invalid.\nPlease make sure it works on the actual COTUS website (http://www.cotus.ford.com) before you register with the auto checker.\n\n'
        email_body += 'The information you entered: {0}'.format(info)

        email_msg = MIMEMultipart()
        email_msg['Subject'] = '[COTUS CHECKER] Invalid Information'
        email_msg['From'] = gmail_user
        email_msg['To'] = email
        email_msg['Date'] = formatdate(localtime=True)
        email_msg.attach(MIMEText(email_body))

        try:
            gmail_server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            gmail_server.ehlo()
            gmail_server.login(gmail_user, gmail_pswd)
            gmail_server.sendmail(email_from, email, email_msg.as_string())
            gmail_server.close()
            return 0, 'SUCCESS'
        except KeyboardInterrupt:
            exit(2)
        except (smtplib.SMTPException, smtplib.SMTPServerDisconnected, smtplib.SMTPResponseException,
                smtplib.SMTPSenderRefused, smtplib.SMTPRecipientsRefused, smtplib.SMTPDataError, smtplib.SMTPConnectError,
                smtplib.SMTPHeloError, smtplib.SMTPNotSupportedError, smtplib.SMTPAuthenticationError):
            return -1, 'FAIL'
