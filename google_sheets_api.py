#!/usr/bin/env python3

from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API for Python'

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
    if args:
      credentials = tools.run_flow(flow, store, args)
    else: # Needed only for compatibility with Python 2.6
      credentials = tools.run(flow, store)
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
  discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
  service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=discoveryUrl)

  spreadsheetId = '1FWYQBZLjvVLFrp88BbPmPJXE0wDZDY_L73y7VQIeRFI'
  rangeName = 'Form Responses 1!B{0}:F'.format(row_num)
  result = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=rangeName).execute()
  values = result.get('values', [])

  if values:
    orders = []
    for row in values:
      if row[1] == 'VIN':
        orders.append(','.join(['vin', row[4].upper().strip(), row[0].lower().strip()]))
      else:
        orders.append(','.join(['num', row[2].upper().strip(), row[3].upper().strip(), row[0].lower().strip()]))

    row_num += len(values)
    open(file_name, 'w').write(str(row_num))
    return orders

  return []
