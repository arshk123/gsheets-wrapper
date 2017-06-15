from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from pprint import pprint
from googleapiclient import discovery

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def get_sheet(spreadsheetId, discoveryUrl, http, service, range):
    """Gets sheet from Google Drive API

    Returns: Spreadsheet object
    """
    return service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=range).execute()


def write_to_sheet(spreadsheetId, service, range):
    values = [["IoTs"]]
    body = {'values': values}
    result = service.spreadsheets().values().update(spreadsheetId=spreadsheetId, range=range, valueInputOption="USER_ENTERED", body=body).execute()
    # print("writing to sheet")

def display_sheet(sheet):
    """ Function to print out rows of the sheet """

    values = sheet.get('values', [])
    if not values:
        print('No data found.')
    else:
        # parse through sheet and print out each row/line
        for row in values:
            for cell in row:
                print(cell, end="")
            print("")

def main():
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    https://docs.google.com/spreadsheets/d/1NuLvG1X3gzGrVU8qAca94ZwrAWt5Gba9ScHLRX0JZfk/edit
    """

    # Attributes
    credentials = get_credentials()
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    http = credentials.authorize(httplib2.Http())

    service = discovery.build('sheets',
                              'v4',
                               http=http,
                               discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '19rkIrd_-VASvonp2MV6WZTio2HlAYTS3IEDA11tlZXM'

    range = "A:V"
    sheet = get_sheet(spreadsheetId=spreadsheetId,
                      http=http,
                      service=service,
                      discoveryUrl=discoveryUrl,
                      range='A:V')

    # display_sheet(sheet)

    write_to_sheet(spreadsheetId,
                   service,
                   range="V2:V")

if __name__ == '__main__':
    main()
