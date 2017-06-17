from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from pprint import pprint
from googleapiclient import discovery
from math import pow

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

# TODO use batch update for updating csv

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


def ind_to_let(index):
    """ Convert numerical index to equivalent column letter """
    letter = ''
    while index > 0:
        temp = (index - 1) % 26
        letter = letter.join(chr(temp + 65))
        index = (index - temp - 1) / 26
    return letter



def let_to_ind(letter):
    """ Function to convert column letter to corresponding numerical index """
    column = 0
    length = len(letter);
    for i in range(length):
        column += (ord(letter[i]) - 64) * pow(26, length - i - 1)
    return int(column)


def get_sheet(spreadsheetId, discoveryUrl, http, service, range):
    """Gets sheet from Google Drive API
    Returns: Spreadsheet object
    """
    return service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=range).execute()


def write_row(spreadsheetId, service, row, values):
    """ Update single row """
    v = [values]
    body = {'values': v}
    result = service.spreadsheets().values().update(spreadsheetId=spreadsheetId,
                                                    range=row,
                                                    valueInputOption="USER_ENTERED",
                                                    body=body).execute()


def fill_data(row, data_cols, fill_cols, labels, service, spreadsheetId, index):
    """ Function to take user input as expected column label
        Input: takes row, the row of the google sheet we are updating
        data_cols, the columns that have filled data that may help identify
        what to fill in...
        fill_cols, the columns we need to fill
        labels, header of csv for identifying what kind of information we have
    """
    for d in data_cols:
        if d <= len(row):
            print(labels[d], row[d])
    answers = []
    for a in fill_cols:
        answer = input("Provide an estimated " + labels[a] + ": ")
        answers.append(answer)

    for idx, ans in enumerate(answers):
        write_row(spreadsheetId,
                   service,
                   row=str(ind_to_let(fill_cols[idx]+1) + str(index)),
                   values=[answers[idx]])


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
    """Shows basic usage of the Sheets wrapper
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

    spreadsheetId = ''

    # specify range of columns to get from API
    range = 'A:Z'

    sheet = get_sheet(spreadsheetId=spreadsheetId,
                      http=http,
                      service=service,
                      discoveryUrl=discoveryUrl,
                      range=range)

    # TODO: make this reusable
    data_cols = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'L', 'M']
    answer_col = [int(let_to_ind('X') - 1)]
    indices = []
    for idx, letter in enumerate(data_cols):
        data_cols[idx] = let_to_ind(letter) - 1

    values = sheet.get('values', [])
    labels = values[0]
    for idx, row_data in enumerate(values[1:]):
        fill_data(row_data, data_cols, answer_col, labels, service, spreadsheetId, index=idx+2)

if __name__ == '__main__':
    main()
