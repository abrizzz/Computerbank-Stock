#!/usr/bin/python3
from __future__ import print_function
import httplib2
import os
import sys
from keys import *
from reader import *
import time
import argparse

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

APPLICATION_NAME = 'Computerbank Stock'


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
                                   JSON_FILE)

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

def writeToSheet(service,stock,range):
    values = []
    for d in stock:
        values.append(d.toArray())

    body = {
      'values': values
    }

    clearSheet(service,range)
    service.spreadsheets().values().update(spreadsheetId=spreadsheetId,
    range=range,valueInputOption='USER_ENTERED',body=body).execute()

def clearSheet(service,range):
    service.spreadsheets().values().clear(spreadsheetId=spreadsheetId,
    range=range,body={}).execute()


def main():
    """Write to Sheet"""

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    desktop_sheet, laptop_sheet = getSheets(ODS_FILE)

    desktop_stock = getDesktops(desktop_sheet)
    writeToSheet(service,desktop_stock,'Desktops!2:100')
    laptop_stock = getLaptops(laptop_sheet)
    writeToSheet(service,laptop_stock,'Laptops!2:100')

    while True:
        desktop_sheet_new, laptop_sheet_new = getSheets(ODS_FILE)
        if(desktop_sheet.array != desktop_sheet_new.array):
            desktop_sheet = desktop_sheet_new
            desktop_stock = getDesktops(desktop_sheet)
            writeToSheet(service,desktop_stock,'Desktops!2:100')

        if(laptop_sheet.array != laptop_sheet_new.array):
            laptop_sheet = laptop_sheet_new
            laptop_stock = getLaptops(laptop_sheet)
            writeToSheet(service,laptop_stock,'Laptops!2:100')
        time.sleep(INTERVAL)




if __name__ == '__main__':
    main()
