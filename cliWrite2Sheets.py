#!/usr/local/bin/python3

from __future__ import print_function

import os.path
import datetime
import sys
## Uncomment webbrowser for local testing
import webbrowser

from column_width import auto_resize_columns

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SHARED_FOLDER_ID = '1WYxIIkLXa5wQNsPZDu6O0drN1FRREh-X' ## 


badge_names = {}
# Read in badge ids and corresponding names
with open('badge_names.txt') as f:
    for line in f:
        badge, name = line.strip().split(' ', maxsplit=1)
        badge_names[badge] = name

# read environment variables non-interactively (e.g. ./write2sheets.py <EventName> <BadgeID> <in>)
event = sys.argv[1]
badgeid = sys.argv[2]
inout = sys.argv[3]

# Retrieve name from badge_names dictionary, using badgeid as the default value if the key does not exist
name = badge_names.get(badgeid, badgeid)

#name = badge_names.get(badgeid)
#if name is None:
    #name = badgeid

now = datetime.datetime.now()
date_str = now.strftime('%Y-%m-%d')
date_time_str = now.strftime('%m/%d/%Y %H:%M:%S') # 12/22/2022 17:46:01 (preferred Google Sheets date/time format)
SHEET_NAME = f'{date_str}_{event}'

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

def main():
    creds = None
    # test for previous authentication 
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
    try:
        sheets_service = build('sheets', 'v4', credentials=creds)
        drive_service = build('drive', 'v3', credentials=creds)

        results = drive_service.files().list(q=f"'{SHARED_FOLDER_ID}' in parents", fields="nextPageToken, files(id, name)").execute()
        items = results.get("files", [])
        file_names = [item["name"] for item in items]
        
        if SHEET_NAME not in file_names:
            # Create new Google Sheet
            file_metadata = {
              'name': SHEET_NAME,
              'mimeType': 'application/vnd.google-apps.spreadsheet',
              'parents': [SHARED_FOLDER_ID]
          }
            file = drive_service.files().create(body=file_metadata).execute()
            file_id = file['id']
            print(f'Sheet created with ID: {file_id}')
            # Write the column headers to the first row of the sheet
            result = sheets_service.spreadsheets().values().update(
                    spreadsheetId=file_id,
                    range='Sheet1!A1:D1',
                    valueInputOption='RAW',
                    body={'values': [['Name', 'CheckinTime', 'CheckoutTime', 'Duration']]}
            ).execute()
            ## Uncomment webbrowser.open for local testing
            webbrowser.open(f'https://docs.google.com/spreadsheets/d/' + file_id, new=2)
        else:
            # Get the ID of the existing sheet
            file_id = [item["id"] for item in items if item["name"] == SHEET_NAME][0]
    
        # Retrieve the values in the sheet using the spreadsheets().values().get() method
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=file_id,
            range='Sheet1!A1:B'  # retrieve the values in columns A and B only
        ).execute()
        rows = result.get('values', [])
        found_row = None
        found_row_values = None
        # Iterate over the rows in the sheet and check if the first cell in each row matches the value of 'name'
        for i, row in enumerate(rows):
            if row[0] == name:
                found_row = i
                found_row_values = row
                break
        if found_row is not None:
            # Check the value of 'inout'
            if inout == 'in':
                # Update the cell in column B of the found row
                result = sheets_service.spreadsheets().values().update(
                    spreadsheetId=file_id,
                    range=f'Sheet1!B{found_row+1}',
                    valueInputOption='USER_ENTERED',
                    body={'values': [[date_time_str]]}
                ).execute()
            elif inout == 'out':
                # Update the cell in column C of the found row
                result = sheets_service.spreadsheets().values().update(
                    spreadsheetId=file_id,
                    range=f'Sheet1!C{found_row+1}',
                    valueInputOption='USER_ENTERED',
                    body={'values': [[date_time_str]]}
                ).execute()
                # Update the cell in column D of the found row with the formula to calculate the duration
                result = sheets_service.spreadsheets().values().update(
                    spreadsheetId=file_id,
                    range=f'Sheet1!D{found_row+1}',
                    valueInputOption='USER_ENTERED',
                    body={'values': [[f"=TEXT(C{found_row+1}-B{found_row+1},\"h:mm:ss\")"]]}
                ).execute()
                auto_resize_columns(sheets_service, file_id, 0, 2, 3)
                auto_resize_columns(sheets_service, file_id, 0, 3, 4)
        else:
            # Append the new row to the sheet using the spreadsheets().values().append() method
            ROWS = [[name, date_time_str]]
            sheet_range = f'Sheet1!A1:D{len(ROWS)+1}'
            sheet = sheets_service.spreadsheets()
            result = sheet.values().append(
                spreadsheetId=file_id,
                range=sheet_range,
                insertDataOption='INSERT_ROWS',
                valueInputOption='USER_ENTERED',
                body={'values': ROWS}
            ).execute()
            print(f'{result["updates"]["updatedRows"]} rows appended to sheet.')
        auto_resize_columns(sheets_service, file_id, 0, 0, 1)
        auto_resize_columns(sheets_service, file_id, 0, 1, 2)
    except HttpError as err:
        print(f'An error occured: {err}')

if __name__ == '__main__':
    main()
