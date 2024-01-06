# Quickstart code via George.  Copied from https://developers.google.com/sheets/api/quickstart/python

from __future__ import print_function

import os.path

import google.auth.transport.requests
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
## upload below from https://stackoverflow.com/questions/11472401/looking-for-example-using-mediafileupload
from apiclient.http import MediaFileUpload
from apiclient.discovery import build
# from oauth2client.service_account import ServiceAccountCredentials
###


# If modifying these scopes, delete the file token.json.
# SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SCOPES = ['https://www.googleapis.com/auth/drive']

def main():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:

        drive_service = build('drive', 'v3', credentials=creds)  # copied from https://stackoverflow.com/questions/11472401/looking-for-example-using-mediafileupload

        ADMIN_RPT_FOLDER = '1_vTXsMYOwK_TLYaRqC7YA_jolBuqDWP4'

        # file_metadata = {'name': 'Mockup of StatesRpt 4 from pdf',
        #                  'mimeType': 'application/vnd.google-apps.document',
        #                  "parents": [ADMIN_RPT_FOLDER]}

        file_metadata = {'name': 'Mockup of StatesRpt 4 from ps',
                         'mimeType': 'application/vnd.google-apps.document',
                         "parents": [ADMIN_RPT_FOLDER]}

        # 'mimeType': 'application/pdf',

        #                  'mimetype': 'application/vnd.google-apps.spreadsheet',
        # media = MediaFileUpload('/Users/Denise/Documents/Sparta Thermostat settings.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') # upload xlsx
        # media = MediaFileUpload('/Users/Denise/Documents/Sparta Thermostat settings.xlsx', mimetype='application/vnd.google-apps.spreadsheet') # converts xlsx

        media = MediaFileUpload('/Users/Denise/Dropbox/Postcard Files/Other/Mockup of StateRpt 4.pdf')

        # media = MediaFileUpload('/Users/Denise/Dropbox/Postcard Files/Other/Mockup of StateRpt 4.pdf',
        #                         mimetype='vnd.google-apps.spreadsheet.') # converts xlsx

            # application/vnd.openxmlformats-officedocument.spreadsheetml.sheet

        file = drive_service.files().create(body=file_metadata, supportsAllDrives= True,
                                            media_body=media,
                                            fields='id').execute()
        print('File ID: %s' % file.get('id'))

    except HttpError as err:
        print(err)

if __name__ == '__main__':
    main()