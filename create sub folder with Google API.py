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


#ROV Requests
SAMPLE_SPREADSHEET_ID = '1TO1WDZSZir4R1CPDJzEyBzswiOf34BgBeOulICr1WXw'
SAMPLE_RANGE_NAME = 'Form Responses 1!O1:C'

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
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

    drive_service = build('drive', 'v3', credentials=creds)  # copied from https://stackoverflow.com/questions/11472401/looking-for-example-using-mediafileupload

    try:
        file_metadata = {
            'name': 'sub folder',
            'parents': ['1OEAdTnhQoAKpyzFqYdM1ztr42vHAqbk5'],
            'mimeType': 'application/vnd.google-apps.folder'
        }

        file = drive_service.files().create(body=file_metadata,
                                            fields='id').execute()
        newFileId = file.get('id')
        print('Folder ID: %s' % file.get('id'))
        quit()

# test api folder id 1OEAdTnhQoAKpyzFqYdM1ztr42vHAqbk5

# DELETE https://www.googleapis.com/drive/v2/files/folderId/children/childId

        # delete all files in folder except excluded
        # dir_id = "my folder Id"
        # file_id = "avoid deleting this file"
        # service.files().update(fileId=file_id, addParents="root", removeParents=dir_id).execute()
        # service.files().delete(fileId=dir_id).execute()  # trash instead of delete moves to trash; delete gone for good

        # gfile = drive_service.CreateFile({'parents': [{'id': '1OEAdTnhQoAKpyzFqYdM1ztr42vHAqbk5'}]})
        # # Read file and set it as the content of this instance.
        # gfile.SetContentFile('CherrySugarCookiescopy.pdf')
        # gfile.Upload()  # Upload the file.

        # file_metadata = {'name': 'CherrySugarCookiescopy.pdf'}

        # file_metadata = {'name': 'CherrySugarCookiescopy.pdf', "parents": ["1OEAdTnhQoAKpyzFqYdM1ztr42vHAqbk5"]}

        file_metadata = {'name': 'Sparta Thermostat settings',
                         'mimeType': 'application/vnd.google-apps.spreadsheet',
                         "parents": ["1OEAdTnhQoAKpyzFqYdM1ztr42vHAqbk5"]}

        #                  'mimetype': 'application/vnd.google-apps.spreadsheet',
        # media = MediaFileUpload('/Users/Denise/Documents/Sparta Thermostat settings.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') # upload xlsx
        # media = MediaFileUpload('/Users/Denise/Documents/Sparta Thermostat settings.xlsx', mimetype='application/vnd.google-apps.spreadsheet') # converts xlsx

        media = MediaFileUpload('/Users/Denise/Dropbox/Postcard Files/VoterLetters Reports/all-parent-campaigns-requests-2022-01-20-W/CA-Alameda County VoterLetters Summary 2022-01-20-W.xlsx',
                                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') # converts xlsx

            # application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
            # application/pdf
            # text/csv
            # image/jpeg or image/png
            # text/plain

        file = drive_service.files().create(body=file_metadata, supportsAllDrives= True,
                                            media_body=media,
                                            fields='id').execute()
        print('File ID: %s' % file.get('id'))

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()