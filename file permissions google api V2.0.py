# V2.0 - assign permissions for each report instead of directory.
# necessitates reading in VL users file to match report with room / organizer
# /Users/Denise/Downloads/enterprise-users-2022-01-21.csv

sendPermissionAddEmail = False

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
import pandas as pd
import glob
import pathlib
import pymsgbox
import ast
from tkinter import Tk  # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from tkinter import messagebox

###

# If modifying these scopes, delete the file token.json.
# SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SCOPES = ['https://www.googleapis.com/auth/drive']
TEST_PERMISSIONS_FOLDER_ID = '1gJKzHtJo6UAtZo4YKP3TBa7NQIY_Lgpz'

def main():

    if False:
        df = pd.read_csv("/Users/Denise/Downloads/enterprise-users-2022-01-21.csv",usecols = ['name','email','role','organization'])
        df = df.sort_values(["organization", "name"], ascending=(True, True))
        orgEmailList = df.loc[df['role'].isin(['organizer']),["organization", "email"]].values.tolist()
    else:
        orgEmailList = [
            ['CA-Peninsula-South Bay','dee@centerforcommonground.org'],
            ['Mountain States','comstockrov@gmail.com'],
            ['CA-Peninsula-South Bay','nancy@centerforcommonground.org'],
            ['NY-NYC and Long Island','kramsman+nycorg@gmail.com']
        ]

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

        dirStartLoc = os.path.expanduser("~/Dropbox/Postcard Files/VL Org Reports/")
        path_input_str = askdirectory(initialdir=dirStartLoc, title="Select INPUT directory")
        if path_input_str == "":
            pymsgbox.alert("No directory chosen - exiting", "Alert")
            exit()
        path_input = pathlib.Path(path_input_str)
        folderName = path_input.parts[-1]  # the last section of the path

        # TODO: below will be replaced with reading in VL user file
        perm_add_list = ['dee@centerforcommonground.org', 'comstockrov@gmail.com', 'nancy@centerforcommonground.org']
        perm_add_list = ['comstockrov@gmail.com']

        # get list of filenames in computer directory
        xlsxReportFiles = glob.glob(str(path_input / '[!~]*.xlsx'))  # xlsx files not starting with ~ (temporary files)
        for fileWPath in xlsxReportFiles:
            # filePath = pathlib.Path('/Users/Denise/Dropbox/Postcard Files/VoterLetters Reports/all-parent-campaigns-requests-2022-01-20-W')
            # path, fileNameWOPath = os.path.split(fileWPath)
            fileName = os.path.basename(fileWPath)  # returns filename with extension
            fileNameWOExt = pathlib.Path(fileWPath).stem
            roomName = fileNameWOExt[0:-34]
            print('Uploading, RoomName ', fileName, roomName)
            # TODO: add code to upload report here

            for org, email in orgEmailList:
                # email = 'briank@kramericore.com'
                if org == roomName:
                    print('ready to add permission for email (org, roomName, email): ', org, roomName, email)
                    payload = {
                        "role": "reader",
                        "type": "user",
                        "emailAddress" : email
                    }
                    if sendPermissionAddEmail:
                        emailMessageMsg = "A new VoterLetters summary report is available for your room.  Click to open."
                    # a NYC report '1arP8lS8j_ovLyDyKV87NPfJcXav1BV7GjaKVLW6xuHY'  # FIXME: need to substitute uploaded file id in create premission below
       #             drive_service.permissions().create(fileId='1arP8lS8j_ovLyDyKV87NPfJcXav1BV7GjaKVLW6xuHY', sendNotificationEmail = True, emailMessage = emailMessageMsg, body=payload).execute()
                    else:
                        a=1
                        # Same as above .create command but do not include email text as not allowed if no email sent
        #             drive_service.permissions().create(fileId='1arP8lS8j_ovLyDyKV87NPfJcXav1BV7GjaKVLW6xuHY', sendNotificationEmail = False, body=payload).execute()

                # TODO: put notification message in: emailMessage = text
                    # drive_service.permissions().create(fileId = TEST_PERMISSIONS_FOLDER_ID, type='user', role='reader', emailAddress=email, sendNotificationEmail=False).execute()

# sendNotificationEmail = false, supportsAllDrives = true, role = reader, type = user, emailAddress
# {'permissions': [
            # {'id': '09934718075527879418', 'displayName': 'Dee Nelson', 'emailAddress': 'dee@centerforcommonground.org', 'role': 'writer'},
            # {'id': '01359456030132609282', 'displayName': 'Nancy Goodban', 'emailAddress': 'nancy@centerforcommonground.org', 'role': 'writer'},
            # {'id': '14513457487962141098', 'displayName': 'Barbara Comstock', 'emailAddress': 'comstockrov@gmail.com', 'role': 'reader'},
            # {'id': '08725835372821687358', 'displayName': 'kramsmann', 'emailAddress': 'kramsman@gmail.com', 'role': 'owner'}
            # ]}

    except HttpError as err:
        print(err)

if __name__ == '__main__':
    main()

    # {'kind': 'drive#permissionList', 'permissions': [
    #     {'id': '14986069575103011050', 'displayName': 'Dee Nelson - CFCG', 'type': 'user',
    #      'kind': 'drive#permission',
    #      'photoLink': 'https://lh3.googleusercontent.com/a-/AOh14GggOlpzHdIL2k5uQuK834Ie1J5HZHjhl0aEQPzR=s64',
    #      'emailAddress': 'rov.scc.ca@gmail.com', 'role': 'writer', 'deleted': False, 'pendingOwner': False},
    #     {'id': '10933979155546302898', 'displayName': 'ROV DCMV', 'type': 'user', 'kind': 'drive#permission',
    #      'photoLink': 'https://lh3.googleusercontent.com/a/default-user=s64',
    #      'emailAddress': 'rov.dcmv@gmail.com', 'role': 'reader', 'deleted': False, 'pendingOwner': False},
    #     {'id': '08725835372821687358', 'displayName': 'kramsmann', 'type': 'user', 'kind': 'drive#permission',
    #      'photoLink': 'https://lh3.googleusercontent.com/a/default-user=s64',
    #      'emailAddress': 'kramsman@gmail.com', 'role': 'owner', 'deleted': False, 'pendingOwner': False}]}
    #
