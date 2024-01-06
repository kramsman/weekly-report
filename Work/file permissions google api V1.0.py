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
###


# If modifying these scopes, delete the file token.json.
# SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SCOPES = ['https://www.googleapis.com/auth/drive']
TEST_PERMISSIONS_FOLDER_ID = '1gJKzHtJo6UAtZo4YKP3TBa7NQIY_Lgpz'

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
        page_token = None
        while True:
            # Master VL report folder 1OEAdTnhQoAKpyzFqYdM1ztr42vHAqbk5
            # Postcarding folder 1ajDYyNLaqTKA8jGXDSjiCjCdFjYTJZ6E
            # stuff sheet 1TO1WDZSZir4R1CPDJzEyBzswiOf34BgBeOulICr1WXw
            # JT Racial Justice folder 1Xw-4BP7g4aAZOhTpuR1bmqLmOWNKW_6d

            perms=drive_service.permissions().list(fileId = TEST_PERMISSIONS_FOLDER_ID,
                    fields='permissions/id,permissions/displayName,permissions/role,permissions/emailAddress').execute()

            page_token = perms.get('nextPageToken', None)
            if page_token is None:
                break
            # result = perms.get('permissions', [])
            # result = perms.get('permissions').get('id')


        perm_id_list =[]
        for p in perms.get('permissions', []):
            # Process change
            # print('Found file: %s (%s)' % (file.get('name'), file.get('id'), file.get('trashed')))
            if p.get('role') != 'owner':
                # perm_id_list.append(p.get('id'))
                perm_id_list.append([p.get('id'),p.get('emailAddress')])
            print('Info:', (p.get('displayName'), p.get('id'),p.get('role')))

        a=1

        for id, email in perm_id_list:
            print('going to delete permission with id, email: ',id, email)
            drive_service.permissions().delete(fileId = TEST_PERMISSIONS_FOLDER_ID, permissionId=id, supportsAllDrives = True).execute()

        perm_add_list = ['dee@centerforcommonground.org', 'comstockrov@gmail.com', 'nancy@centerforcommonground.org']
        perm_add_list = ['comstockrov@gmail.com']
        perm_add_list = ['briank@kramericore.com']

        for email in perm_add_list:
            print('ready to add permission for email: ', email)
            payload = {
                "role": "reader",
                "type": "user",
                "emailAddress" : email
            }
            emailMessageMsg = "This is a link to the new VL report.  Click it to open"
            drive_service.permissions().create(fileId=TEST_PERMISSIONS_FOLDER_ID, sendNotificationEmail = True, emailMessage = emailMessageMsg, body=payload).execute()
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
