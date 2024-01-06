# Created 1/23/2022 - take reports from Python/Dropbox and copy them to google drive
# TODO: add this code to the report producing program
# TODO: Could we assign read permission on a file level by looking up Orgs? - Set message to say a new VL report is available?


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
REPORT_FOLDER_ID = "1OEAdTnhQoAKpyzFqYdM1ztr42vHAqbk5"  # Folder holding VoterLetters Organizer Reports

def getFolderId(foldernm, par):
    q1 = "mimeType='application/vnd.google-apps.folder' and trashed=false and name='" + foldernm + "' and parents = '" + par +"'"
    idList = []
    page_token = None
    while True:
        response = drive_service.files().list(
            q=q1,
            spaces='drive',
            fields="nextPageToken, files(id, name, trashed, parents)",
            pageToken=page_token).execute()
        print(response)
        for file in response.get('files', []):
            idList.append(file.get('id'))
            print('Found folder (name,id,trashed,parent):', (file.get('name'), file.get('id'), file.get('trashed'), file.get('parents')))

        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break
    return idList

def shareWithEveryone(folderId, service):
    payload = {
        "role": "reader",
        "type": "anyone"
    }
    drive_service.permissions().create(fileId=folderId, body=payload).execute()

### MAIN ###
def main():
    choice = pymsgbox.confirm("Run Reports or upload files to Google Drive", 'Run, Upload or Exit?', ["Run", 'Upload', 'Exit'])
    if choice == "Exit":
        pymsgbox.alert("Exiting", "Alert")
        exit()
    elif choice == "Run":
        # TODO: report production goes here
        pymsgbox.alert("** Run Here  **", "Alert")
        exit()

    Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    # path_input_str = '/Users/Denise/Dropbox/Postcard Files/VoterLetters Reports/all-parent-campaigns-requests-2022-01-20-W'

    dirStartLoc = os.path.expanduser("~/Dropbox/Postcard Files/VL Org Reports/")
    path_input_str = askdirectory(initialdir=dirStartLoc, title="Select INPUT directory")
    if path_input_str == "":
        pymsgbox.alert("No directory chosen - exiting", "Alert")
        exit()
    path_input = pathlib.Path(path_input_str)
    folderName = path_input.parts[-1]  # the last section of the path

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    # if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    global drive_service
    drive_service = build('drive', 'v3', credentials=creds)  # copied from https://stackoverflow.com/questions/11472401/looking-for-example-using-mediafileupload

    subFolderIdsInReportFolder = getFolderId(folderName, REPORT_FOLDER_ID) # folder name beig searched, list to put ids into, drive service, parent id

    try:

        # Send all sub folders from above with specified name and parent to trash (trashes files too)
        for folderId in subFolderIdsInReportFolder:
            updated_file = drive_service.files().update(fileId=folderId, body={'trashed': True}).execute()
        # pymsgbox.alert("trashed the following folder: \n\n Need to insert list here", "Trashed files")

        # create sub folder for this week's reports under Master Report Folder
        file_metadata = {
            'name': folderName,
            'parents': [REPORT_FOLDER_ID],
            'mimeType': 'application/vnd.google-apps.folder'
        }

        file = drive_service.files().create(body=file_metadata,
                                            fields='id').execute()
        newFileId = file.get('id')
        print('New Folder ID: %s' % newFileId)


        # get list of filenames in computer directory
        xlsxFiles = glob.glob(str(path_input / '[!~]*.xlsx'))  # xlsx files not starting with ~ (temporary files)
        for  fileWPath in xlsxFiles:
            # filePath = pathlib.Path('/Users/Denise/Dropbox/Postcard Files/VoterLetters Reports/all-parent-campaigns-requests-2022-01-20-W')
            # path_parent = reportDirToUpload.resolve().parent
            # path, fileNameWOPath = os.path.split(fileWPath)
            fileName = os.path.basename(fileWPath)  # returns filename with extension
            fileNameWOExt = pathlib.Path(fileWPath).stem
            print('Uploading ', fileName)
            # fileWPath = os.path('/Users/Denise/Dropbox/Postcard Files/VoterLetters Reports/all-parent-campaigns-requests-2022-01-20-W/CA-Alameda County VoterLetters Summary 2022-01-20-W.xlsx')
            str_file_metadata = "{'name': '" + fileNameWOExt +"' ,"\
                        "'mimeType': 'application/vnd.google-apps.spreadsheet' ," +\
                        "\"parents\": [\"" + newFileId + "\"]}"

            file_metadata = ast.literal_eval(str_file_metadata) # convert string to dictionary object

            media = MediaFileUpload(fileWPath, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            file = drive_service.files().create(body=file_metadata, supportsAllDrives= True,
                                            media_body=media,
                                            fields='id').execute()
            # print('File ID: %s' % file.get('id'))

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()