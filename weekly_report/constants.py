"""Constants and configuration values for the weekly report program."""
import sys
from pathlib import Path

CREDS_DIR = "~/.config/weekly-report"

######################
# WEEKLY messages
######################

# ORG_WEEKLY_MSG = ("A new WEEKLY Sincere summary report, THE LAST WEEKLY UNTIL THE MIDTERMS IN MARCH.  "
#                   "To access the sheet you will need to be logged in to Google.  Do this by using the Chrome browser or by going to google.com in another browser.")
# ORG_WEEKLY_MSG = ("The FINAL WEEKLY Sincere summary report is available for your room. "
#                         "IT IS THE LAST UNTIL THE MIDTERMS IN 2026."
#                   "To access the sheet you will need to be logged in to Google.  Do this by using the Chrome browser or by going to google.com in another browser."
#                         )
ORG_WEEKLY_MSG = (f"A new WEEKLY Sincere summary report is available for your room. "
                  f"To access the sheet you will need to be logged in to Google.  Do this by using the Chrome browser or by going to google.com in another browser."
                  f"Click to open.")

# CORE_WEEKLY_MSG = (f"A new WEEKLY ROV-WIDE Sincere summary report, THE LAST UNTIL THE VA GENERAL END OF SUMMER, has been sent to the CORE GROUP. "
#                   f"Click to open.")
# CORE_WEEKLY_MSG = ("The FINAL WEEKLY ROV-WIDE Sincere summary report has been sent to the CORE GROUP, THE LAST UNTIL THE MIDTERMS IN 2026.")
CORE_WEEKLY_MSG = ("A new WEEKLY ROV-WIDE Sincere summary report has been sent to the CORE GROUP. "
                        "Click to open.")

######################
# MONTHLY messages
######################
# CORE_MONTHLY_MSG = (f"A new MONTHLY ROV-WIDE Sincere summary report has been sent to the CORE GROUP, the last monthly for this campaign cycle. "
#                   f"Click to open.")
CORE_MONTHLY_MSG = (f"A new MONTHLY ROV-WIDE Sincere summary report has been sent to the CORE GROUP. "
                  f"Click to open.")

ORG_MONTHLY_MSG = ("A new MONTHLY Sincere summary report is available for your room. "
                  "To access the sheet you will need to be logged in to Google.  Do this by using the Chrome browser or by going to google.com in another browser."
                  "Click to open.")
# ORG_MONTHLY_MSG = ("A new MONTHLY Sincere summary report is available for your room, the last monthly for this campaign cycle.. "
#                   "To access the sheet you will need to be logged in to Google.  Do this by using the Chrome browser or by going to google.com in another browser."
#                   "Click to open.")

# CORE_EMAIL_LIST = ['kramsman@yahoo.com',
#                    'rovkatyhickman@gmail.com',
#                    'Andrea@centerforcommonground.org',
#                    'dee@centerforcommonground.org',
#                    'comstockrov@gmail.com',
#                    'nancy@centerforcommonground.org',
#                    'bill.becky.rov@gmail.com',
#                    'carey@harmonicsystems.net',
#                    'gideon.asher1@gmail.com',
#                    'gabriel@centerforcommonground.org',
#                    'josi@centerforcommonground.org',
#                    ]
CORE_EMAIL_LIST = ['kramsman@yahoo.com']
# CORE_EMAIL_LIST = ['kramsman@yahoo.com', 'gideon.asher1@gmail.com']
# CORE_EMAIL_LIST = ['gideon.asher1@gmail.com']
# CORE_EMAIL_LIST = ['test@test.com', 'bkramer@kramericore.com']
# CORE_EMAIL_LIST = ['kramsman+test@Gmail.com']


# Google cloud console for authorizations: https://console.cloud.google.com/home/dashboard?project=voterletters-reports,
# then 'APi and Services' (upper left)>credentials.  +Create > Oauth > desktop
# delete token.json / credentials.json.  Rename to credentials, copy in

# determine if application is running as a script file or frozen exe
if getattr(sys, 'frozen', False):
    # assume .exe is moved from subdirectory ./dist so reports go to same
    # ROOT_PATH = os.path.dirname(sys.executable)
    ROOT_PATH = Path(sys.executable).parents[0]
elif __file__:
    # ROOT_PATH = os.path.dirname(__file__)
    ROOT_PATH = Path(__file__).parents[0]
else:
    ROOT_PATH = None
# ROOT_PATH = str(ROOT_PATH)

# string to filter factories.  Can not filter for 'not locked' because that would exclude some current year closed
# campaigns
FACTORY_FILTER_STRING = '-2026'

# for google drive api
SEND_PERMISSION_EMAIL_FLAG = True  # send permission granted emails
SCOPES = ['https://www.googleapis.com/auth/drive']

# ROOM_REPORT_FOLDER_ID = "1OEAdTnhQoAKpyzFqYdM1ztr42vHAqbk5"  # Folder holding Sincere Organizer Reports
ROOM_REPORT_FOLDER_ID = "1BcvXzwyEKGiiWSt27NLaassLE0DMgioP"  # 'VoterLetters Organizer Reports' copied to Tech@CFCG
# ADMIN_REPORT_FOLDER_ID = '1_vTXsMYOwK_TLYaRqC7YA_jolBuqDWP4'
# ADMIN_REPORT_FOLDER_ID = '1uILHx5lllvFqYjpzzYI-9JtTivHRitdx'  # test move folder from CFTech to Tech@CFCG
# ADMIN_REPORT_FOLDER_ID = '0AOkQpyahstHzUk9PVA'  # FAILED - would not give permission - share drive on Tech@CFCG
# ADMIN_REPORT_FOLDER_ID = '1hyArqMhrgCJeRoN7tHR8Rmkzn4BvCmpq'  # WORKS - Tech@CFCG my drive folder (NOT share)
ADMIN_REPORT_FOLDER_ID = '1PU8hcYfE3Vlh5v8Cq60Gup_8lfaFff4b'  # 'VoterLetters Admin Reports' copied to Tech@CFCG

SINCERE_DOWNLOAD_DIR = "~/Downloads/"
OUTPUT_DIR_ADMIN = "~/Dropbox/Postcard Files/VL Admin Reports"
OUTPUT_DIR_REPORTS = "~/Dropbox/Postcard Files/VL Org Reports"

state_list = ["AL", "AK", "AS", "AZ", "AR", "CA", "CO", "CT", "DE", "DC", "FL", "GA", "GU", "HI", "ID", "IL", "IN",
              "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM",
              "NY", "NC", "ND", "MP", "OH", "OK", "OR", "PA", "PR", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA",
              "VI", "WA", "WV", "WI", "WY"]


noteLines = [
    '',
    'WHAT IS IN THIS REPORT',
    '',
    'What you\u2019ll get:',
    '   - Each spreadsheet presents the address and request counts in your room: by time-period,'
    'state, campaign, team, and writer.',
    '   - The tabs (located at the bottom of the sheet) provide an overview then gets more specific:',
    '       the first few tabs show data for your whole room.  Tabs after that are one for each team.',
    '   - A weekly version showing counts by week will be produced each Monday. Monthly time periods',
    '       will be shown in a report produced on the first day of each month.',
    '   - You will still get the daily summary reports from Sincere.',
    '',
    'The overall report:',
    '   - The sheet shows only the address and request counts for your room.',
    '   - Each tab shows totals over time, either in weekly or monthly increments.  Dates are endings.',
    '   - The titles give the',
    '         - name of your room',
    '         - name of the team',
    '         - time increment (weekly or monthly)',
    '         - dates included in the report',
    '   - The data will be summarized in each row',
    '   - The reports show total state, campaign, teams, and writers',
    '',
    'The individual tabs of the report are:',
    '   - Tab 1, named  "Notes" explains the report'
    '   - Tab 2, named  "Campaigns" shows requested addresses broken by Campaign',
    '   - Tab 3, "All Team Sums" by Teams',
    '   - Tab 4, "Teams w Writers" by Writers within Teams',
    '   - Tab 5, "Teams w Counts" number of requests by Writers within Teams. ',
    '   - The tabs following are specific to each team (including individual writers).',
    '       These can be copied and pasted for team leads.  '
]
