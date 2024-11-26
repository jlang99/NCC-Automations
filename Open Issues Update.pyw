import os
import pyodbc
import datetime
import time

from icecream import ic

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
openIssuesSheet = '1zTAHG7iNR_sdpLajTrdHObhXcL_ylMxl4ild4nbVtSI'

def dbcnxn():
    global db, connect_db, c
    #Connect to DB
    db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039.accdb;'
    connect_db = pyodbc.connect(db)
    c = connect_db.cursor()

    
def update_Sheet():
    global tracking
    dbcnxn()
    credentials = None
    if os.path.exists("OpenIssue-token.json"):
        credentials = Credentials.from_authorized_user_file("OpenIssue-token.json", SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(r"C:\Users\OMOPS.AzureAD\Python Codes\NCC-AutomationCredentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("OpenIssue-token.json", "w+") as token:
            token.write(credentials.to_json())
    
    #DATABASE Block
    c.execute('SELECT * from [Open Issue Updater]')
    data = c.fetchall()
    formatted_data = [list(row) for row in data]
    for row in formatted_data:
        for i, item in enumerate(row):
            if isinstance(item, datetime.datetime) and item:
                row[i] = item.strftime('%m/%d/%y')

    #ic(formatted_data)


    clear_body = {
        "requests": [
            {
                "updateCells": {
                    "range": {
                        "sheetId": 0,
                        "startRowIndex": 2,
                        "endRowIndex": 500
                    },
                    "fields": "userEnteredValue"
                }
            }
        ]
    }



    #WRITE TO SHEET
    try:
        service = build("sheets", "v4", credentials= credentials)
        sheets = service.spreadsheets()
    except HttpError as error:
        print("Error accessing the sheet:", error)

    try:
        sheets.batchUpdate(spreadsheetId=openIssuesSheet, body=clear_body).execute()
    except HttpError as error:
        print("Error clearing the sheet:", error)
    
    try:
        body = {"values": formatted_data}
        range = "Dump!A3:H" + str(len(formatted_data) +2)
                    
        sheets.values().update(spreadsheetId=openIssuesSheet, range=range, valueInputOption="USER_ENTERED", body=body).execute()     

    except HttpError as error:
        print("Erroring Writing to Sheet", error)
    
    tracking += 1
    print("Consecutive Successes:", tracking)
    print("Update Complete!")
    time.sleep(300)
    print("Updating...")












tracking = 0

while True:
    window = datetime.datetime.now().hour
    if window < 19 and window > 7:
        update_Sheet()
    else:
        time.sleep(3600)
