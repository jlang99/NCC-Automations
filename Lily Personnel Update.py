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
test_sheet_ID = '1Uq8LSya5w6xiAbFhqeCa1upT2Zs-ueQmYFP6cxNkOkI'

def dbcnxn():
    global db, connect_db, c
    #Connect to DB
    db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\Shared drives\Narenco Projects\O&M Projects\NCC\NCC\NCC 039.accdb;'
    connect_db = pyodbc.connect(db)
    c = connect_db.cursor()

    
def update_Sheet():
    dbcnxn()
    credentials = None
    if os.path.exists("XelioUpdate-token.json"):
        credentials = Credentials.from_authorized_user_file("XelioUpdate-token.json", SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(r"C:\Users\OMOPS.AzureAD\Python Codes\NCC-AutomationCredentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("XelioUpdate-token.json", "w+") as token:
            token.write(credentials.to_json())
    
    c.execute('SELECT * from [XELIOupdate]')
    data = c.fetchall()

    if data == []:
        exit()
    # Function to format datetime to date or time
    def format_datetime(dt, fmt):
        if fmt == 'date':
            return dt.date()
        elif fmt == 'time':
            return dt.time()

    # Dictionary to hold processed data
    processed_data = {}

    for entry in data:
        log_id = entry[0]
        description = entry[8]
            
        # Format dates and times
        formatted_entry = (
            log_id,
            entry[1], #Company
            entry[2], #Name
            format_datetime(entry[3], 'date'),
            format_datetime(entry[4], 'time'),
            format_datetime(entry[5], 'date'),
            format_datetime(entry[6], 'time'),
            entry[7], #Total Work hours Float value
            description
        )
        dictkey = str(log_id) + formatted_entry[2]
        # Merge descriptions for the same log_id
        if dictkey in processed_data and processed_data[dictkey][2] == entry[2]:
            old_entry = processed_data[dictkey]
            new_description = old_entry[8] + '; ' + description
            new_entry = old_entry[:8] + (new_description,) 
            processed_data[dictkey] = new_entry
        else:
            processed_data[dictkey] = formatted_entry

    # Convert dictionary back to list of tuples
    final_data = list(processed_data.values())
    #ic(final_data)

    def get_column_by_day(date):
        day_mapping = {
            0: 'D',  # Monday
            1: 'E',  # Tuesday
            2: 'F',  # Wednesday
            3: 'G',  # Thursday
            4: 'H',  # Friday
            5: 'I',  # Saturday
            6: 'J'   # Sunday
        }
        day_of_week = date.weekday()
        return day_mapping.get(day_of_week)

    # Prepare data for sheet
    new_final_data = []
    for entry in final_data:
        # entry format: (log_id, col1, col2, date, time1, date2, time2, col7, description)
        day_column = get_column_by_day(entry[3])  # entry[3] is the date for column C
        time_concat = entry[4].strftime('%H:%M') + "-" + entry[6].strftime('%H:%M')
        row = [
            entry[1],  # Column A
            entry[2],  # Column B
            entry[3].strftime('%m/%d/%Y'),  # Column C
            '',  # Column based on day of week
            '', '', '', '', '', '', # Fill columns until K
            entry[7],  # Column K
            '',  # Column L
            entry[8],  # Column M
            'Yes'  # Column N
        ]
        # Insert the time_concat into the correct day column
        day_column_index = ord(day_column) - ord('A')  # Convert column letter to index
        row[day_column_index] = time_concat
        new_final_data.append(row)

    #ic(new_final_data)
    #WRITE TO SHEET
    try:
        service = build("sheets", "v4", credentials= credentials)
        sheets = service.spreadsheets()
    
        body = {"values": new_final_data}
        
        sheets.values().append(
            spreadsheetId=test_sheet_ID, 
            range="X-Elio",  # Specify the sheet and range where data should be appended
            valueInputOption="USER_ENTERED", 
            body=body
        ).execute()
        
    except HttpError as error:
        print(error)
    print("Sent!")
    time.sleep(1)
    print("Closing...")
    time.sleep(0.5)
    

print("Updating X-Elio's Lily Sign in Spreadsheet")
update_Sheet()


