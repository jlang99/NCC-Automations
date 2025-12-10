import pandas as pd
import tkinter as tk
import sys, os
from tkinter import filedialog, messagebox
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import time as ty
from bs4 import BeautifulSoup
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(parent_dir)
from PythonTools import get_google_credentials

JOSEPH_SHEET = "1iuBuPvvJn_H9_PBVGQn9FPQ3jixOiJc01Hd22vlvncw"
JACOB_SHEET = "19BQEQy6mIH9pvh2qc0K7WDOuZCt0o715FVNMVd52xPw"
DEFAULT_SHEET = "1CVFRuRCBqtzNq-3J3DSjoKrxR5TXSPpPYLFRV5ANls4"  
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# TODO: Populate these lists with the site names for each person.
JOSEPH_SITES = {"Lily", "Bluebird", "Bulloch 1A", "Bulloch 1B", "CDIA", "Cougar", "Harding", "Harrison", "Holly Swamp", "JEFFERSON", "LongLeaf Pine Solar",
                "Marshall", "Mclean", "Sunflower", "Thunderhead", "Upson", "Violet", "Wayne II", "Wayne III", "Wellons Farm", "Whitehall", "Whitetail"}
JACOB_SITES = {"Lily", "Hayes", "Hickory", "BISHOPVILLE", "Cardinal", "Cherry Blossom", "Conetoe 1", "Duplin", "Elk",  "Freight Line", "Gray Fox", 
               "HICKSON", "OGBURN", "PG Solar", "Richmond Cadle", "Shorthorn", "Tedder", "Van Buren", "Warbler", "Washington", "Wayne I", "Williams"}

EXTRA_SITES = {"Charter GM", "Shoe Show", "Omnidian Target", "Pivot Energy"}


def update_google_sheet(service, spreadsheet_id, sheet_name, dataframe=False):
    """Updates a Google Sheet with the given dataframe."""
    print(sheet_name)
    # Get spreadsheet metadata to check for existing sheets
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')
    
    sheet_exists = any(s['properties']['title'] == sheet_name for s in sheets)

    if sheet_exists:
        # Clear existing data
        service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=sheet_name,
            body={}
        ).execute()
    else:
        # Create new sheet
        requests = [{'addSheet': {'properties': {'title': sheet_name}}}]
        body = {'requests': requests}
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
    if dataframe is not False:
        # Write dataframe to sheet
        body = {
            'values': [dataframe.columns.values.tolist()] + dataframe.values.tolist()
        }
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=sheet_name,
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()

    # To avoid hitting rate limits
    ty.sleep(2)


def process_file(file_path):
    if not file_path:
        return
    df = pd.read_excel(file_path)
    
    if 'Site' not in df.columns:
        messagebox.showerror("Error", "The Excel file must have a 'Site' column.")
        return
    
    if 'Work Description (Text Only)' in df.columns:
        df['Work Description (Text Only)'] = df['Work Description (Text Only)'].apply(
            lambda x: BeautifulSoup(x, 'html.parser').get_text(separator=' ', strip=True) if isinstance(x, str) else x
        )

    # Replace NaN with None for JSON compatibility
    df = df.where(pd.notna(df), None)

    # Convert datetime columns to strings to prevent JSON serialization errors
    for col in df.select_dtypes(include=['datetime64[ns]']).columns:
        # Format to string, preserving NaT (Not a Time) values
        df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M:%S').replace({pd.NaT: None})
    for col in df.select_dtypes(include=['object']).columns:
        if df[col].apply(lambda x: isinstance(x, pd.Timestamp)).any():
            df[col] = df[col].apply(lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if isinstance(x, pd.Timestamp) else x)

    creds = get_google_credentials()
    service = build('sheets', 'v4', credentials=creds)

    grouped = df.groupby('Site')

    processed_sites = set()

    for site_name, site_df in grouped:
        processed_sites.add(site_name)
        if site_name in JOSEPH_SITES:
            update_google_sheet(service, JOSEPH_SHEET, site_name, site_df)
        if site_name in JACOB_SITES:
            update_google_sheet(service, JACOB_SHEET, site_name, site_df)
        if site_name not in JOSEPH_SITES and site_name not in JACOB_SITES:
            update_google_sheet(service, DEFAULT_SHEET, site_name, site_df)

    # Clear sheets for any of Joseph's + Jacob's sites that were not in the report
    for site_name in JOSEPH_SITES:
        if site_name not in processed_sites:
                update_google_sheet(service, JOSEPH_SHEET, site_name)
    for site_name in JACOB_SITES:
        if site_name not in processed_sites:
            update_google_sheet(service, JACOB_SHEET, site_name)
    for site_name in EXTRA_SITES:
        if site_name not in processed_sites:
            update_google_sheet(service, DEFAULT_SHEET, site_name)


    messagebox.showinfo("Success", "Google Sheets have been updated successfully.")
    root.destroy()


def select_file():
    file_path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
    )
    if file_path:
        process_file(file_path)

root = tk.Tk()
root.title("Weekly Updates")
root.geometry("300x150")

browse_button = tk.Button(root, text="Select File and Process", command=select_file)
browse_button.pack(pady=20, fill='both', expand=True)

root.mainloop()
# End of Weekly Updates.py