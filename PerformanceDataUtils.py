import re
import time as ty
import numpy as np
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import datetime as dt # Already imported
from bs4 import BeautifulSoup
from tkinter import messagebox
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os, sys
#my Package
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(parent_dir)
from PythonTools import EMAILS, CREDS, get_google_credentials

JOSEPH_SHEET = "1iuBuPvvJn_H9_PBVGQn9FPQ3jixOiJc01Hd22vlvncw"
JACOB_SHEET = "19BQEQy6mIH9pvh2qc0K7WDOuZCt0o715FVNMVd52xPw"
DEFAULT_SHEET = "1CVFRuRCBqtzNq-3J3DSjoKrxR5TXSPpPYLFRV5ANls4"  
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# TODO: Populate these lists with the site names for each person.
JOSEPH_SITES = {"Lily", "Bluebird", "Bulloch 1A", "Bulloch 1B", "CDIA", "Cougar", "Harding", "Harrison", "Holly Swamp", "JEFFERSON", "LongLeaf Pine Solar",
                "Marshall", "Mclean", "Sunflower", "Thunderhead", "Upson", "Violet", "Wayne II", "Wayne III", "Wellons Farm", "Whitehall", "Whitetail"}
JACOB_SITES = {"Lily", "Hayes", "Hickory", "BISHOPVILLE", "Cardinal", "Cherry Blossom", "Conetoe 1", "Duplin", "Elk",  "Freight Line", "Gray Fox", 
               "HICKSON", "OGBURN", "PG Solar", "Richmond Cadle", "Shorthorn", "Tedder", "Van Buren", "Warbler", "Washington", "Wayne I", "WILLIAMS"}

EXTRA_SITES = {"Charter GM", "Shoe Show", "Omnidian Target", "Pivot Energy"}


def update_issue_tracking_google_sheet(service, spreadsheet_id, sheet_name, dataframe=False):
    """Updates a Google Sheet with the given dataframe."""
    print(sheet_name)
    # Get spreadsheet metadata to check for existing sheets
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata.get('sheets', '')
    
    sheet_id = None
    for s in sheets:
        if s['properties']['title'] == sheet_name:
            sheet_id = s['properties']['sheetId']
            break

    if sheet_id is not None:
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
        response = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
        sheet_id = response['replies'][0]['addSheet']['properties']['sheetId']
    if dataframe is not False:
        dataframe = dataframe.sort_values(by='WO Date', ascending=False)
        # Write dataframe to sheet
        values = [dataframe.columns.values.tolist()] + dataframe.values.tolist()
        body = {'values': values}
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_name}'!B1",
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()

        ty.sleep(1) #Added to prevent Rate Limits

        if len(values) > 1:
            requests = [
                # 1. Request to set the header in cell A1 (Row index 0, Column index 0)
                {
                    "updateCells": {
                        "rows": [
                            {
                                "values": [
                                    {
                                        "userEnteredValue": {
                                            "stringValue": "Known?"
                                        }
                                    }
                                ]
                            }
                        ],
                        "fields": "userEnteredValue",
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 0,
                            "endRowIndex": 1, # Row 1 is exclusive, so this covers only Row 0 (A1)
                            "startColumnIndex": 0,
                            "endColumnIndex": 1  # Column 1 is exclusive, so this covers only Column 0 (A)
                        }
                    }
                },
                # 2. Request to apply the checkbox data validation to rows A2 onwards
                {
                    "setDataValidation": {
                        "range": {
                            "sheetId": sheet_id,
                            # Start at row index 1 (A2) to skip the new header in A1
                            "startRowIndex": 1,
                            "endRowIndex": len(values),
                            "startColumnIndex": 0,
                            "endColumnIndex": 1
                        },
                        "rule": {
                            "condition": {"type": "BOOLEAN"},
                            "showCustomUi": True
                        }
                    }
                }
            ]
            
            # Execute the batch update with both requests
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'requests': requests}
            ).execute()
    

    # To avoid hitting rate limits
    ty.sleep(2)


def process_WO_issue_tracking_file(file_path, creds):
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

    service = build('sheets', 'v4', credentials=creds)

    grouped = df.groupby('Site')

    processed_sites = set()

    for site_name, site_df in grouped:
        processed_sites.add(site_name) # Add the site to the set of processed sites
        if site_name in JOSEPH_SITES: # Check if the site belongs to Joseph
            update_issue_tracking_google_sheet(service, JOSEPH_SHEET, site_name, site_df) # Update Joseph's sheet
        if site_name in JACOB_SITES: # Check if the site belongs to Jacob
            update_issue_tracking_google_sheet(service, JACOB_SHEET, site_name, site_df) # Update Jacob's sheet
        if site_name not in JOSEPH_SITES and site_name not in JACOB_SITES: # If the site doesn't belong to either Joseph or Jacob
            update_issue_tracking_google_sheet(service, DEFAULT_SHEET, site_name, site_df) # Update the default sheet

    # Clear sheets for any of Joseph's + Jacob's sites that were not in the report
    for site_name in JOSEPH_SITES:
        if site_name not in processed_sites:
            update_issue_tracking_google_sheet(service, JOSEPH_SHEET, site_name)
    for site_name in JACOB_SITES:
        if site_name not in processed_sites:
            update_issue_tracking_google_sheet(service, JACOB_SHEET, site_name)
    for site_name in EXTRA_SITES:
        if site_name not in processed_sites:
            update_issue_tracking_google_sheet(service, DEFAULT_SHEET, site_name)




INV_PERFORMANCE_SHEET = '1sp8SJNbMf0AhtEn2-WZmCUKHCGZrqRDgNFNOtoUkuHE'
INVERTER_GROUPS = {
    'Bishopville II': {
        '15 String': {"3-6", "3-7"},
        '14 String': {"1-1", "1-2", "1-3", "1-5", "2-2", "2-3", "2-5", "2-7", "2-8", "2-9", "4-4", "4-6", "4-8", "4-9"},
        '13 String': {"1-4", "1-6", "1-7", "1-8", "1-9", "2-1", "2-4", "2-6", "3-1", "3-2", "3-3", "3-4", "3-5", "3-8", "3-9", "4-1", "4-2", "4-3", "4-5", "4-7"},
    },
    'Bulloch 1A': {
        '11 String': {7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24},
        '10 String': {1, 2, 3, 4, 5, 6},
    },
    'Bulloch 1B': {
        '11 String': {2, 3, 4, 5, 6, 7, 8, 13, 14, 15, 16, 18, 19, 20, 21, 22, 23, 24},
        '10 String': {1, 9, 10, 11, 12, 17},
    },
    'Cardinal': {
        '94 Limit': {15, 16, 17, 18, 19, 20, 21, 36, 37, 38, 39, 40, 41, 42, 54, 55, 56, 57, 58, 59},
        '95 Limit': {8, 9, 10, 11, 12, 13, 14, 29, 30, 31, 32, 33, 34, 35, 48, 49, 50, 51, 52, 53},
        '96 Limit': {1, 2, 3, 4, 5, 6, 7, 22, 23, 24, 25, 26, 27, 28, 43, 44, 45, 46, 47},
    },
    'Cougar Solar, LLC': {
        '14 String': {"1-1", "1-2", "1-3", "1-4", "1-5", "2-1", "2-2", "2-3", "2-4", "2-5", "2-6", "3-1", "3-2", "3-3", "3-4", "3-5", "4-1", "4-2", "4-3", "4-4", "4-5", "5-1", "5-4", "5-5", "6-1", "6-2", "6-3", "6-4", "6-5"},
        '13 String': {"5-2", "5-3"},
    },
    'Elk': {
        '11 String': {1, 2, 3, 4, 5, 6, 7, 8, 14, 15, 16, 17, 18, 19, 20, 23, 24, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43},
        '10 String': {9, 10, 11, 12, 13, 21, 22, 25, 26, 27, 28, 29},
    },
    'Freight Line': {
        '19 String': {1, 4, 5, 8, 9, 10, 11, 12, 15, 16, 17, 18},
        '18 String': {3},
        '18 String Limited 66%': {2},
        '10 String': {7, 13},
        '9 String': {6, 14},
    },
    'Gray Fox': {
        '12 String': {"1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10", "1.11", "2.2", "2.3", "2.4", "2.5", "2.6", "2.7", "2.8", "2.9", "2.10", "2.11"},
        '11 String': {"1.12", "1.13", "1.14", "1.15", "1.16", "1.17", "1.18", "1.19", "1.20", "2.1", "2.12", "2.13", "2.14", "2.15", "2.16", "2.17", "2.18", "2.19", "2.20"},
    },
    'Harding': {
        '13 String': {4, 5, 6, 10, 11, 12, 13, 14, 15, 17, 18, 19},
        '12 String': {1, 2, 3, 7, 8, 9, 16, 20, 21, 22, 23, 24},
    },
    'Harrison': {
        '17 String': {2, 3, 4, 5, 6, 7, 12, 13, 14, 15, 18, 19, 20, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 43},
        '16 String': {1, 8, 9, 10, 11, 16, 17, 21, 40, 41, 42},
    },
    'Hayes': {
        '16 String': {5, 6, 7, 8, 9, 18, 19, 20, 21, 22},
        '15 String': {1, 2, 3, 4, 10, 11, 12, 13, 14, 15, 16, 17, 23, 24, 25, 26},
    },
    'Hickson': {
        '13 String': {"1-7", "1-8", "1-9", "1-12", "1-13", "1-14", "1-15", "1-16"},
        '12 String': {"1-1", "1-2", "1-3", "1-4", "1-5", "1-6", "1-7", "1-8", "1-9", "1-10", "1-11"},
    },
    'Jefferson': {
        '18 String': {"A1.1", "A1.2", "A1.3", "A1.4", "A1.6", "A1.13", "A1.14", "A2.1", "A2.2", "A2.4", "A2.5", "A2.6", "A2.7", "A2.10", "A2.11", "A2.12", "A2.13", "A2.14", "A2.15", "A2.16", "A3.8", "A3.9", "A3.10", "A3.11", "A3.12", "A3.13", "A3.14", "A3.15", "A4.1", "A4.2", "A4.3", "A4.4", "A4.5", "A4.6", "A4.7", "A4.8"},
        '17 String': {"A1.5", "A1.7", "A1.8", "A1.9", "A1.10", "A1.11", "A1.12", "A1.15", "A1.16", "A2.3", "A2.8", "A2.9", "A3.1", "A3.2", "A3.3", "A3.4", "A3.5", "A3.6", "A3.7", "A3.16", "A4.9", "A4.10", "A4.11", "A4.12", "A4.13", "A4.14", "A4.15", "A4.16"},
    },
    'Longleaf Pine': {
        '12 String': {"A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "B1", "B19", "B20"},
        '11 String': {"A12", "A13", "A14", "A15", "A16", "A17", "A18", "A19", "A20", "B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "B17", "B18"},
    },
    'McLean': {
        '11 String': {3, 4, 5, 6, 7, 8, 9, 10, 13, 14, 15, 16, 21, 24, 25, 26, 27, 29, 31},
        '10 String': {1, 2, 11, 12, 17, 18, 19, 20, 22, 23, 28, 30, 32, 33, 34, 35, 36, 37, 38, 39, 40},
    },
    'PG': {
        'Limited .66': {1, 2, 3, 4, 5, 6},
        '100': {7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18},
    },
    'Shorthorn': {
        '13 String': {6, 16, 21, 25, 29, 30, 31, 34, 35, 36, 44, 49, 50, 51, 54, 55, 56, 57, 67, 68, 69, 70, 71, 72},
        '12 String': {1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 12, 13, 14, 15, 17, 18, 19, 20, 22, 23, 24, 26, 27, 28, 32, 33, 37, 38, 39, 40, 41, 42, 43, 45, 46, 47, 48, 52, 53, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66},
    },
    'Sunflower': {
        '12 String': {1, 2, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 35, 36, 37, 38, 39, 40, 41, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 71, 78, 79, 80},
        '13 String': {3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 34, 42, 62, 63, 64, 65, 66, 67, 68, 69, 70, 73, 74, 75, 76, 77},
    },
    'Richmond': {
        '11 String': {1, 2, 3, 4, 5, 6, 7, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21},
        '10 String': {8, 9, 10, 22, 23, 24},
    },
    'Tedder': {
        '16 String': {5, 6, 7, 10, 11, 12, 13, 14},
        '15 String': {1, 2, 3, 4, 8, 15, 16},
    },
    'Upson': {
        '11 String': {1, 2, 3, 4, 5, 9, 10, 11, 12, 13, 14, 15, 16, 17, 21, 22, 23, 24},
        '10 String': {6, 7, 8, 18, 19, 20},
    },
    'Washington': {
        '13 String': {4, 5, 6, 7, 8, 9, 10, 11, 12, 15, 16, 17, 18, 19, 21, 22, 23, 24, 40},
        '12 String': {1, 2, 3, 13, 14, 20, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39},
    },
    'Whitehall': {
        '13 String': {2, 6, 7, 8, 9, 10, 11, 12},
        '36 String': {1, 3, 4, 5, 13, 14, 15, 16},
    },
    'Whitetail': {
        '16 String': {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 21, 22, 23, 24, 25, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 49, 50, 51, 57, 61, 62, 63, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80},
        '17 String': {13, 14, 15, 16, 17, 18, 19, 20, 26, 27, 28, 29, 30, 31, 43, 44, 45, 46, 47, 48, 52, 53, 54, 55, 56, 58, 59, 60, 64},
    },
    'Williams Solar, LLC': {
        '11 String': {"A1", "A2", "A3", "A5", "A7", "A12", "A13", "A14", "A15", "A16", "A17", "A18", "A19", "A20", "B1", "B2", "B3", "B4", "B5", "B6", "B7", "B11", "B12", "B13", "B14", "B15", "B16", "B17", "B18", "B19"},
        '10 String': {"A4", "A6", "A8", "A9", "A10", "A11", "B8", "B9", "B10", "B20"},
    }
}

SITE_DATA = {
    'Bishopville II Solar':{
        'GHI': 'Hukseflux SR30 - POA (PY0) [Sun] (GHI)',
        'POA': 'Hukseflux SR30 - POA (PY0) [Sun] (POA)',
        'inv_num': 36,
    }, 
    'Bluebird Solar':{
        'GHI': 'Kipp & Zonen SMP-6 (GHI) (GHI)',
        'POA': 'Kipp & Zonen SMP-6 (POA) (POA)',
        'inv_num': 24,
    },
    'Bulloch 1A':{
        'GHI': 'Kipp & Zonen SMP-12 (GHI) - PY2 (GHI)',
        'POA': 'Kipp & Zonen SMP-12 (POA) - PY0 (RMA14237) (POA)',
        'inv_num': 24,
    },
    'Bulloch 1B':{
        'GHI': 'Kipp & Zonen SMP-12 (GHI) - PY2 (GHI)',
        'POA': 'Kipp & Zonen SMP-12 (POA) - PY0 (POA)',
        'inv_num': 24,
    },
    'Cardinal':{
        'GHI': 'Hukseflux SR30 GHI (RMA11424) (GHI)',
        'POA': 'Hukesflux SR30 POA (POA)',
        'inv_num': 59,
    },
    'Cherry Blossom Solar, LLC':{
        'GHI': 'Weather Station 3 (Dual Mod Temp) (GHI)',
        'POA': 'Weather Station 3 (Dual Mod Temp) (POA)',
        'inv_num': 4,
    },
    'Conetoe':{
        'GHI': None,
        'POA': None,
        'inv_num': 16,
    },
    'Cougar Solar, LLC':{
        'GHI': 'Reference Cell (POA, GHI)',
        'POA': 'KZPOA (POA, GHI)',
        'inv_num': 31,
    },
    'Duplin':{
        'GHI': 'Weather Station (POA GHI) (GHI)',
        'POA': 'Weather Station (POA GHI) (POA)',
        'inv_num': 21,
    },
    'Elk Solar':{
        'GHI': 'Weather Station (GHI)',
        'POA': 'Weather Station (POA)',
        'inv_num': 43,
    },
    'Freight Line Solar':{
        'GHI': 'Kipp & Zonen SMP6 - GHI (GHI)',
        'POA': 'Kipp & Zonen SMP6 - POA (GHI)',
        'inv_num': 18,
    },
    'Gray Fox Solar':{
        'GHI': 'PYRANOMETER - GHI [Sun2] (GHI)',    
        'POA': 'PYRANOMETER - GHI [Sun2] (POA)',
        'inv_num': 40,
    },
    'Harding Solar':{
        'GHI': 'PYRANOMETER - GHI (GHI)',
        'POA': 'PYRANOMETER - POA (POA)',
        'inv_num': 24,
    },
    'Harrison Solar':{
        'GHI': 'Hukseflux SR-30 GHI (GHI)',
        'POA': 'Huksaflux SR-30 POA UP (POA)',
        'inv_num': 43,
    },
    'Hayes':{
        'GHI': 'Hukseflux SR30 (GHI) (POA, GHI)',
        'POA': 'Albedometer POA (POA Sensor)',
        'inv_num': 26,
    },
    'Hickory Solar, LLC':{
        'GHI': 'Weather Station Modbus 17 (Hukseflux GHI w/ Mod Temp)# (GHI)',
        'POA': 'Weather Station Modbus 17 (Hukseflux GHI w/ Mod Temp)# (POA)',
        'inv_num': 2,
    },
    'Hickson Solar Farm':{
        'GHI': 'Hukseflux SR05 (GHI) (GHI)',
        'POA': 'Hukseflux SR05 (POA) (POA)',
        'inv_num': 16,
    },
    'Holly Swamp Solar':{
        'GHI': 'Kipp & Zonen SMP6 - GHI (GHI)',
        'POA': 'Kipp & Zonen SMP6 - POA (POA)',
        'inv_num': 16,
    },
    'Jefferson Solar':{
        'GHI': 'Hukseflux SR05 - GHI (GHI)',
        'POA': 'Hukseflux SR05 - GHI (POA)',
        'inv_num': 64,
    },
    'Longleaf Pine Solar, LLC':{
        'GHI': 'PYRANOMETER - GHI (GHI)',
        'POA': 'PYRANOMETER - POA (POA)',
        'inv_num': 40,
    },
    'Marshall Solar':{
        'GHI': 'Hukseflux SR05 GHI (GHI)',
        'POA': 'Hukseflux SR05 POA (POA)',
        'inv_num': 16,
    },
    'McLean':{
        'GHI': 'PYRANOMETER - GHI (PY0) (GHI)',
        'POA': 'PYRANOMETER - POA (PY1) (SO84054) (POA)',
        'inv_num': 40,
    },
    'Ogburn Solar Farm':{
        'GHI': 'Hukseflux SR05 (GHI) (GHI)',
        'POA': 'Hukseflux SR05 (POA) (POA)',
        'inv_num': 16,
    },
    'PG Solar':{
        'GHI': 'Kipp & Zonen SMP6 - GHI (POA, GHI)',
        'POA': 'Kipp & Zonen SMP6 - POA (POA, GHI)',
        'inv_num': 18,
    },
    'Richmond':{
        'GHI': 'Kipp & Zonen SMP-12 (GHI) - PY1 (GHI)',
        'POA': 'Kipp & Zonen SMP-12 (POA) - PY0 (POA)',
        'inv_num': 24,
    },  
    'Shorthorn':{
        'GHI': 'PYRANOMETER - GHI (GHI)',
        'POA': 'PYRANOMETER - POA (POA)',
        'inv_num': 72,
    },
    'Sunflower Solar':{
        'GHI': 'PYRANOMETER - GHI (PY0) (GHI)',
        'POA': 'PYRANOMETER - POA (PY1) (POA)',
        'inv_num': 80,
    },
    'Tedder Solar':{
        'GHI': 'Kipp & Zonen SMP-12 (GHI) - PY1 (GHI)',
        'POA': 'Kipp & Zonen SMP-12 (POA) - PY0 (POA)',
        'inv_num': 16,
    },  
    'Thunderhead Solar':{
        'GHI': 'Hukseflux SR05 (GHI) (GHI)',
        'POA': 'Hukseflux SR05 (GHI) (POA)',
        'inv_num': 16,
    },  
    'Upson':{
        'GHI': 'Kipp & Zonen SMP-12 (GHI) - PY1 (GHI)',
        'POA': 'Kipp & Zonen SMP-12 (POA) - PY0 (POA)',
        'inv_num': 24,
    },  
    'Van Buren Solar':{
        'GHI': 'Hukseflux SR30 - GHI (SO38592) (GHI)',
        'POA': 'Hukseflux SR30 - POA (SO38592) (POA)',
        'inv_num': 17,
    },  
    'Violet Solar, LLC':{
        'GHI': 'KZ POA (GHI)',
        'POA': 'KZ POA (POA)',
        'inv_num': 2,
    },
    'Warbler':{
        'GHI': 'Hukseflux SR30 GHI (GHI)',
        'POA': 'Hukseflux SR30 POA (POA)',
        'inv_num': 32,
    },
    'Washington Solar':{
        'GHI': 'PYRANOMETER - GHI (GHI)',
        'POA': 'PYRANOMETER - POA (POA)',
        'inv_num': 40,  
    },
    'Wayne 1':{
        'GHI': 'Weather Station (POA GHI) (GHI)',
        'POA': 'Weather Station (POA GHI) (POA)',
        'inv_num': 4,
    }, 
    'Wayne 2':{
        'GHI': 'Weather Station (POA GHI) (GHI)',
        'POA': 'Weather Station (POA GHI) (POA)',
        'inv_num': 4,
    },
    'Wayne 3':{
        'GHI': 'Weather Station (POA GHI) (GHI)',
        'POA': 'Weather Station (POA GHI) (POA)',
        'inv_num': 4,
    },
    'Wellons Solar, LLC':{
        'GHI': 'Weather Station GHI & POA (GHI)',
        'POA': 'Weather Station POA (POA)',
        'inv_num': 6,
    }, 
    'Whitehall Solar':{
        'GHI': 'PYRANOMETER - GHI (GHI)',
        'POA': 'PYRANOMETER - POA (POA)',
        'inv_num': 16,
    }, 
    'Whitetail':{
        'GHI': 'Hukseflux SR-30 - GHI (GHI)',
        'POA': 'Hukseflux SR-30 - POA (POA)',
        'inv_num': 80,
    }, 
    'Williams Solar, LLC':{
        'GHI': 'PYRANOMETER - GHI (GHI)',
        'POA': 'PYRANOMETER - POA (POA)',
        'inv_num': 40,
    }, 
}

CB_SHEET = '1RGUwARwDdfDoC8VcNQgb5KC6gPOA-rcMaPsu6vlzg9o'
CB_GROUPS = {
    'Conetoe': {
        '24 String': {
            '1.1', '1.2', '1.3', '1.4', '1.5', '1.6', '1.7', '1.8', '1.9', '1.10',
            '1.11', '1.12', '2.1', '2.2', '2.3', '2.4', '2.5', '2.6', '2.7', '2.8',
            '2.9', '2.10', '2.11', '2.12', '3.1', '3.2', '3.3', '3.4', '3.5', '3.6',
            '3.7', '3.8', '3.9', '3.10', '3.11', '3.12', '4.1', '4.2', '4.3', '4.4',
            '4.5', '4.6', '4.7', '4.8', '4.9', '4.10', '4.11', '4.12'
        },
        '20 String': {'1.13'}
    },
    'Hickory Solar, LLC': {
        '15 String': {
            'INV1A1 A', 'INV1A2 A', 'INV1A3 A', 'INV1A4 A', 'INV1A5 A', 'INV1A6 A',
            'INV1A7 A', 'INV1A8 A', 'INV1A9 A', 'INV1B13 B', 'INV1B14 B', 'INV1B15 B',
            'INV1B16 B', 'INV1B17 B', 'INV1B18 B', 'INV1B19 B', 'INV1B20 B', 'INV1B21 B',
            'INV2A1 A', 'INV2A2 A', 'INV2A6 A', 'INV2A7 A', 'INV2A8 A', 'INV2A9 A',
            'INV2A10 A', 'INV2A11 A', 'INV2A12 A', 'INV2B13 B', 'INV2B16 B', 'INV2B17 B',
            'INV2B23 B'
        },
        '12 String': {'INV1A10 A', 'INV1B23 B', 'INV1B24 B', 'INV2A3 A', 'INV2A4 A'},
        '18 String': {'INV1A11 A', 'INV2B22 B'},
        '20 String': {'INV1A12 A'},
        '11 String': {'INV2A5 A'},
        '16 String': {'INV2B14 B'},
        '17 String': {'INV2B15 B', 'INV2B19 B'},
        '14 String': {'INV2B20 B', 'INV2B24 B'}
    },
    'Wellons Solar, LLC': {
        '22 String': {
            'CB1-1', 'CB1-2', 'CB1-3', 'CB1-4', 'CB1-5', 'CB1-6', 'CB1-9', 'CB1-10',
            'CB1-11', 'CB1-12', 'CB1-13', 'CB1-14', 'CB1-15', 'CB1-16', 'CB2-1', 'CB2-2',
            'CB2-3', 'CB2-4', 'CB2-5', 'CB2-6', 'CB2-7'
        },
        '18 String': {
            'CB1-8', 'CB2-10', 'CB2-12', 'CB2-13', 'CB2-14', 'CB2-15', 'CB2-16', 'CB2-17'
        },
        '20 String': {
            'CB1-17', 'CB1-18', 'CB2-8', 'CB2-9', 'CB2-11', 'CB2-18'
        },
        '28 String': {
            'CB3-1', 'CB3-2', 'CB3-3', 'CB3-4', 'CB3-5', 'CB3-6', 'CB3-7', 'CB3-8',
            'CB3-9', 'CB3-10', 'CB3-11', 'CB3-12', 'CB3-13', 'CB3-14'
        }
    },
    'Violet Solar, LLC': {
        '18 String': {
            'dc_11', 'dc_13', 'dc_14', 'dc_15', 'dc_16', 'dc_17', 'dc_18', 'dc_19',
            'dc_110', 'dc_111', 'dc_112', 'dc_22', 'dc_23', 'dc_24', 'dc_25', 'dc_26',
            'dc_27', 'dc_212', 'dc_213', 'dc_41', 'dc_42', 'dc_43', 'dc_44', 'dc_45',
            'dc_46', 'dc_47'
        },
        '16 String': {'dc_21', 'dc_28', 'dc_29', 'dc_210', 'dc_37', 'dc_410', 'dc_48', 'dc_311', 'dc_312'},
        '14 String': {'dc_38', 'dc_39', 'dc_310', 'dc_12'},
        '20 String': {'dc_32', 'dc_33', 'dc_34'},
        '28 String': {'dc_35', 'dc_36'},
        'NA': {'dc_211', 'dc_411', 'dc_412'} # These seem to be un-grouped
    }
}




def natural_sort_key(s):
    """
    Creates a sort key that prioritizes grouping by the first number found,
    then sorts naturally.
    """
    # Normalize the string to handle variations like 'Inv 3' vs 'Inv3' and 'A_CB' vs 'ACB'
    normalized_s = str(s).replace(' ', '').replace('_', '')
    parts = re.findall(r'\d+|\D+', normalized_s)
    
    # Convert digit strings to integers for correct numerical sorting
    natural_key = [int(part) if part.isdigit() else part.lower() for part in parts]
    
    # Find the first number in the key
    first_number = -1
    for item in natural_key:
        if isinstance(item, int):
            first_number = item
            break
            
    # Return a tuple: (first number, full natural key).
    return (first_number, natural_key)

def update_cb_sheet(metrics, credentials):
    """Writes the collected metrics to a Google Sheet."""
    service = build('sheets', 'v4', credentials=credentials)

    # Helper to format for JSON
    def format_val(v):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return None
        if isinstance(v, (np.integer, np.int64)):
            return int(v)
        if isinstance(v, (np.floating, np.float64)):
            return float(v)
        return v

    for sheet_name, site_metrics in metrics.items():
        if not site_metrics:
            print(f"No metrics to write for {sheet_name}. Skipping.")
            continue

        avg_peak_timestamp = site_metrics.get('avg_peak_timestamp')
        abs_peak_timestamp = site_metrics.get('abs_peak_timestamp')
        # --- Prepare data for writing ---
        # Get a unique, sorted list of all CB names for this site
        all_cb_names = sorted(list(set(
            # The order of replacements is important here to avoid partial string matching.
            # '_abs_peak_loss' must be replaced before '_peak_loss' to ensure the full suffix is removed correctly.
            key.replace('cb_', '').replace('_sum_loss', '').replace('_abs_peak_loss', '').replace('_peak_loss', '').replace('_sum', '')
            for key in site_metrics.keys() if key.startswith('cb_')
        )))

        cb_data_for_sorting = []
        for cb_name in all_cb_names:
            sum_loss = site_metrics.get(f"cb_{cb_name}_sum_loss", None)
            peak_loss = site_metrics.get(f"cb_{cb_name}_peak_loss", None)
            abs_peak_loss = site_metrics.get(f"cb_{cb_name}_abs_peak_loss", None)
            cb_sum = site_metrics.get(f"cb_{cb_name}_sum", None)
            
            cb_data_for_sorting.append({
                'cb_name': cb_name,
                'cb_sum': cb_sum,
                'sum_loss': sum_loss,
                'peak_loss': peak_loss,
                'abs_peak_loss': abs_peak_loss
            })

        # Sort by group (first digit), then by the maximum of the three loss columns (descending), then by combiner box name (ascending)
        cb_data_for_sorting.sort(key=lambda x: (
            natural_sort_key(x['cb_name'])[1],
            -max(
                (x['sum_loss'] if x['sum_loss'] is not None and not np.isnan(x['sum_loss']) else -1),
                (x['peak_loss'] if x['peak_loss'] is not None and not np.isnan(x['peak_loss']) else -1),
                (x['abs_peak_loss'] if x['abs_peak_loss'] is not None and not np.isnan(x['abs_peak_loss']) else -1)
            ),
            natural_sort_key(x['cb_name'])
        ))

        data_to_write = []
        for cb_entry in cb_data_for_sorting:
            data_to_write.append([
                cb_entry['cb_name'],
                format_val(cb_entry['cb_sum']),
                format_val(cb_entry['sum_loss']),
                format_val(cb_entry['peak_loss']),
                format_val(cb_entry['abs_peak_loss'])
            ])
        headers = [["Combiner Box", "5-day Sum (kWh)", "Sum-Based Loss (5-day)", "Avg Peak-Based Loss", "Abs Peak-Based Loss"]]
        final_data = headers + data_to_write

        # --- Write to Google Sheet ---
        try:
            # Get spreadsheet metadata to find sheetId and current rowCount
            spreadsheet = service.spreadsheets().get(spreadsheetId=CB_SHEET).execute()
            sheet_id = None
            current_row_count = 0
            for s in spreadsheet.get('sheets', []):
                if s.get('properties', {}).get('title') == sheet_name:
                    sheet_id = s['properties']['sheetId']
                    current_row_count = s['properties']['gridProperties']['rowCount']
                    break
            
            if sheet_id is None:
                print(f"Sheet '{sheet_name}' not found. Cannot update.")
                continue

            # Calculate required rows (data + 1 for timestamp)
            required_rows = len(final_data) + 1 # +1 for the timestamp row

            # If required rows exceed current row count, append new rows
            if required_rows > current_row_count:
                rows_to_add = required_rows - current_row_count
                append_rows_request = {'appendDimension': {'sheetId': sheet_id, 'dimension': 'ROWS', 'length': rows_to_add}}
                service.spreadsheets().batchUpdate(spreadsheetId=CB_SHEET, body={'requests': [append_rows_request]}).execute()
                print(f"Appended {rows_to_add} rows to sheet '{sheet_name}'.")
                ty.sleep(1)

            # Clear the sheet before writing new data
            clear_range = f"{sheet_name}!A1:E" # Clear up to the required rows
            service.spreadsheets().values().clear(spreadsheetId=CB_SHEET, range=clear_range).execute()
            ty.sleep(1) # API rate limit

            # Write all data at once
            range_to_update = f'{sheet_name}!A1'
            body = {'values': final_data}
            service.spreadsheets().values().update(
                spreadsheetId=CB_SHEET,
                range=range_to_update,
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            print(f"Successfully updated all ratios for {sheet_name}.")
            ty.sleep(1)
            
            # Now, add the timestamps at the end of the respective columns
            avg_peak_timestamp_str = avg_peak_timestamp.strftime('%Y-%m-%d %H:%M') if pd.notna(avg_peak_timestamp) else ''
            if avg_peak_timestamp_str:
                timestamp_cell_d = f'{sheet_name}!D{len(final_data) + 1}'
                timestamp_body_d = {'values': [[avg_peak_timestamp_str]]}
                service.spreadsheets().values().update(
                    spreadsheetId=CB_SHEET,
                    range=timestamp_cell_d,
                    valueInputOption='USER_ENTERED',
                    body=timestamp_body_d
                ).execute()
                print(f"Successfully added average peak timestamp for {sheet_name} at {timestamp_cell_d}.")
                ty.sleep(1)

            abs_peak_timestamp_str = abs_peak_timestamp.strftime('%Y-%m-%d %H:%M') if pd.notna(abs_peak_timestamp) else ''
            if abs_peak_timestamp_str:
                timestamp_cell_e = f'{sheet_name}!E{len(final_data) + 1}'
                timestamp_body_e = {'values': [[abs_peak_timestamp_str]]}
                service.spreadsheets().values().update(
                    spreadsheetId=CB_SHEET,
                    range=timestamp_cell_e,
                    valueInputOption='USER_ENTERED',
                    body=timestamp_body_e
                ).execute()
                print(f"Successfully added absolute peak timestamp for {sheet_name} at {timestamp_cell_e}.")
                ty.sleep(1)

        except HttpError as err:
            print(f"An error occurred while writing to sheet {sheet_name}: {err}")
    email_cherryCB_report(credentials)

def process_cb_file(file_path, credentials):
    xls = pd.ExcelFile(file_path)
    all_site_metrics = {}

    for sheet_name in xls.sheet_names:
        if sheet_name == "Sheet 1":
            continue

        print(f"Processing sheet: {sheet_name}")
        
        # 1. Read and prepare dataframes
        df = pd.read_excel(xls, sheet_name=sheet_name, skiprows=2)
        df = df.iloc[1:].reset_index(drop=True)
        
        timestamp_col = df.columns[0]
        df[timestamp_col] = pd.to_datetime(df[timestamp_col], errors='coerce')
        df.dropna(subset=[timestamp_col], inplace=True)

        # Filter for peak production hours (9am to 5pm)
        peak_dataframe = df[(df[timestamp_col].dt.hour >= 10) & (df[timestamp_col].dt.hour <= 15)].copy()
        if peak_dataframe.empty:
            print(f"No data for peak hours in {sheet_name}. Skipping.")
            continue

        # Create sum_dataframe for the last 5 days
        max_date = peak_dataframe[timestamp_col].max()
        last_5_days_start_date = max_date - pd.Timedelta(days=5)
        sum_dataframe = df[df[timestamp_col] >= last_5_days_start_date].copy()

        # Identify all combiner box columns
        cb_columns = [col for col in df.columns[1:] if 'dc_' in col or 'CB' in col or '.' in col or 'INV' in col]
        
        # Convert all CB columns to numeric
        for col in cb_columns:
            peak_dataframe[col] = pd.to_numeric(peak_dataframe[col], errors='coerce')
            sum_dataframe[col] = pd.to_numeric(sum_dataframe[col], errors='coerce')

        # 2. Calculate the three data points for all CBs
        cb_sums = {col: sum_dataframe[col].sum() for col in cb_columns}

        peak_dataframe['CB_Row_Average'] = peak_dataframe[cb_columns].mean(axis=1)
        max_avg_index = peak_dataframe['CB_Row_Average'].idxmax()
        cb_peaks = {}
        if pd.notna(max_avg_index):
            max_production_row = peak_dataframe.loc[max_avg_index]
            cb_peaks = {col: max_production_row[col] for col in cb_columns}
            avg_peak_timestamp = peak_dataframe.loc[max_avg_index, timestamp_col]
        else:
            avg_peak_timestamp = None

        max_value_in_df = peak_dataframe[cb_columns].max().max()
        cb_abs_peaks = {}
        if pd.notna(max_value_in_df):
            max_col_name = peak_dataframe[cb_columns].max().idxmax()
            max_value_row_index = peak_dataframe[max_col_name].idxmax()
            best_row_by_abs_max = peak_dataframe.loc[max_value_row_index]
            cb_abs_peaks = {col: best_row_by_abs_max[col] for col in cb_columns}
            abs_peak_timestamp = peak_dataframe.loc[max_value_row_index, timestamp_col]
        else:
            abs_peak_timestamp = None

        # 3. Process ratios based on groups
        site_metrics = {}
        site_metrics['avg_peak_timestamp'] = avg_peak_timestamp
        site_metrics['abs_peak_timestamp'] = abs_peak_timestamp

        for cb_name in cb_columns:
            site_metrics[f"cb_{cb_name}_sum"] = cb_sums.get(cb_name, 0)

        groups_to_process = {}

        if sheet_name in CB_GROUPS:
            groups_to_process = CB_GROUPS[sheet_name]
        elif sheet_name == "Cherry Blossom Solar, LLC":
            temp_groups = {}
            cb_pattern = r'ST(\d+)'
            for col in cb_columns:
                match = re.search(cb_pattern, col)
                if match:
                    group_id = f"Inverter {match.group(1)}"
                    if group_id not in temp_groups:
                        temp_groups[group_id] = set()
                    temp_groups[group_id].add(col)
            groups_to_process = temp_groups
        else:
            print(f"Warning: No grouping defined for {sheet_name}. Skipping ratio calculation.")

        for group_name, cb_set in groups_to_process.items():
            if group_name == 'NA': continue

            group_sums = [cb_sums.get(cb_name, 0) for cb_name in cb_set]
            max_sum = max(group_sums) if group_sums else 0
            for cb_name in cb_set: # Calculate loss instead of ratio
                loss = (max_sum - cb_sums.get(cb_name, 0)) if max_sum > 0 else 0
                site_metrics[f"cb_{cb_name}_sum_loss"] = loss # Change key to _sum_loss

            group_peaks = [cb_peaks.get(cb_name, 0) for cb_name in cb_set]
            max_peak = max(group_peaks) if group_peaks else 0
            for cb_name in cb_set: # Calculate loss instead of ratio
                loss = (max_peak - cb_peaks.get(cb_name, 0)) if max_peak > 0 else 0
                site_metrics[f"cb_{cb_name}_peak_loss"] = loss # Change key to _peak_loss

            group_abs_peaks = [cb_abs_peaks.get(cb_name, 0) for cb_name in cb_set]
            max_abs_peak = max(group_abs_peaks) if group_abs_peaks else 0
            for cb_name in cb_set: # Calculate loss instead of ratio
                loss = (max_abs_peak - cb_abs_peaks.get(cb_name, 0)) if max_abs_peak > 0 else 0
                site_metrics[f"cb_{cb_name}_abs_peak_loss"] = loss # Change key to _abs_peak_loss

        all_site_metrics[sheet_name] = site_metrics

    # 4. Write all results to Google Sheet
    update_cb_sheet(all_site_metrics, credentials)













def natural_sort_key(s):
    """Sorts strings with numbers in a natural, human-friendly order."""
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]



def update_performance_sheet(metrics, credentials):
    """Writes the collected metrics to a Google Sheet."""
    service = build('sheets', 'v4', credentials=credentials)

    for site_name, site_metrics in metrics.items():
        # Extract and sort inverter ratios using natural sort
        inverter_keys = []
        for key, value in site_metrics.items():
            if key.startswith('inverter_') and key.endswith('_ratio') and '_peak_' not in key:
                inv_id_str = key.replace('inverter_', '').replace('_ratio', '')
                inverter_keys.append((inv_id_str, value))
        
        # Sort by inverter ID using the natural sort key
        inverter_keys.sort(key=lambda x: natural_sort_key(x[0]))
        # Replace NaN with None for JSON compatibility
        ratios = [[None if isinstance(value, float) and np.isnan(value) else value] for _, value in inverter_keys]
        #print(f"Prepared performance ratios for {site_name}: {ratios}")
        if ratios:
            range_end = 1 + len(ratios) # B2 to B(num_inverters + 1)
            range_to_update = f'{site_name}!B2:B{range_end}'
            body = {'values': ratios}
            
            service.spreadsheets().values().update(
                spreadsheetId=INV_PERFORMANCE_SHEET,
                range=range_to_update,
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            print(f"Successfully updated performance ratios for {site_name} in range {range_to_update}.")
            ty.sleep(1)  # Pause to avoid hitting API rate limits
        # --- Ratios based on AVERAGE (for column C) ---
        avg_peak_timestamp = site_metrics.get('avg_peak_timestamp')
        avg_peak_timestamp_str = avg_peak_timestamp.strftime('%Y-%m-%d %H:%M') if pd.notna(avg_peak_timestamp) else ''

        avg_inverter_keys = []
        for key, value in site_metrics.items():
            if key.startswith('inverter_') and key.endswith('_peak_ratio') and '_abs_' not in key:
                inv_id_str = key.replace('inverter_', '').replace('_peak_ratio', '')
                avg_inverter_keys.append((inv_id_str, value))
        
        avg_inverter_keys.sort(key=lambda x: natural_sort_key(x[0]))
        # Replace NaN with None for JSON compatibility
        avg_ratios = [[None if isinstance(value, float) and np.isnan(value) else value] for _, value in avg_inverter_keys]
        #print(f"Prepared average performance ratios for {site_name}: {avg_ratios}")
        if avg_ratios:
            range_end = 1 + len(avg_ratios)
            range_to_update = f'{site_name}!C2:C{range_end}'
            body = {'values': avg_ratios}
            service.spreadsheets().values().update(
                spreadsheetId=INV_PERFORMANCE_SHEET,
                range=range_to_update,
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            print(f"Successfully updated average performance ratios for {site_name} in range {range_to_update}.")
            ty.sleep(1)
            # Now, add the timestamp at the end of the column
            timestamp_cell = f'{site_name}!C{max(SITE_DATA[site_name]["inv_num"] + 2, range_end + 1)}'
            timestamp_body = {'values': [[avg_peak_timestamp_str]]}
            service.spreadsheets().values().update(
                spreadsheetId=INV_PERFORMANCE_SHEET,
                range=timestamp_cell,
                valueInputOption='USER_ENTERED',
                body=timestamp_body
            ).execute()
            print(f"Successfully added timestamp for average peak ratios at {timestamp_cell}.")
            ty.sleep(1)

        # --- Ratios based on ABSOLUTE PEAK (for column D) ---
        abs_peak_timestamp = site_metrics.get('abs_peak_timestamp')
        abs_peak_timestamp_str = abs_peak_timestamp.strftime('%Y-%m-%d %H:%M') if pd.notna(abs_peak_timestamp) else ''

        # --- Ratios based on ABSOLUTE PEAK (for column D) ---
        abs_peak_inverter_keys = []
        for key, value in site_metrics.items():
            if key.startswith('inverter_') and key.endswith('_abs_peak_ratio'):
                inv_id_str = key.replace('inverter_', '').replace('_abs_peak_ratio', '')
                abs_peak_inverter_keys.append((inv_id_str, value))
        
        abs_peak_inverter_keys.sort(key=lambda x: natural_sort_key(x[0]))
        abs_peak_ratios = [[None if isinstance(value, float) and np.isnan(value) else value] for _, value in abs_peak_inverter_keys]
        
        if abs_peak_ratios:
            range_end = 1 + len(abs_peak_ratios)
            range_to_update = f'{site_name}!D2:D{range_end}'
            body = {'values': abs_peak_ratios}
            service.spreadsheets().values().update(
                spreadsheetId=INV_PERFORMANCE_SHEET,
                range=range_to_update,
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            print(f"Successfully updated absolute peak performance ratios for {site_name} in range {range_to_update}.")
            ty.sleep(1)

            # Now, add the timestamp at the end of the column
            timestamp_cell = f'{site_name}!D{max(SITE_DATA[site_name]["inv_num"] + 2, range_end + 1)}'
            timestamp_body = {'values': [[abs_peak_timestamp_str]]}
            service.spreadsheets().values().update(
                spreadsheetId=INV_PERFORMANCE_SHEET,
                range=timestamp_cell,
                valueInputOption='USER_ENTERED',
                body=timestamp_body
            ).execute()
            print(f"Successfully added timestamp for absolute peak ratios at {timestamp_cell}.")
            ty.sleep(1)




def find_inverter_num(column_name):
    """
    Extracts the inverter number/identifier from a column name string using regex.
    It looks for patterns like 'Inverter 1-2', 'INV 2.4', 'Inverter - 11', or 'A1.1'
    """
    # This regex looks for "Inverter" or "INV" (case-insensitive) OR a pattern like "A1.1" at the start,
    # followed by an optional hyphen, and then captures an identifier that can include
    # letters, numbers, hyphens, and periods.
    match = re.search(r'(?i)(?:(?:inverter|inv)\s*-?\s*([A-Za-z0-9-.]+)|(^[A-Z]\d\.\d+))', column_name)
    if match:
        # The regex has two capturing groups, one will be None
        inverter_id = match.group(1) or match.group(2)
        # The first captured group is the inverter identifier
        #print(f"Extracted inverter identifier: {inverter_id} from column name: {column_name}")
        return inverter_id
    return None # Return None if no match is found

def process_INV_performance_xlsx(file_path, credentials):
    """Process the selected Excel file."""
    performance_metrics = {}
    xls = pd.ExcelFile(file_path)
    for sheet_name in xls.sheet_names:
        #print("Loop Start, Sheet Name: ", sheet_name)

        # Skip sheets that are not relevant for performance checks.
        # Added Wayne sites as they have very few inverters and can skew averages.
        if sheet_name in {"Sheet 1", "Charlotte Airport", "Wayne 1", "Wayne 2", "Wayne 3"}:
            continue  # Skip sheets that we can't determine perfomance metircs based on average due to minimal data. Small Inv Groups            
        dataframe = pd.read_excel(xls, sheet_name=sheet_name, skiprows=2)
        #print(f"Created DataFrame for sheet: {sheet_name}")
        timestamp_col = dataframe.columns[0]


        # Convert Timestamps to Datetime Objs to run filter on time of day.
        dataframe[timestamp_col] = pd.to_datetime(dataframe[timestamp_col], errors='coerce')

        # New filtering logic
        # Prioritize POA, then GHI, then fall back to time.
        poa = SITE_DATA.get(sheet_name, {}).get('POA')
        ghi = SITE_DATA.get(sheet_name, {}).get('GHI')

        # Initialize conditions to False
        poa_condition = pd.Series([False] * len(dataframe), index=dataframe.index)
        ghi_condition = pd.Series([False] * len(dataframe), index=dataframe.index)

        if poa and poa in dataframe.columns:
            poa_condition = pd.to_numeric(dataframe[poa], errors='coerce') >= 400
        if ghi and ghi in dataframe.columns:
            ghi_condition = pd.to_numeric(dataframe[ghi], errors='coerce') >= 300

        time_condition = (dataframe[timestamp_col].dt.hour >= 10) & (dataframe[timestamp_col].dt.hour <= 15)

        # 1. Identify inverter data columns based on SITE_DATA
        inv_num = SITE_DATA.get(sheet_name, {}).get('inv_num', 0)
        inverter_data_columns = dataframe.columns[1:inv_num+1]

        for col in inverter_data_columns:
            dataframe[col] = pd.to_numeric(dataframe[col], errors='coerce')

        #print(sheet_name," | ", dataframe.columns[1:inv_num+1])
        # Time is hard and fast.
        use_poa = False
        if poa and poa in dataframe.columns:
            if (time_condition & poa_condition).any():
                use_poa = True

        use_ghi = False
        if not use_poa and ghi and ghi in dataframe.columns:
            if (time_condition & ghi_condition).any():
                use_ghi = True

        if use_poa:
            peak_dataframe = dataframe[time_condition & poa_condition].copy()
        elif use_ghi:
            peak_dataframe = dataframe[time_condition & ghi_condition].copy()
        else:
            peak_dataframe = dataframe[time_condition].copy()
        
        #print(sheet_name," | ", peak_dataframe.columns[1:inv_num+1])

        if peak_dataframe.empty:
            print(f"No data found for {sheet_name} after filtering. Skipping.")
            continue

        sum_dataframe = peak_dataframe.copy()
        max_date = sum_dataframe[timestamp_col].max()
        #print(f"Max date in sheet '{sheet_name}': {max_date}")
        last_week_start_date = max_date - dt.timedelta(days=5)
        #print(f"Last week start date: {last_week_start_date}")
        
        # Apply the 5-day filter to the sum_dataframe
        sum_dataframe = sum_dataframe[sum_dataframe[timestamp_col] >= last_week_start_date]

        #print(sum_dataframe.head())
        # 2. Check if the remaining columns exist or have data
        if inverter_data_columns.empty:
            print(f"{inverter_data_columns}")
            # This covers the case where the sheet only had the timestamp column and nothing else.
            print(f"Skipping sheet '{sheet_name}' as no inv data found.")
            continue
        # Calculate inverter averages and sums

        # 2. Convert these columns to numeric, coercing errors to NaN
        # This is crucial for accurate averaging

        # 3. Calculate the row-wise average across ONLY the inverter columns
        # axis=1 specifies a row-wise operation
        peak_dataframe['Inverter_Row_Average'] = peak_dataframe[inverter_data_columns].mean(axis=1)
        
        # 3. Find the row (index) with the maximum 'Inverter_Row_Average' value
        max_avg_index = peak_dataframe['Inverter_Row_Average'].idxmax()
        if pd.isna(max_avg_index) or np.isnan(max_avg_index):
            print(f"Skipping sheet '{sheet_name}' as no valid inverter data found after filtering.")
            continue
        
        # 4. Select the row corresponding to the time of peak average production
        # .loc ensures the index is used correctly to get the single row of data
        max_production_row = peak_dataframe.loc[max_avg_index]
        avg_peak_timestamp = max_production_row[timestamp_col]
        
        # 5. Extract the production values for all inverters from that single row
        inverter_production_at_peak = max_production_row[inverter_data_columns]

        # 6. Find the maximum production value among the inverters in that specific row
        max_production_in_row = inverter_production_at_peak.max()
        inverter_peaks = {
            find_inverter_num(col): max_production_row[col]
            for col in inverter_data_columns
        }

        # --- New process: Find best row by absolute max value in any inverter column ---
        # Find the maximum value across all inverter columns in the peak_dataframe
        max_value_in_df = peak_dataframe[inverter_data_columns].max().max()

        # Find the column that contains this overall maximum value
        max_col_name = peak_dataframe[inverter_data_columns].max().idxmax()

        # Find the index of the row where this maximum value occurs
        max_value_row_index = peak_dataframe[max_col_name].idxmax()

        # Get the entire row of data for that time
        best_row_by_abs_max = peak_dataframe.loc[max_value_row_index]
        abs_peak_timestamp = best_row_by_abs_max[timestamp_col]

        # Create a dictionary of inverter production values from this new "best" row
        inverter_abs_peaks = {
            find_inverter_num(col): best_row_by_abs_max[col]
            for col in inverter_data_columns
        }

        inverter_sums = {find_inverter_num(col): pd.to_numeric(sum_dataframe[col], errors='coerce').sum() for col in sum_dataframe.columns[1:inv_num+1]}
        #if sheet_name == 'Marshall':
            #print(f"Inverter sums for {sheet_name}: {inverter_sums}")
        if sheet_name in INVERTER_GROUPS:
            performance_metrics[sheet_name] = {}
            performance_metrics[sheet_name]['avg_peak_timestamp'] = avg_peak_timestamp
            performance_metrics[sheet_name]['abs_peak_timestamp'] = abs_peak_timestamp
            for group_name, inverter_numbers_set in INVERTER_GROUPS[sheet_name].items():
                group_sums = [inverter_sums.get(str(inv_num), 0) for inv_num in inverter_numbers_set]
                max_sum = max(group_sums) if group_sums else 0
                performance_metrics[sheet_name][group_name] = max_sum if max_sum else 0
                for inv_num in inverter_numbers_set:
                    inv_sum = inverter_sums.get(str(inv_num), 0)
                    ratio = (inv_sum / max_sum) if max_sum > 0 else 0
                    performance_metrics[sheet_name][f"inverter_{inv_num}_ratio"] = ratio
                # --- Peak Ratio Calculation (NEW) ---
                # Get the peak production for each inverter in the group
                group_peaks = [inverter_peaks.get(str(inv_num), 0) for inv_num in inverter_numbers_set]
                # Filter out NaNs before calculating max
                valid_peaks = [p for p in group_peaks if pd.notna(p)]
                max_peak = max(valid_peaks) if valid_peaks else 0

                # Store the max peak for the group
                performance_metrics[sheet_name][f"{group_name}_MaxPeak"] = max_peak

                # Calculate ratio for each inverter in the group based on the group's max peak
                for inv_num in inverter_numbers_set:
                    inv_peak = inverter_peaks.get(str(inv_num), 0)
                    if pd.isna(inv_peak):
                        peak_ratio = 0
                    else:
                        peak_ratio = (inv_peak / max_peak) if max_peak > 0 else 0
                    performance_metrics[sheet_name][f"inverter_{inv_num}_peak_ratio"] = peak_ratio

                # --- Absolute Peak Ratio Calculation (NEW) ---
                # Get the absolute peak production for each inverter in the group
                group_abs_peaks = [inverter_abs_peaks.get(str(inv_num), 0) for inv_num in inverter_numbers_set]
                valid_abs_peaks = [p for p in group_abs_peaks if pd.notna(p)]
                max_abs_peak = max(valid_abs_peaks) if valid_abs_peaks else 0

                # Store the max absolute peak for the group
                performance_metrics[sheet_name][f"{group_name}_MaxAbsPeak"] = max_abs_peak

                # Calculate ratio for each inverter in the group based on the group's max absolute peak
                for inv_num in inverter_numbers_set:
                    inv_abs_peak = inverter_abs_peaks.get(str(inv_num), 0)
                    if pd.isna(inv_abs_peak):
                        abs_peak_ratio = 0
                    else:
                        abs_peak_ratio = (inv_abs_peak / max_abs_peak) if max_abs_peak > 0 else 0
                    performance_metrics[sheet_name][f"inverter_{inv_num}_abs_peak_ratio"] = abs_peak_ratio
        elif sheet_name == "Duplin":
            performance_metrics[sheet_name] = {}
            performance_metrics[sheet_name]['avg_peak_timestamp'] = avg_peak_timestamp
            performance_metrics[sheet_name]['abs_peak_timestamp'] = abs_peak_timestamp
            central_inverters = {}
            central_inverter_peaks = {}
            central_counter = 1
            string_inverters = {}
            string_inverter_peaks = {}
            string_counter = 1
            central_inverter_abs_peaks = {}
            string_inverter_abs_peaks = {}

            for col in dataframe.columns[1:inv_num+1]:
                inv_sum = pd.to_numeric(sum_dataframe[col], errors='coerce').sum()
                inv_peak = max_production_row[col]
                inv_abs_peak = best_row_by_abs_max[col]
                if 'central' in col.lower():
                    central_inverters[central_counter] = inv_sum
                    central_inverter_peaks[central_counter] = inv_peak
                    central_inverter_abs_peaks[central_counter] = inv_abs_peak
                    #print(f"{col} Counter: {central_counter}, Sum: {inv_sum}")
                    central_counter += 1
                elif 'string' in col.lower():
                    string_inverters[string_counter] = inv_sum
                    string_inverter_peaks[string_counter] = inv_peak
                    string_inverter_abs_peaks[string_counter] = inv_abs_peak
                    #print(f"{col} Counter: {string_counter}, Sum: {inv_sum}")
                    string_counter += 1

            for group_name, group_data in [("Central Inverters", central_inverters), ("String Inverters", string_inverters)]:
                if group_data:
                    max_sum = max(group_data.values())
                    performance_metrics[sheet_name][group_name] = max_sum if max_sum else 0
                    for inv_num, inv_sum in group_data.items():
                        ratio = (inv_sum / max_sum) if max_sum > 0 else 0
                        # Adjust inverter number for string inverters to be unique
                        if group_name == "String Inverters":
                            inv_num += len(central_inverters)
                        performance_metrics[sheet_name][f"inverter_{inv_num}_ratio"] = ratio if ratio else 0
                # --- Peak Ratio Calculation (NEW) ---
                group_peaks = central_inverter_peaks if group_name == "Central Inverters" else string_inverter_peaks
                if group_peaks:
                    valid_peaks = [p for p in group_peaks.values() if pd.notna(p)]
                    max_peak = max(valid_peaks) if valid_peaks else 0
                    performance_metrics[sheet_name][f"{group_name}_MaxPeak"] = max_peak

                    for inv_num, inv_peak in group_peaks.items():
                        if pd.isna(inv_peak):
                            peak_ratio = 0
                        else:
                            peak_ratio = (inv_peak / max_peak) if max_peak > 0 else 0
                        # Adjust inverter number for string inverters to be unique
                        if group_name == "String Inverters":
                            inv_num += len(central_inverter_peaks)
                        performance_metrics[sheet_name][f"inverter_{inv_num}_peak_ratio"] = peak_ratio if peak_ratio else 0
            # --- Absolute Peak Ratio Calculation (NEW) ---
            for group_name, group_abs_peaks in [("Central Inverters", central_inverter_abs_peaks), ("String Inverters", string_inverter_abs_peaks)]:
                if group_abs_peaks:
                    valid_abs_peaks = [p for p in group_abs_peaks.values() if pd.notna(p)]
                    max_abs_peak = max(valid_abs_peaks) if valid_abs_peaks else 0
                    performance_metrics[sheet_name][f"{group_name}_MaxAbsPeak"] = max_abs_peak

                    for inv_num, inv_abs_peak in group_abs_peaks.items():
                        if pd.isna(inv_abs_peak):
                            abs_peak_ratio = 0
                        else:
                            abs_peak_ratio = (inv_abs_peak / max_abs_peak) if max_abs_peak > 0 else 0
                        # Adjust inverter number for string inverters to be unique
                        if group_name == "String Inverters":
                            inv_num += len(central_inverter_abs_peaks)
                        performance_metrics[sheet_name][f"inverter_{inv_num}_abs_peak_ratio"] = abs_peak_ratio if abs_peak_ratio else 0
            #print(performance_metrics[sheet_name])

        elif sheet_name == "Conetoe":
            performance_metrics[sheet_name] = {}
            performance_metrics[sheet_name]['avg_peak_timestamp'] = avg_peak_timestamp
            performance_metrics[sheet_name]['abs_peak_timestamp'] = abs_peak_timestamp
            # For Conetoe, sum parts like '1.1', '1.2' into a main inverter '1'
            main_inverter_sums = {}
            for inv_part, inv_sum in inverter_sums.items():
                if inv_part and '.' in inv_part:
                    main_inv_num = inv_part.split('.')[0]
                    if main_inv_num not in main_inverter_sums:
                        main_inverter_sums[main_inv_num] = 0
                    main_inverter_sums[main_inv_num] += inv_sum

            if main_inverter_sums:
                # Now calculate ratios based on the summed values of main inverters
                all_main_sums = list(main_inverter_sums.values())
                max_sum = max(all_main_sums) if all_main_sums else 0
                performance_metrics[sheet_name]["All Inverters"] = max_sum

                for main_inv_num, total_sum in main_inverter_sums.items():
                    ratio = (total_sum / max_sum) if max_sum > 0 else 0
                    performance_metrics[sheet_name][f"inverter_{main_inv_num}_ratio"] = ratio
            
            # --- Peak Ratio Calculation (NEW) ---
            main_inverter_peaks = {}
            for inv_part, inv_peak in inverter_peaks.items():
                if inv_part and '.' in inv_part:
                    main_inv_num = inv_part.split('.')[0]
                    if main_inv_num not in main_inverter_peaks:
                        main_inverter_peaks[main_inv_num] = 0
                    if pd.notna(inv_peak):
                        main_inverter_peaks[main_inv_num] += inv_peak # Summing the peak production for parts

            if main_inverter_peaks:
                all_main_peaks = list(main_inverter_peaks.values())
                max_peak = max(all_main_peaks) if all_main_peaks else 0
                performance_metrics[sheet_name]["All Inverters_MaxPeak"] = max_peak

                for main_inv_num, total_peak in main_inverter_peaks.items():
                    peak_ratio = (total_peak / max_peak) if max_peak > 0 else 0
                    performance_metrics[sheet_name][f"inverter_{main_inv_num}_peak_ratio"] = peak_ratio
            # --- Absolute Peak Ratio Calculation (NEW) ---
            main_inverter_abs_peaks = {}
            for inv_part, inv_abs_peak in inverter_abs_peaks.items():
                if inv_part and '.' in inv_part:
                    main_inv_num = inv_part.split('.')[0]
                    if main_inv_num not in main_inverter_abs_peaks:
                        main_inverter_abs_peaks[main_inv_num] = 0
                    if pd.notna(inv_abs_peak):
                        main_inverter_abs_peaks[main_inv_num] += inv_abs_peak # Summing the peak production for parts

            if main_inverter_abs_peaks:
                all_main_abs_peaks = list(main_inverter_abs_peaks.values())
                max_abs_peak = max(all_main_abs_peaks) if all_main_abs_peaks else 0
                performance_metrics[sheet_name]["All Inverters_MaxAbsPeak"] = max_abs_peak
                for main_inv_num, total_abs_peak in main_inverter_abs_peaks.items():
                    abs_peak_ratio = (total_abs_peak / max_abs_peak) if max_abs_peak > 0 else 0
                    performance_metrics[sheet_name][f"inverter_{main_inv_num}_abs_peak_ratio"] = abs_peak_ratio
            #print(performance_metrics[sheet_name])
        # This block should only run for sites NOT in INVERTER_GROUPS.
        else:
            
            print(f"DEBUG: Site {sheet_name} falling back to 'All Inverters' group. Check INVERTER_GROUPS key.")
            # If the site is not in INVERTER_GROUPS and not an exception, treat all inverters as one group.
            performance_metrics[sheet_name] = {}
            performance_metrics[sheet_name]['avg_peak_timestamp'] = avg_peak_timestamp
            performance_metrics[sheet_name]['abs_peak_timestamp'] = abs_peak_timestamp
            all_sums = list(inverter_sums.values())
            max_sum = max(all_sums) if all_sums else 0
            performance_metrics[sheet_name]["All Inverters"] = max_sum
            #if sheet_name == 'Marshall':
            #    print(f"{sheet_name} - All Inverters max sum: {max_sum}\n{all_sums}")

            for inv_num, inv_sum in inverter_sums.items():
                ratio = (inv_sum / max_sum) if max_sum > 0 else 0
                performance_metrics[sheet_name][f"inverter_{inv_num}_ratio"] = ratio
                #print(f"{sheet_name} - Inverter {inv_num} sum: {inv_sum}, max:{max_sum}, ratio: {ratio}\n {len(list(performance_metrics[sheet_name].items()))}")

            # --- Peak Ratio Calculation (NEW) ---
            all_peaks = list(inverter_peaks.values())
            valid_peaks = [p for p in all_peaks if pd.notna(p)]
            max_peak = max(valid_peaks) if valid_peaks else 0
            performance_metrics[sheet_name]["All Inverters_MaxPeak"] = max_peak

            for inv_num, inv_peak in inverter_peaks.items():
                if pd.isna(inv_peak):
                    peak_ratio = 0
                else:
                    peak_ratio = (inv_peak / max_peak) if max_peak > 0 else 0
                performance_metrics[sheet_name][f"inverter_{inv_num}_peak_ratio"] = peak_ratio

            # --- Absolute Peak Ratio Calculation (NEW) ---
            all_abs_peaks = list(inverter_abs_peaks.values())
            valid_abs_peaks = [p for p in all_abs_peaks if pd.notna(p)]
            max_abs_peak = max(valid_abs_peaks) if valid_abs_peaks else 0
            performance_metrics[sheet_name]["All Inverters_MaxAbsPeak"] = max_abs_peak

            for inv_num, inv_abs_peak in inverter_abs_peaks.items():
                if pd.isna(inv_abs_peak):
                    abs_peak_ratio = 0
                else:
                    abs_peak_ratio = (inv_abs_peak / max_abs_peak) if max_abs_peak > 0 else 0
                performance_metrics[sheet_name][f"inverter_{inv_num}_abs_peak_ratio"] = abs_peak_ratio

    print("Loop End")
    #loop completes then calls this 
    update_performance_sheet(performance_metrics, credentials)

def create_list_of_sheet_names(service, spreadsheet_id):
    """Create a list of all sheet names in the spreadsheet."""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        return [sheet['properties']['title'] for sheet in spreadsheet.get('sheets', [])]
    except HttpError as err:
        print(f"An error occurred: {err}")
        return []

def get_performance_sites_list(creds):
    service = build('sheets', 'v4', credentials=creds)
    return create_list_of_sheet_names(service, INV_PERFORMANCE_SHEET)
    
def email_cherryCB_report(creds):
    # 1. Check if today is Monday (0) or Tuesday (1)
    current_weekday = dt.datetime.now().weekday()
    if current_weekday not in [0, 1]: # Monday or Tuesday
        print("Cherry Blossom CB report email skipped: Not Monday or Tuesday.")
        return

    print("Attempting to send Cherry Blossom CB report email...")

    sheets_service = build('sheets', 'v4', credentials=creds)
    
    sheet_title = "Cherry Blossom Solar, LLC" # The specific sheet to email
    spreadsheet_id = CB_SHEET # Defined globally in PerformanceDataUtils.py

    try:
        # Get the sheet ID for "Cherry Blossom"
        spreadsheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_id = None
        for s in spreadsheet_metadata.get('sheets', []):
            if s.get('properties', {}).get('title') == sheet_title:
                sheet_id = s.get('properties', {}).get('sheetId')
                break

        if sheet_id is None:
            print(f"Sheet '{sheet_title}' not found in spreadsheet '{spreadsheet_id}'. Cannot generate PDF.")
            return

        # Construct the URL to download the sheet as a PDF
        headers = {'Authorization': 'Bearer ' + creds.token} # The user wants to export only columns A:G
        pdf_url = f'https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=pdf&gid={sheet_id}&range=A:G'
        
        print(f"Downloading PDF for '{sheet_title}' from '{pdf_url}'...")
        response = requests.get(pdf_url, headers=headers)
        response.raise_for_status() # Raise an exception for bad status codes
        pdf_content = response.content
        print("PDF downloaded successfully.")

        # Email details
        sender_email = EMAILS['NCC Desk']
        recipients = [EMAILS['Newman Segars'], EMAILS['Parker Wilson'], EMAILS['Jacob Budd'], EMAILS['Joseph Lang']] # Or a specific list for Cherry Blossom
        smtp_password = CREDS['shiftsumEmail']
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = ', '.join(recipients)
        msg['Subject'] = f"Cherry Blossom CB Report - {dt.datetime.now().strftime('%Y-%m-%d')}"

        # Static link to the Cherry Blossom CB sheet
        cherry_blossom_sheet_link = f"https://docs.google.com/spreadsheets/d/1RGUwARwDdfDoC8VcNQgb5KC6gPOA-rcMaPsu6vlzg9o/edit?gid=1637179539#gid=1637179539"
        html_body = f"<html><body><p>Hello Team,</p><p>Please find the Cherry Blossom Combiner Box report for {dt.datetime.now().strftime('%Y-%m-%d')} attached.</p><p>You can also view the live sheet here: <a href='{cherry_blossom_sheet_link}'>Cherry Blossom CB Sheet</a></p><p>Thank you,<br>NCC Automation</p></body></html>"
        msg.attach(MIMEText(html_body, 'html'))

        # Attach the PDF
        attachment = MIMEApplication(pdf_content, _subtype="pdf")
        attachment.add_header('Content-Disposition', 'attachment', filename=f"{sheet_title} CB Report {dt.datetime.now().strftime('%Y-%m-%d')}.pdf")
        msg.attach(attachment)

        # Send the email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, smtp_password)
            server.send_message(msg)
        print(f"Cherry Blossom CB report email sent successfully to {', '.join(recipients)}!")

    except requests.exceptions.RequestException as e:
        print(f"Error downloading PDF for {sheet_title}: {e}")
    except HttpError as e:
        print(f"Error accessing Google Sheets API for {sheet_title}: {e}")
    except Exception as e:
        print(f"Failed to send email: {e}")
