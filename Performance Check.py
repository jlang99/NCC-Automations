import ctypes
import time
import pandas as pd
import sys, os
import re
import tkinter as tk
from tkinter import filedialog

#my Package
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(parent_dir)
from PythonTools import get_google_credentials

#Google Sheets API
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

performanceSheet = '1sp8SJNbMf0AhtEn2-WZmCUKHCGZrqRDgNFNOtoUkuHE'

myappid = 'NCC.Performance.Check'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)


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
    'Cougar': {
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
        '16 String': {5, 6, 7, 9, 18, 19, 20, 21},
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
    'Williams': {
        '11 String': {"A1", "A2", "A3", "A5", "A7", "A12", "A13", "A14", "A15", "A16", "A17", "A18", "A19", "A20", "B1", "B2", "B3", "B4", "B5", "B6", "B7", "B11", "B12", "B13", "B14", "B15", "B16", "B17", "B18", "B19"},
        '10 String': {"A4", "A6", "A8", "A9", "A10", "A11", "B8", "B9", "B10", "B20"},
    }
}

SITE_DATA = {
    'Bishopville II':{
        'GHI': 'Hukseflux SR30 - POA (PY0) [Sun] (GHI)',
        'POA': 'Hukseflux SR30 - POA (PY0) [Sun] (POA)',
        'inv_num': 36,
    }, 
    'Bluebird':{
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
    'Cherry Blossom':{
        'GHI': 'Weather Station 3 (Dual Mod Temp) (GHI)',
        'POA': 'Weather Station 3 (Dual Mod Temp) (POA)',
        'inv_num': 4,

    },
    'Conetoe':{
        'GHI': None,
        'POA': None,
        'inv_num': 16,
    },
    'Cougar':{
        'GHI': 'Reference Cell (POA, GHI)',
        'POA': 'KZPOA (POA, GHI)',
        'inv_num': 31,
    },
    'Duplin':{
        'GHI': 'Weather Station (POA GHI) (GHI)',
        'POA': 'Weather Station (POA GHI) (POA)',
        'inv_num': 18,
    },
    'Elk':{
        'GHI': 'Weather Station (GHI)',
        'POA': 'Weather Station (POA)',
        'inv_num': 43,
    },
    'Freight Line':{
        'GHI': 'Kipp & Zonen SMP6 - GHI (GHI)',
        'POA': 'Kipp & Zonen SMP6 - POA (GHI)',
        'inv_num': 18,
    },
    'Gray Fox':{
        'GHI': 'PYRANOMETER - GHI [Sun2] (GHI)',    
        'POA': 'PYRANOMETER - GHI [Sun2] (POA)',
        'inv_num': 40,
    },
    'Harding':{
        'GHI': 'PYRANOMETER - GHI (GHI)',
        'POA': 'PYRANOMETER - POA (POA)',
        'inv_num': 24,
    },
    'Harrison':{
        'GHI': 'Hukseflux SR-30 GHI (GHI)',
        'POA': 'Huksaflux SR-30 POA UP (POA)',
        'inv_num': 43,
    },
    'Hayes':{
        'GHI': 'Hukseflux SR30 (GHI) (POA, GHI)',
        'POA': 'Albedometer POA (POA Sensor)',
        'inv_num': 26,
    },
    'Hickory':{
        'GHI': 'Weather Station Modbus 17 (Hukseflux GHI w/ Mod Temp)# (GHI)',
        'POA': 'Weather Station Modbus 17 (Hukseflux GHI w/ Mod Temp)# (POA)',
        'inv_num': 2,
    },
    'Hickson':{
        'GHI': 'Hukseflux SR05 (GHI) (GHI)',
        'POA': 'Hukseflux SR05 (POA) (POA)',
        'inv_num': 16,
    },
    'Holly Swamp':{
        'GHI': 'Kipp & Zonen SMP6 - GHI (GHI)',
        'POA': 'Kipp & Zonen SMP6 - POA (POA)',
        'inv_num': 16,
    },
    'Jefferson':{
        'GHI': 'Hukseflux SR05 - GHI (GHI)',
        'POA': 'Hukseflux SR05 - GHI (POA)',
        'inv_num': 64,
    },
    'Longleaf Pine':{
        'GHI': 'PYRANOMETER - GHI (GHI)',
        'POA': 'PYRANOMETER - POA (POA)',
        'inv_num': 40,
    },
    'Marshall':{
        'GHI': 'Hukseflux SR05 GHI (GHI)',
        'POA': 'Hukseflux SR05 POA (POA)',
        'inv_num': 16,
    },
    'McLean':{
        'GHI': 'PYRANOMETER - GHI (PY0) (GHI)',
        'POA': 'PYRANOMETER - POA (PY1) (SO84054) (POA)',
        'inv_num': 40,
    },
    'Ogburn':{
        'GHI': 'Hukseflux SR05 (GHI) (GHI)',
        'POA': 'Hukseflux SR05 (POA) (POA)',
        'inv_num': 16,
    },
    'PG':{
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
    'Sunflower':{
        'GHI': 'PYRANOMETER - GHI (PY0) (GHI)',
        'POA': 'PYRANOMETER - POA (PY1) (POA)',
        'inv_num': 80,
    },
    'Tedder':{
        'GHI': 'Kipp & Zonen SMP-12 (GHI) - PY1 (GHI)',
        'POA': 'Kipp & Zonen SMP-12 (POA) - PY0 (POA)',
        'inv_num': 16,
    },  
    'Thunderhead':{
        'GHI': 'Hukseflux SR05 (GHI) (GHI)',
        'POA': 'Hukseflux SR05 (GHI) (POA)',
        'inv_num': 16,
    },  
    'Upson':{
        'GHI': 'Kipp & Zonen SMP-12 (GHI) - PY1 (GHI)',
        'POA': 'Kipp & Zonen SMP-12 (POA) - PY0 (POA)',
        'inv_num': 24,
    },  
    'Van Buren':{
        'GHI': 'Hukseflux SR30 - GHI (SO38592) (GHI)',
        'POA': 'Hukseflux SR30 - POA (SO38592) (POA)',
        'inv_num': 17,
    },  
    'Violet':{
        'GHI': 'KZ POA (GHI)',
        'POA': 'KZ POA (POA)',
        'inv_num': 2,
    },
    'Warbler':{
        'GHI': 'Hukseflux SR30 GHI (GHI)',
        'POA': 'Hukseflux SR30 POA (POA)',
        'inv_num': 32,
    },
    'Washington':{
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
    'Wellons':{
        'GHI': 'Weather Station GHI & POA (GHI)',
        'POA': 'Weather Station POA (POA)',
        'inv_num': 6,
    }, 
    'Whitehall':{
        'GHI': 'PYRANOMETER - GHI (GHI)',
        'POA': 'PYRANOMETER - POA (POA)',
        'inv_num': 16,
    }, 
    'Whitetail':{
        'GHI': 'Hukseflux SR-30 - GHI (GHI)',
        'POA': 'Hukseflux SR-30 - POA (POA)',
        'inv_num': 80,
    }, 
    'Williams':{
        'GHI': 'PYRANOMETER - GHI (GHI)',
        'POA': 'PYRANOMETER - POA (POA)',
        'inv_num': 40,
    }, 
}

credentials = get_google_credentials()

def natural_sort_key(s):
    """Sorts strings with numbers in a natural, human-friendly order."""
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]





def update_performance_sheet(metrics):
    """Writes the collected metrics to a Google Sheet."""
    try:
        service = build('sheets', 'v4', credentials=credentials)
        for site_name, site_metrics in metrics.items():
            # Extract and sort inverter ratios using natural sort
            inverter_keys = []
            for key, value in site_metrics.items():
                if key.startswith('inverter_') and key.endswith('_ratio'):
                    inv_id_str = key.replace('inverter_', '').replace('_ratio', '')
                    inverter_keys.append((inv_id_str, value))
            
            # Sort by inverter ID using the natural sort key
            inverter_keys.sort(key=lambda x: natural_sort_key(x[0]))
            ratios = [[value] for _, value in inverter_keys]
            if ratios:
                range_end = 1 + len(ratios) # B2 to B(num_inverters + 1)
                range_to_update = f'{site_name}!B2:B{range_end}'
                body = {'values': ratios}
                
                service.spreadsheets().values().update(
                    spreadsheetId=performanceSheet,
                    range=range_to_update,
                    valueInputOption='USER_ENTERED',
                    body=body
                ).execute()
                print(f"Successfully updated performance ratios for {site_name} in range {range_to_update}.")
                time.sleep(0.7)  # Pause to avoid hitting API rate limits
        
        root.destroy()  # Close the Tkinter root window after processing is completed

    except HttpError as err:
        print(f"An error occurred while updating the sheet: {err}")


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

def process_xlsx(file_path):
    """Process the selected Excel file."""
    try:
        performance_metrics = {}
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            if sheet_name in {"Sheet 1", "Charlotte Airport", "Wayne 1", "Wayne 2", "Wayne 3"}:
                continue  # Skip sheets that we can't determine perfomance metircs based on average due to minimal data. Small Inv Groups            
            clean_sheet_name = sheet_name.replace(" Solar", "").replace(", LLC", "").replace(" Farm", "")
            dataframe = pd.read_excel(xls, sheet_name=sheet_name, skiprows=2)
            #print(f"Created DataFrame for sheet: {sheet_name}")
            
            # Convert Timestamps to Datetime Objs to run filter on time of day.
            dataframe.iloc[:, 0] = pd.to_datetime(dataframe.iloc[:, 0], errors='coerce')

            # New filtering logic
            # Prioritize POA, then GHI, then fall back to time.
            poa = SITE_DATA.get(clean_sheet_name, {}).get('POA')
            ghi = SITE_DATA.get(clean_sheet_name, {}).get('GHI')

            # Initialize conditions to False
            poa_condition = pd.Series([False] * len(dataframe), index=dataframe.index)
            ghi_condition = pd.Series([False] * len(dataframe), index=dataframe.index)

            if poa and poa in dataframe.columns:
                poa_condition = pd.to_numeric(dataframe[poa], errors='coerce') >= 200
            if ghi and ghi in dataframe.columns:
                ghi_condition = pd.to_numeric(dataframe[ghi], errors='coerce') >= 150

            time_condition = (dataframe.iloc[:, 0].dt.hour >= 9) & (dataframe.iloc[:, 0].dt.hour <= 15)

            # Apply filters in order of priority
            # A row is kept if it meets any of these conditions.
            dataframe = dataframe[
                poa_condition |
                (~poa_condition & ghi_condition) |
                (~(poa_condition | ghi_condition) & time_condition)
            ]

            inv_num = SITE_DATA.get(clean_sheet_name, {}).get('inv_num', 0)
            inverter_sums = {find_inverter_num(col): pd.to_numeric(dataframe[col], errors='coerce').sum() for col in dataframe.columns[1:inv_num+1]}
            if clean_sheet_name == 'Marshall':
                print(f"Inverter sums for {clean_sheet_name}: {inverter_sums}")
            if clean_sheet_name in INVERTER_GROUPS:
                performance_metrics[clean_sheet_name] = {}
                for group_name, inverter_numbers_set in INVERTER_GROUPS[clean_sheet_name].items():
                    group_sums = [inverter_sums.get(str(inv_num), 0) for inv_num in inverter_numbers_set]
                    max_sum = max(group_sums) if group_sums else 0
                    performance_metrics[clean_sheet_name][group_name] = max_sum if max_sum else 0
                    for inv_num in inverter_numbers_set:
                        inv_sum = inverter_sums.get(str(inv_num), 0)
                        ratio = (inv_sum / max_sum) if max_sum > 0 else 0
                        performance_metrics[clean_sheet_name][f"inverter_{inv_num}_ratio"] = ratio
            elif clean_sheet_name == "Duplin":
                performance_metrics[clean_sheet_name] = {}
                central_inverters = {}
                central_counter = 1
                string_inverters = {}
                string_counter = 1

                for col in dataframe.columns[1:]:
                    inv_sum = pd.to_numeric(dataframe[col], errors='coerce').sum()
                    if 'central' in col.lower():
                        central_inverters[central_counter] = inv_sum
                        #print(f"{col} Counter: {central_counter}, Sum: {inv_sum}")
                        central_counter += 1
                    elif 'string' in col.lower():
                        string_inverters[string_counter] = inv_sum
                        #print(f"{col} Counter: {string_counter}, Sum: {inv_sum}")
                        string_counter += 1

                for group_name, group_data in [("Central Inverters", central_inverters), ("String Inverters", string_inverters)]:
                    if group_data:
                        max_sum = max(group_data.values())
                        performance_metrics[clean_sheet_name][group_name] = max_sum if max_sum else 0
                        for inv_num, inv_sum in group_data.items():
                            ratio = (inv_sum / max_sum) if max_sum > 0 else 0
                            # Adjust inverter number for string inverters to be unique
                            if group_name == "String Inverters":
                                inv_num += len(central_inverters)
                            performance_metrics[clean_sheet_name][f"inverter_{inv_num}_ratio"] = ratio if ratio else 0
                #print(performance_metrics[clean_sheet_name])

            elif sheet_name == "Conetoe":
                performance_metrics[clean_sheet_name] = {}
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
                    performance_metrics[clean_sheet_name]["All Inverters"] = max_sum

                    for main_inv_num, total_sum in main_inverter_sums.items():
                        ratio = (total_sum / max_sum) if max_sum > 0 else 0
                        performance_metrics[clean_sheet_name][f"inverter_{main_inv_num}_ratio"] = ratio
                #print(performance_metrics[clean_sheet_name])
                
            else:
                # If the site is not in INVERTER_GROUPS and not an exception, treat all inverters as one group.
                performance_metrics[clean_sheet_name] = {}
                all_sums = list(inverter_sums.values())
                max_sum = max(all_sums) if all_sums else 0
                performance_metrics[clean_sheet_name]["All Inverters"] = max_sum
                if clean_sheet_name == 'Marshall':
                    print(f"{clean_sheet_name} - All Inverters max sum: {max_sum}\n{all_sums}")

                for inv_num, inv_sum in inverter_sums.items():
                    ratio = (inv_sum / max_sum) if max_sum > 0 else 0
                    performance_metrics[clean_sheet_name][f"inverter_{inv_num}_ratio"] = ratio

        
        update_performance_sheet(performance_metrics)

    except Exception as e:
        print(f"Error processing Excel file '{file_path}': {e}")



def browse_file():
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")], initialdir=downloads_folder)
    if file_path:
        process_xlsx(file_path)

def create_list_of_sheet_names(service, spreadsheet_id):
    """Create a list of all sheet names in the spreadsheet."""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        return [sheet['properties']['title'] for sheet in spreadsheet.get('sheets', [])]
    except HttpError as err:
        print(f"An error occurred: {err}")
        return []





# Create the main window
root = tk.Tk()
root.title("Performance Check")
try:
    root.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\tracker_3KU_icon.ico")
except Exception as e:
    print(f"Error loading icon: {e}")

browse_button = tk.Button(root, text="Browse files", command=browse_file, width= 50, height= 10)
browse_button.pack(fill='both')
frame = tk.Frame(root)
frame.pack()

root.mainloop()