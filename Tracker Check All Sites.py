#Tracker Check 
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from icecream import ic
import tkinter as tk
from tkinter import filedialog, ttk
import os
import ctypes
import re

from datetime import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
trackerSheet = '1EzgmTCAAkCuOIVVRTUtNOp_jXdZpT7Hob-DUhEuYmrE'

myappid = 'NCC.Tracker.Check'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)


credentials = None
if os.path.exists("Tracker-token.json"):
    credentials = Credentials.from_authorized_user_file("Tracker-token.json", SCOPES)
if not credentials or not credentials.valid:
    if credentials and credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(r"C:\Users\OMOPS.AzureAD\Python Codes\NCC-AutomationCredentials.json", SCOPES)
        credentials = flow.run_local_server(port=0)
    with open("Tracker-token.json", "w+") as token:
        token.write(credentials.to_json())




# DEFINING WHAT TR GOES WITH WHAT INV AND THEN DETERMINING THE IMPACT TO THE SITE.
tracker_data_dict = {
    "Shorthorn": {
        1: [25, 26, 27, 28], 2: [21, 22, 23, 24], 3: [1, 2, 3, 4], 4: [5, 6, 7, 8], 5: [9, 10, 11, 12], 6: [13, 14, 15, 20], 7: [16, 17, 18, 19], 8: [29, 30, 31, 32], 9: [33, 34, 35, 36], 10: [39, 40, 41, 42],
        11: [43, 44, 45, 46], 12: [57, 58, 59, 60], 13: [53, 54, 55, 56], 14: [49, 50, 51, 52], 15: [37, 38, 47, 48], 16: [79, 80, 81, 82], 17: [120, 121, 122, 123], 18: [116, 117, 118, 119], 19: [112, 113, 114, 115], 20: [108, 109, 110, 111],
        21: [61, 62, 63, 64, 65, 66], 22: [67, 68, 69, 70], 23: [71, 72, 73, 74], 24: [75, 76, 77, 78], 25: [124, 125, 126, 127, 128, 129], 26: [130, 131, 132, 133], 27: [134, 135, 136, 137], 28: [83, 84, 85, 86], 29: [87, 88, 89, 90], 30: [91, 92, 93, 94, 95],
        31: [96, 97, 98, 99], 32: [100, 101, 102, 103], 33: [104, 105, 106, 107], 34: [138, 139, 140, 141], 35: [142, 143, 144, 145, 146], 36: [147, 148, 149, 150], 37: [151, 152, 153, 154], 38: [155, 156, 157, 158], 39: [159, 160, 161, 162], 40: [163, 164, 165, 166],
        41: [167, 168, 169, 170], 42: [171, 172, 173, 174], 43: [175, 176, 177, 178], 44: [216, 217, 218, 219, 220], 45: [212, 213, 214, 215], 46: [208, 209, 210, 211], 47: [204, 205, 206, 207], 48: [200, 201, 202, 203], 49: [179, 180, 181, 182], 50: [183, 184, 185, 186, 187],
        51: [188, 189, 190, 191], 52: [192, 193, 194, 195], 53: [196, 197, 198, 199], 54: [224, 223, 222, 221], 55: [229, 228, 227, 226, 225], 56: [233, 232, 231, 230], 57: [237, 236, 235, 234], 58: [241, 240, 239, 238], 59: [245, 244, 243, 242], 60: [250, 249, 248, 247],
        61: [254, 253, 252, 251], 62: [258, 257, 256, 255], 63: [262, 261, 260, 259], 64: [266, 265, 264, 263], 65: [299, 298, 297, 296], 66: [295, 294, 293, 246], 67: [292, 291, 290, 289], 68: [288, 287, 286, 285, 284], 69: [283, 282, 281, 280], 70: [279, 278, 277, 276],
        71: [275, 274, 273, 272, 271], 72: [270, 269, 268, 267]
    },
    "Bulloch 1A": {
        1: [68, 69, 36], 2: [70, 71, 72, 73], 3: [74, 75, 76], 4: [77, 78, 79], 5: [80, 81, 82, 83], 6: [84, 85, 86], 7: [64, 65, 66, 67], 8: [61, 62, 63], 9: [57, 58, 59, 60], 10: [53, 54, 55, 56],
        11: [50, 51 , 52], 12: [46, 47, 48, 49], 13: [37, 38, 39, 40], 14: [42, 43, 44, 45], 15: [32, 33, 34, 35], 16: [29, 30, 31], 17: [25, 26, 27, 28], 18: [21, 22, 23, 24], 19: [18, 19, 20], 20: [14, 15, 16, 17],
        21: [10, 11, 12, 13], 22: [7, 8, 9], 23: [3, 4, 5, 6], 24: [41, 1, 2]
    },
    "Bulloch 1B": {
        1: [29, 28, 27, 26], 2: [25, 24, 23], 3: [22, 21, 20, 19], 4: [18, 17, 16, 15], 5: [14, 13, 12], 6: [11, 10, 9, 8], 7: [7, 6, 5], 8: [4, 3, 2, 1], 9: [42, 41, 40], 10: [39, 38, 37, 36],
        11: [35, 34, 33], 12: [32, 31, 30], 13: [83, 84, 85, 86], 14: [80, 81, 82], 15: [76, 77, 78, 79], 16: [71, 72, 73, 74], 17: [68, 69, 70], 18: [64, 65, 66, 67], 19: [61, 62, 63], 20: [57, 58, 59, 60],
        21: [53, 54, 55, 56], 22: [50, 51, 52], 23: [46, 47, 48, 49], 24: [45, 44, 43, 75]
    },
    "Gray Fox Solar": {
        1: [1, 2, 3, 4], 2: [5, 6, 7, 8], 3: [9, 10, 11, 12], 4: [13, 14, 15, 16], 5: [20, 17, 18, 19], 6: [24, 21, 22, 23], 7: [28, 25, 26, 27], 8: [32, 29, 30, 31], 9: [36, 33, 34, 35], 10: [40, 37, 38, 39],
        11: [41, 42, 43, 44], 12: [47, 48, 45, 46], 13: [51, 49, 50], 14: [55, 52, 53, 54], 15: [58, 59, 56, 57], 16: [60, 61, 62], 17: [66, 63, 64, 65], 18: [69, 67, 68, 70], 19: [73, 71, 72], 20: [74, 75, 76, 77],
        21: [78, 79, 80, 81], 22: [82, 83, 84, 85], 23: [86, 87, 88, 89], 24: [90, 91, 92, 93], 25: [94, 95, 96, 97], 26: [98, 99, 100, 101], 27: [102, 103, 104, 105], 28: [106, 107, 108, 109], 29: [110, 111, 112, 113], 30: [114, 115, 116, 117],
        31: [118, 119, 120, 121], 32: [122, 123, 124, 125], 33: [126, 127, 128], 34: [129, 130, 131, 132], 35: [133, 134, 135, 136], 36: [137, 138, 139], 37: [140, 141, 142, 143], 38: [144, 145, 146, 147], 39: [148, 149, 150], 40: [151, 152, 153, 154]
    },
    "Harding Solar": {
        1: [81, 82, 83, 84], 2: [85, 86, 87, 88], 3: [89, 90, 91, 92], 4: [93, 94, 95, 96], 5: [97, 98, 99, 100, 101], 6: [102, 103, 104, 105], 7: [58, 57, 56, 55, 54, 53], 8: [52, 51, 50, 49, 48], 9: [47, 46, 45, 44], 10: [43, 42, 41, 40],
        11: [39, 38, 37, 36, 35], 12: [34, 33, 32, 31], 13: [27, 28, 29, 30], 14: [22, 23, 24, 25, 26], 15: [18, 19, 20, 21], 16: [14, 15, 16, 17], 17: [10, 11, 12, 13], 18: [5, 6, 7, 8, 9], 19: [1, 2, 3, 4], 20: [59, 60, 61, 62, 63, 64],
        21: [65, 66, 67, 68], 22: [69, 70, 71, 72], 23: [73, 74, 75, 76], 24: [77, 78, 79, 80]
    },
    "McLean": {
        1: [72, 76, 80], 2: [59, 62, 65], 3: [47, 50, 53, 56], 4: [32, 35, 38, 41, 44], 5: [19, 21, 23, 25], 6: [27, 29, 31], 7: [34, 37, 40, 43], 8: [46, 49, 52, 55], 9: [58, 61, 64], 10: [67, 71, 75, 79],
        11: [119, 123, 127], 12: [130, 133, 136, 139], 13: [128, 131, 134, 137], 14: [116, 120, 124], 15: [100, 104, 108, 112], 16: [113, 117, 121, 125], 17: [101, 105, 109], 18: [93, 97, 92, 96], 19: [84, 87, 86, 83], 20: [69, 73, 77, 81, 68],
        21: [7, 9, 11, 13], 22: [1, 2, 3, 4, 5], 23: [6, 8, 10, 12, 14], 24: [15, 17, 16, 18], 25: [20, 22, 24], 26: [26, 28, 30, 33], 27: [36, 39, 42, 45], 28: [48, 51, 54], 29: [57, 60, 63, 66], 30: [70, 74, 78],
        31: [82, 85, 88, 90], 32: [94, 98, 102], 33: [106, 110, 114], 34: [118, 122, 126, 129], 35: [132, 135, 138], 36: [147, 148, 149, 150], 37: [142, 144, 146], 38: [141, 143, 145, 140], 39: [107, 111, 115], 40: [89, 91, 95, 99, 103]
    },
    "Richmond": {
        1: [1, 2, 43, 42], 2: [3, 4, 5, 6], 3: [7, 8, 9], 4: [10, 11, 12, 13], 5: [14, 15, 16, 17], 6: [18, 19, 20], 7: [21, 22, 23, 24], 8: [27, 25, 26], 9: [31, 30, 29, 28], 10: [34, 33, 32],
        11: [38, 37, 36, 35], 12: [41, 40, 39], 13: [47, 46, 45, 44], 14: [50, 49, 48], 15: [54, 53, 52, 51], 16: [58, 57, 56, 55], 17: [61, 60, 59], 18: [65, 64, 63, 62], 19: [69, 68, 67, 66], 20: [72, 71, 70],
        21: [76, 75, 74, 73], 22: [77, 78, 79], 23: [80, 81, 82, 83], 24: [84, 85, 86]
    },
    #I can't understand the legend for the sheet field #6-1 when AE is 1-355
    "Sunflower Solar": {
        1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [],
        11: [], 12: [], 13: [], 14: [], 15: [], 16: [], 17: [], 18: [], 19: [], 20: [],
        21: [], 22: [], 23: [], 24: [], 25: [], 26: [], 27: [], 28: [], 29: [], 30: [],
        31: [], 32: [], 33: [], 34: [], 35: [], 36: [], 37: [], 38: [], 39: [], 40: [],
        41: [], 42: [], 43: [], 44: [], 45: [], 46: [], 47: [], 48: [], 49: [], 50: [],
        51: [], 52: [], 53: [], 54: [], 55: [], 56: [], 57: [], 58: [], 59: [], 60: [],
        61: [], 62: [], 63: [], 64: [], 65: [], 66: [], 67: [], 68: [], 69: [], 70: [],
        71: [], 72: [], 73: [], 74: [], 75: [], 76: [], 77: [], 78: [], 79: [], 80: []
    },
    "Upson": {
        1: [1, 2, 3, 4], 2: [5, 6, 7], 3: [8, 9, 10, 11], 4: [12, 13, 14, 15], 5: [16, 17, 18], 6: [19, 20, 21, 22], 7: [23, 24, 25, 26], 8: [27, 28, 29], 9: [30, 31, 32, 33], 10: [34, 35, 36, 37],
        11: [38, 39, 40], 12: [41, 42, 43, 44], 13: [45, 46, 47, 48], 14: [49, 50, 51], 15: [52, 53, 54, 55], 16: [56, 57, 58, 59], 17: [60, 61, 62], 18: [63, 64, 65, 66], 19: [67, 68, 69, 70], 20: [71, 72, 73],
        21: [74, 75, 76, 77], 22: [78, 79, 80, 81], 23: [82, 83, 84], 24: [85, 86, 87, 88]
    },
    #Layout not Representative of the As Builts.
    "Washington Solar": {
        1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [],
        11: [], 12: [], 13: [], 14: [], 15: [], 16: [], 17: [], 18: [], 19: [], 20: [],
        21: [], 22: [], 23: [], 24: [], 25: [], 26: [], 27: [], 28: [], 29: [], 30: [],
        31: [], 32: [], 33: [], 34: [], 35: [], 36: [], 37: [], 38: [], 39: [], 40: []
    },
    "Whitehall Solar": {
        1: [1, 2, 3, 4], 2: [52, 53, 54, 55, 56, 57], 3: [5, 6, 7, 8], 4: [58, 59, 60, 61, 62, 63], 5: [9, 10, 11, 12], 6: [13, 14, 15, 16], 7: [17, 18, 19, 20, 21], 8: [22, 23, 24, 25], 9: [26, 27, 28, 29], 10: [30, 31, 32, 33, 34],
        11: [35, 36, 37, 38], 12: [39, 64, 65, 66, 67, 68], 13: [40, 41, 42, 43], 14: [69, 70, 71, 72, 73, 74], 15: [44, 45, 46, 47], 16: [48, 49, 50, 51]
    },
    #Either I made an error on the first row that I did or the As Built is wrong. Either way theres no comms with GC and AE for now so I'm not going to wokr on this util that is fixed. 
    "Whitetail": {
        1: [234, 237, 239, 242, 245], 2: [], 3: [231, 226, 223, 222, 218, 215], 4: [], 5: [215, 213, 210, 207, 203], 6: [], 7: [201, 198, 195, 190, 188], 8: [], 9: [185, 182, 179, 176, 174], 10: [],
        11: [171, 168, 165, 162, 159], 12: [], 13: [155, 152, 150, 147, 144], 14: [], 15: [140, 138, 134, 133, 132], 16: [], 17: [130, 126, 123, 120, 117, 114], 18: [], 19: [111, 109, 106, 103, 98, 95], 20: [],
        21: [], 22: [], 23: [], 24: [], 25: [], 26: [], 27: [], 28: [], 29: [], 30: [],
        31: [], 32: [], 33: [], 34: [], 35: [], 36: [], 37: [], 38: [], 39: [], 40: [],
        41: [], 42: [], 43: [], 44: [26, 23, 20, 17, 14, 11], 45: [41, 38, 35, 32, 29], 46: [59, 56, 53, 50, 47, 44], 47: [77, 74, 71, 68, 65, 62], 48: [92, 89, 86, 83, 80], 49: [], 50: [],
        51: [], 52: [], 53: [], 54: [], 55: [], 56: [], 57: [], 58: [], 59: [], 60: [],
        61: [], 62: [], 63: [], 64: [], 65: [], 66: [], 67: [], 68: [], 69: [], 70: [],
        71: [], 72: [], 73: [], 74: [], 75: [], 76: [], 77: [], 78: [], 79: [], 80: []
    },
    #Extra Sites with Sungrows these will get the 10% EST for Inv loss and then we can calc the Site %
    "Bluebird Solar": {
        1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [],
        11: [], 12: [], 13: [], 14: [], 15: [], 16: [], 17: [], 18: [], 19: [], 20: [],
        21: [], 22: [], 23: [], 24: []
    },
    "Cardinal": {
        1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [],
        11: [], 12: [], 13: [], 14: [], 15: [], 16: [], 17: [], 18: [], 19: [], 20: [],
        21: [], 22: [], 23: [], 24: [], 25: [], 26: [], 27: [], 28: [], 29: [], 30: [],
        31: [], 32: [], 33: [], 34: [], 35: [], 36: [], 37: [], 38: [], 39: [], 40: [],
        41: [], 42: [], 43: [], 44: [], 45: [], 46: [], 47: [], 48: [], 49: [], 50: [],
        51: [], 52: [], 53: [], 54: [], 55: [], 56: [], 57: [], 58: [], 59: []
    },
    "Freight Line Solar": {
        1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [],
        11: [], 12: [], 13: [], 14: [], 15: [], 16: [], 17: [], 18: []
    },
    "Hayes": {
        1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [],
        11: [], 12: [], 13: [], 14: [], 15: [], 16: [], 17: [], 18: [], 19: [], 20: [],
        21: [], 22: [], 23: [], 24: [], 25: [], 26: []
    },
    "Hickson Solar Farm": {
        1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [],
        11: [], 12: [], 13: [], 14: [], 15: [], 16: []
    },
    "Holly Swamp Solar": {
        1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [],
        11: [], 12: [], 13: [], 14: [], 15: [], 16: []
    },
    "Jefferson Solar": {
        1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [],
        11: [], 12: [], 13: [], 14: [], 15: [], 16: [], 17: [], 18: [], 19: [], 20: [],
        21: [], 22: [], 23: [], 24: [], 25: [], 26: [], 27: [], 28: [], 29: [], 30: [],
        31: [], 32: [], 33: [], 34: [], 35: [], 36: [], 37: [], 38: [], 39: [], 40: [],
        41: [], 42: [], 43: [], 44: [], 45: [], 46: [], 47: [], 48: [], 49: [], 50: [],
        51: [], 52: [], 53: [], 54: [], 55: [], 56: [], 57: [], 58: [], 59: [], 60: [],
        61: [], 62: [], 63: [], 64: []
    },
    "Marshall Solar": {
        1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [],
        11: [], 12: [], 13: [], 14: [], 15: [], 16: []
    },
    "Ogburn Solar Farm": {
        1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [],
        11: [], 12: [], 13: [], 14: [], 15: [], 16: []
    },
    "PG Solar": {
        1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [],
        11: [], 12: [], 13: [], 14: [], 15: [], 16: [], 17: [], 18: []
    },
    "Tedder Solar": {
        1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [],
        11: [], 12: [], 13: [], 14: [], 15: [], 16: []
    },
    "Thunderhead Solar": {
        1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [],
        11: [], 12: [], 13: [], 14: [], 15: [], 16: []
    },
    "Van Buren Solar": {
        1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: [], 8: [], 9: [], 10: [],
        11: [], 12: [], 13: [], 14: [], 15: [], 16: [], 17: []
    }
}
#Maybe another dictionary or something like chekcing the len of the site dictionary to see if Central or not for the altered % effect on Inv and therefore site % Loss





def format_tracker_result(match, sheet_name):
    if match:
        groups = match.groups()
        if len(groups) == 2 and sheet_name in ['Cherry Blossom Solar, LLC', 'Conetoe']:
                return f"Tracker {groups[0]} Motor {groups[1]}"
        elif len(groups) == 2 and sheet_name in ['Hickory Solar, LLC', 'Shorthorn', 'Sunflower Solar']:
            if sheet_name in ['Shorthorn', 'Sunflower Solar']:
                groups = list(groups)  # Convert tuple to list to allow modification
                groups[0] = str(int(groups[0]) + 1) 
            return f"NCU {groups[0]} Tracker {groups[1]}"
        elif sheet_name == 'Whitetail':
            if len(groups) == 1:
                return f"Master 1: Tracker {groups[0]}"
            else:
                return f"Master 2: Tracker {groups[1]}"
        elif len(groups) == 2:
                return f"Zone {groups[0]} Tracker {groups[1]}"
        elif len(groups) == 1:
                return f"Tracker {groups[0]}"
    else:
        return "ERROR: Tracker # Not Found"


def process_file(file_path):
    # Load the Excel file with multiple sheets
    xls = pd.ExcelFile(file_path)
    print("Started", datetime.now())

    for sheet_name in xls.sheet_names:
        if sheet_name != "Sheet 1":
            df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=2)
            # Remove columns that include the word 'Setpoint' or 'Target' in the field name
            df = df.loc[:, ~df.columns.str.contains('SetPoint|Target', case=False)]
            results = []
            for column in df.columns[1:]:# Iterate through each column except the first one
                days4 = df.iloc[1:][column].tail(36) #Hours from Midnight Last Night counting back
                days4 = pd.to_numeric(days4, errors='coerce') #Sets to a number, correcting negative values
                

                #Need Patterns for Upson Whitehall Harding Gray Fox and Washington atleast
                if sheet_name in ["Bulloch 1A", "Bulloch 1B", "Richmond", "McLean", "Upson"]:
                    pattern = r"A\d S (\d+)"
                elif sheet_name in ['Hickory Solar, LLC']:
                    pattern = r"Tracker (\d+).*A\d S (\d+)"
                elif sheet_name in ["Bishopville II Solar", "Jefferson Solar"]:
                    pattern = r"\(ZC(\d)\).*Angle (\d+)"
                elif sheet_name in ["Harding Solar", "Washington Solar", "Whitehall Solar", "Gray Fox Solar"]:
                    pattern = r"TCU (\d+)"
                elif sheet_name in ["Bluebird Solar", "Cardinal", "Freight Line Solar", "Hayes", "Holly Swamp Solar", "PG Solar", "Van Buren Solar"]:
                    pattern = r"angle (\d+)"
                elif sheet_name in ['Cherry Blossom Solar, LLC', 'Conetoe']:
                    pattern = r"Tracker.*(\d+).*Motor (\d+)"
                elif sheet_name in ['Marshall Solar', 'Tedder Solar', 'Thunderhead Solar']:
                    pattern = r"Angle (\d+)"
                elif sheet_name in ['Shorthorn']:
                    pattern = r"\(ST(\d)\).*A\d S (\d+)"
                elif sheet_name in ['Sunflower Solar']:
                    pattern = r"\(ST(\d)\).*TCU (\d+)"
                elif sheet_name in ['Whitetail', "Ogburn Solar Farm", "Hickson Solar Farm"]:
                    pattern = r"(\d+)"



                if pattern:
                    inv_match = re.search(pattern, column)
                    tr_num = inv_match.group(1) if inv_match else 0
                    if sheet_name in tracker_data_dict:
                        for key, value_list in tracker_data_dict[sheet_name].items():
                            if int(tr_num) in value_list:
                                inv = key
                                #Define 50% as max loss of Inverter or Performance % lost.
                                #50% is my guess
                                #8/2/23 Holly Swamp 45W Example 43% Loss
                                # 6/3/23 Perfect Day   Holly Swamp 
                                # Flat 
                                inv_loss_fl = (.43/len(value_list))
                                inv_loss = f"{round(inv_loss_fl * 100, 2)}%"   #as a Percentage. 
                                site_loss_fl = inv_loss_fl/len(tracker_data_dict[sheet_name])
                                site_loss = f"{round(site_loss_fl, 5)}%"
                                break
                            else:
                                inv = "Unknown"
                                inv_loss = "Est. 15%"
                                inv_loss_fl = .15
                                site_loss_fl = inv_loss_fl/len(tracker_data_dict[sheet_name])
                                site_loss = f"Est. {round(site_loss_fl, 5)}%"
                                
                                
                    else:
                        inv = "Unknown"
                        inv_loss = "No Data"
                        site_loss = "No Data"
                        inv_loss_fl = "No Data"
                        site_loss_fl = "No Data"
                else:
                    inv = "Unknown"
                    inv_loss = "No Data"
                    site_loss = "No Data"
                    inv_loss_fl = "No Data"
                    site_loss_fl = "No Data"




                
                if days4.apply(lambda x: pd.isna(x) or x == '').all(): #This find the Lost Comms Trackers
                    for i in range(len(df)-1, -1, -1): #Finds the Date it stopped tracking
                        date = "Over a Month ago"
                        value = pd.to_numeric(df.iloc[i][column], errors='coerce') # Ensure the value is numeric before performing the subtraction
                        if pd.notna(value): #Skips NA Values 
                            date = df.iloc[i, 0] #Selects the Cell
                    tracker_obj = re.search(pattern, column)
                    tracker = format_tracker_result(tracker_obj, sheet_name)
                    results.append([tracker, "Lost Comms", str(date), f"Inverter: {inv}", f"Performance Loss on INV: {inv_loss} SITE: {site_loss}", inv, inv_loss_fl, site_loss_fl, str(date)]) 
                else:
                    position = days4.mean() #Avg position accross data set                
                
                    if (days4.max() - days4.min()) <= 5: #Finds/Indicates Stuck tracker and initiates adding data to results              
                        for i in range(len(df)-1, -1, -1): #Finds the Date it stopped tracking
                            date = "Over a Month ago"
                            value = pd.to_numeric(df.iloc[i][column], errors='coerce') # Ensure the value is numeric before performing the subtraction
                            if pd.isna(value): #Skips NA Values 
                                continue
                            if abs(value - position) > 5: 
                                date = df.iloc[i, 0] #Selects the Cell
                                break
                        tracker_obj = re.search(pattern, column)
                        tracker = format_tracker_result(tracker_obj, sheet_name)
                        results.append([tracker, f"Angle: {round(position, 3)}", f"Since: {str(date)}", f"Inverter: {inv}", f"Performance Loss on INV: {inv_loss} SITE: {site_loss}", inv, inv_loss_fl, site_loss_fl, str(date)]) 
            clear_past(sheet_name) 
            if results:
                if results[0]:
                    #ic(sheet_name, results)
                    write_results(sheet_name, results)
    print("Finished", datetime.now())
    root.destroy() #Finished moving data to Sheet

def clear_past(sheet_name):
    print("Clearing", sheet_name)
    try:
        service = build('sheets', 'v4', credentials=credentials)
        sheet = service.spreadsheets()

        # Clear the contents of the sheet
        clear_range = f"{sheet_name}!A1:I1000"  # Adjust the range as needed
        sheet.values().clear(spreadsheetId=trackerSheet, range=clear_range).execute()

    except HttpError as err:
        print(err)
    except IndexError as outbounds:
        print(outbounds)
            
def write_results(sheet_name, data):
    print("Writing", sheet_name)
    try:
        service = build('sheets', 'v4', credentials=credentials)
        sheet = service.spreadsheets()

        # Write new data to the sheet
        body = {"values": data}
        sheet.values().update(spreadsheetId=trackerSheet, range=f"{sheet_name}!A1", valueInputOption="USER_ENTERED", body=body).execute()

    except HttpError as err:
        print(err)


def browse_file():
    notes = tk.Label(frame, text="I am Processing a large data set which may take 30 minutes or more...") # Started at 09:14 end 09:32
    notes2 = tk.Label(frame, text="During which, my window title may say Tracker Check (Not Responding)")
    notes3 = tk.Label(frame, text="Feel Free to Use the PC like you normally would")

    notes.pack()
    notes2.pack()
    notes3.pack()
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")], initialdir=downloads_folder)
    process_file(file_path)






#This is where the Script Starts. Predefined Shit is Above.
# Create the main window
root = tk.Tk()
root.title("Tracker Check")
try:
    root.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\tracker_3KU_icon.ico")
except Exception as e:
    print(f"Error loading icon: {e}")

browse_button = tk.Button(root, text="Browse files", command=browse_file, width= 50, height= 10)
browse_button.pack(fill='both')
frame = tk.Frame(root)
frame.pack()

root.mainloop()