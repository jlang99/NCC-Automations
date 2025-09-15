#Tracker Check 
import pandas as pd
import numpy as np
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from icecream import ic
import tkinter as tk
from tkinter import filedialog, ttk
import os
import ctypes
import re
import time

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
cbSheet = '1RGUwARwDdfDoC8VcNQgb5KC6gPOA-rcMaPsu6vlzg9o'

myappid = 'NCC.CB.Check'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
conetoe_24strs = {
    '1.1', '1.2', '1.3', '1.4', '1.5', '1.6', '1.7', '1.8', '1.9', '1.10', 
    '1.11', '1.12', '2.1', '2.2', '2.3', '2.4', '2.5', '2.6', '2.7', '2.8', 
    '2.9', '2.10', '2.11', '2.12', '3.1', '3.2', '3.3', '3.4', '3.5', '3.6', 
    '3.7', '3.8', '3.9', '3.10', '3.11', '3.12', '4.1', '4.2', '4.3', '4.4', 
    '4.5', '4.6', '4.7', '4.8', '4.9', '4.10', '4.11', '4.12'
}
conetoe_20strs = {'1.13'}
conetoe_cbs = [conetoe_20strs, conetoe_24strs]
hickory_15strs = {
    'INV1A1 A', 'INV1A2 A', 'INV1A3 A', 'INV1A4 A', 'INV1A5 A', 'INV1A6 A', 
    'INV1A7 A', 'INV1A8 A', 'INV1A9 A', 'INV1B13 B', 'INV1B14 B', 'INV1B15 B', 
    'INV1B16 B', 'INV1B17 B', 'INV1B18 B', 'INV1B19 B', 'INV1B20 B', 'INV1B21 B', 
    'INV2A1 A', 'INV2A2 A', 'INV2A6 A', 'INV2A7 A', 'INV2A8 A', 'INV2A9 A', 
    'INV2A10 A', 'INV2A11 A', 'INV2A12 A', 'INV2B13 B', 'INV2B16 B', 'INV2B17 B', 
    'INV2B23 B'
}

hickory_12strs = {
    'INV1A10 A', 'INV1B23 B', 'INV1B24 B', 'INV2A3 A', 'INV2A4 A'
}

hickory_18strs = {
    'INV1A11 A', 'INV2B22 B'
}

hickory_20strs = {
    'INV1A12 A'
}

hickory_11strs = {
    'INV2A5 A'
}

hickory_16strs = {
    'INV2B14 B'
}

hickory_17strs = {
    'INV2B15 B', 'INV2B19 B'
}

hickory_14strs = {
    'INV2B20 B', 'INV2B24 B'
}
hickory_cbs = [hickory_11strs, hickory_12strs, hickory_14strs, hickory_15strs, hickory_16strs, hickory_17strs, hickory_18strs, hickory_20strs]
wellons_22strs = {
    'CB1-1', 'CB1-2', 'CB1-3', 'CB1-4', 'CB1-5', 'CB1-6', 'CB1-9', 'CB1-10', 
    'CB1-11', 'CB1-12', 'CB1-13', 'CB1-14', 'CB1-15', 'CB1-16', 'CB2-1', 'CB2-2', 
    'CB2-3', 'CB2-4', 'CB2-5', 'CB2-6', 'CB2-7'
}

wellons_18strs = {
    'CB1-8', 'CB2-10', 'CB2-12', 'CB2-13', 'CB2-14', 'CB2-15', 'CB2-16', 'CB2-17'
}

wellons_20strs = {
    'CB1-17', 'CB1-18', 'CB2-8', 'CB2-9', 'CB2-11', 'CB2-18'
}

wellons_28strs = {
    'CB3-1', 'CB3-2', 'CB3-3', 'CB3-4', 'CB3-5', 'CB3-6', 'CB3-7', 'CB3-8', 
    'CB3-9', 'CB3-10', 'CB3-11', 'CB3-12', 'CB3-13', 'CB3-14'
}
wellons_cbs = [wellons_18strs, wellons_20strs, wellons_22strs, wellons_28strs]
violet18strs = {
    'dc_11', 'dc_13', 'dc_14', 'dc_15', 'dc_16', 'dc_17', 'dc_18', 'dc_19', 
    'dc_110', 'dc_111', 'dc_112', 'dc_22', 'dc_23', 'dc_24', 'dc_25', 'dc_26', 
    'dc_27', 'dc_212', 'dc_213', 'dc_41', 'dc_42', 'dc_43', 'dc_44', 'dc_45', 
    'dc_46', 'dc_47'}
violet16strs = {'dc_21', 'dc_28', 'dc_29', 'dc_210', 'dc_37', 'dc_410', 'dc_48', 'dc_311', 'dc_312'}
violet14strs = {'dc_38', 'dc_39', 'dc_310', 'dc_12'}
violet20strs = {'dc_32', 'dc_33', 'dc_34'}
violet28strs = {'dc_35', 'dc_36'}
violetNA = {'dc_211', 'dc_411', 'dc_412'}
violetcbs = [violet18strs, violet16strs, violet14strs, violet20strs, violet28strs]

credentials = None
if os.path.exists("aeCB-token.json"):
    credentials = Credentials.from_authorized_user_file("aeCB-token.json", SCOPES)
if not credentials or not credentials.valid:
    try:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            # This block will run if credentials are None, invalid, or don't have a refresh token.
            flow = InstalledAppFlow.from_client_secrets_file(r"G:\Shared drives\O&M\NCC Automations\Daily Automations\NCC-AutomationCredentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("aeCB-token.json", "w") as token:
            token.write(credentials.to_json())
    except Exception as e: # Catches google.auth.exceptions.RefreshError and others
        if 'invalid_grant' in str(e):
            print("Refresh token is invalid. Deleting token file and re-authenticating.")
            os.remove("aeCB-token.json")
            # After deleting the token, we explicitly start the flow again.
            flow = InstalledAppFlow.from_client_secrets_file(r"G:\Shared drives\O&M\NCC Automations\Daily Automations\NCC-AutomationCredentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
            with open("aeCB-token.json", "w") as token:
                token.write(credentials.to_json())
        else:
            raise # Re-raise the exception if it's not an invalid_grant error.
        flow = InstalledAppFlow.from_client_secrets_file(r"G:\Shared drives\O&M\NCC Automations\Daily Automations\NCC-AutomationCredentials.json", SCOPES)
        credentials = flow.run_local_server(port=0)
    with open("aeCB-token.json", "w+") as token:
        token.write(credentials.to_json())




def process_file(file_path):
    # Load the Excel file with multiple sheets
    xls = pd.ExcelFile(file_path)

    for sheet_name in xls.sheet_names:
        if sheet_name != "Sheet 1":
            try:
                service = build('sheets', 'v4', credentials=credentials)
                sheet = service.spreadsheets()

                # Clear the contents of the sheet
                clear_range = f"{sheet_name}!A1:D1000"  # Adjust the range as needed
                sheet.values().clear(spreadsheetId=cbSheet, range=clear_range).execute()
            except HttpError as err:
                print(err)
        if sheet_name == "Violet Solar, LLC":


            df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=2)
            dfnew = df.iloc[1:] #Clear the first row from Data frame as it holds the Units for the Column
            # Convert the first column to datetime
            dfnew.iloc[:, 0] = pd.to_datetime(dfnew.iloc[:, 0], errors='coerce')

            # Filter rows based on time of day
            dfnew = dfnew[(dfnew.iloc[:, 0].dt.time >= pd.to_datetime('09:00 AM').time()) & 
                          (dfnew.iloc[:, 0].dt.time <= pd.to_datetime('05:00 PM').time())]
            underperformers = {}
            for column in dfnew.columns[1:]:# Iterate through each column except the first one
                if column not in violetNA:
                    for cb_list in violetcbs:
                        for nam_cb in cb_list:
                            if column == nam_cb:
                                if cb_list == violet18strs:
                                    strings = 18
                                elif cb_list == violet14strs:
                                    strings = 14
                                elif cb_list == violet16strs:
                                    strings = 16
                                elif cb_list == violet20strs:
                                    strings = 20
                                elif cb_list == violet28strs:
                                    strings = 28
                                underperformers[f'{sheet_name} {strings} Strings'] = []
            for column in dfnew.columns[1:]:# Iterate through each column except the first one
                if column not in violetNA:
                    for cb_list in violetcbs:
                        for nam_cb in cb_list:
                            if column == nam_cb:
                                numeric_column = pd.to_numeric(dfnew[column], errors='coerce') #Sets to a number, correcting negative values
                                cbavg = numeric_column.mean(skipna=True)
                                if np.isnan(cbavg):
                                    cbavg = "No Comms or Device Offline, for a Month"  

                                if cb_list == violet18strs:
                                    strings = 18
                                elif cb_list == violet14strs:
                                    strings = 14
                                elif cb_list == violet16strs:
                                    strings = 16
                                elif cb_list == violet20strs:
                                    strings = 20
                                elif cb_list == violet28strs:
                                    strings = 28
                                underperformers[f'{sheet_name} {strings} Strings'].append([column, cbavg])
            for key, list_avgs in underperformers.items(): #This is all CB's in lists sorted by Number of Strings
                if list_avgs:
                    results = []
                    if len(list_avgs) > 1:
                        new_list = [nameavg[1] for nameavg in list_avgs]
                        filtered_list = [cb for cb in new_list if cb != "No Comms or Device Offline, for a Month"]
                        siteMean = pd.Series(filtered_list).mean()
                        comparison_value = siteMean * 0.90
                        for cb in list_avgs:
                            if cb[1] == "No Comms or Device Offline, for a Month":
                                results.append([cb[0], cb[1], key])
                            elif cb[1] < comparison_value:  # Check if the value is less than 5% of siteMean
                                ratio = round((cb[1]/siteMean) * 100, 1)
                                results.append([cb[0], f"Average Performance: {round(cb[1], 2)} KWH  {ratio}% of the Average", key + f" Average KWH: {round(siteMean, 2)}", ratio])        
                        if results:
                            if results[0]:
                                ic(results)
                                time.sleep(2.5)
                                write_results(sheet_name, results)
                    else:
                        ic("I am Unique", key, list_avgs[0][0], list_avgs[0][1])
                        if list_avgs[0][1] == "No Comms or Device Offline, for a Month":
                            results.append([list_avgs[0][0], list_avgs[0][1], key])
                        elif list_avgs[0][1]:
                            kwh = round(list_avgs[0][1], 1)
                            ratio = round((kwh/3) * 100, 1)
                            results.append([list_avgs[0][0], f"Last Month's Average KWH: {kwh} No other CB's with this # of Strings", key, ratio])
                        if results:
                            if results[0]:
                                ic(results)
                                time.sleep(2.5)
                                write_results(sheet_name, results)
        elif sheet_name == "Conetoe":
            df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=2)
            dfnew = df.iloc[1:] #Clear the first row from Data frame as it holds the Units for the Column
            # Convert the first column to datetime
            dfnew.iloc[:, 0] = pd.to_datetime(dfnew.iloc[:, 0], errors='coerce')

            # Filter rows based on time of day
            dfnew = dfnew[(dfnew.iloc[:, 0].dt.time >= pd.to_datetime('09:00 AM').time()) & 
                          (dfnew.iloc[:, 0].dt.time <= pd.to_datetime('05:00 PM').time())]
            underperformers = {}
            for column in dfnew.columns[1:]:# Iterate through each column except the first one
                    for cb_list in conetoe_cbs:
                        for nam_cb in cb_list:
                            if column == nam_cb:
                                if cb_list == conetoe_20strs:
                                    strings = 20
                                elif cb_list == conetoe_24strs:
                                    strings = 24
                                underperformers[f'{sheet_name} {strings} Strings'] = []
            for column in dfnew.columns[1:]:# Iterate through each column except the first one
                    for cb_list in conetoe_cbs:
                        for nam_cb in cb_list:
                            if column == nam_cb:
                                numeric_column = pd.to_numeric(dfnew[column], errors='coerce') #Sets to a number, correcting negative values
                                cbavg = numeric_column.mean(skipna=True)
                                if np.isnan(cbavg):
                                    cbavg = "No Comms or Device Offline, for a Month"  

                                if cb_list == conetoe_20strs:
                                    ic(column)
                                    strings = 20
                                elif cb_list == conetoe_24strs:
                                    strings = 24
                                underperformers[f'{sheet_name} {strings} Strings'].append([column, cbavg])
            for key, list_avgs in underperformers.items(): #This is all CB's in lists sorted by Number of Strings
                if list_avgs:
                    results = []
                    if len(list_avgs) > 1:
                        new_list = [nameavg[1] for nameavg in list_avgs]
                        filtered_list = [cb for cb in new_list if cb != "No Comms or Device Offline, for a Month"]
                        ic(new_list, filtered_list)
                        siteMean = pd.Series(filtered_list).mean()
                        comparison_value = siteMean * 0.90
                        for cb in list_avgs:
                            if cb[1] == "No Comms or Device Offline, for a Month":
                                results.append([cb[0], cb[1], key])
                            elif cb[1] < comparison_value:  # Check if the value is less than 5% of siteMean
                                ratio = round((cb[1]/siteMean) * 100, 1)
                                results.append([cb[0], f"Average Performance: {round(cb[1], 2)} KWH  {ratio}% of the Average", key + f" Average KWH: {round(siteMean, 2)}", ratio])        
                        if results:
                            if results[0]:
                                ic(results)
                                time.sleep(2.5)
                                write_results(sheet_name, results)
                    else:
                        ic("I am Unique", key, list_avgs[0][0], list_avgs[0][1])
                        if list_avgs[0][1] == "No Comms or Device Offline, for a Month":
                            results.append([list_avgs[0][0], list_avgs[0][1], key])
                        elif list_avgs[0][1]:
                            kwh = round(list_avgs[0][1], 1)
                            results.append([list_avgs[0][0], f"Last Month's Average KWH: {kwh} No other CB's with this # of Strings", key])
                        if results:
                            if results[0]:
                                ic(results)
                                time.sleep(2.5)
                                write_results(sheet_name, results)
        elif sheet_name == "Hickory Solar, LLC":
            df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=2)
            dfnew = df.iloc[1:] #Clear the first row from Data frame as it holds the Units for the Column
            # Convert the first column to datetime
            dfnew.iloc[:, 0] = pd.to_datetime(dfnew.iloc[:, 0], errors='coerce')

            # Filter rows based on time of day
            dfnew = dfnew[(dfnew.iloc[:, 0].dt.time >= pd.to_datetime('09:00 AM').time()) & 
                          (dfnew.iloc[:, 0].dt.time <= pd.to_datetime('05:00 PM').time())]
            underperformers = {}
            for column in dfnew.columns[1:]:# Iterate through each column except the first one
                    for cb_list in hickory_cbs:
                        for nam_cb in cb_list:
                            if column == nam_cb:
                                if cb_list == hickory_20strs:
                                    strings = 20
                                elif cb_list == hickory_14strs:
                                    strings = 14
                                elif cb_list == hickory_12strs:
                                    strings = 12
                                elif cb_list == hickory_11strs:
                                    strings = 11
                                elif cb_list == hickory_15strs:
                                    strings = 15
                                elif cb_list == hickory_16strs:
                                    strings = 16
                                elif cb_list == hickory_17strs:
                                    strings = 17
                                elif cb_list == hickory_18strs:
                                    strings = 18
                                underperformers[f'{sheet_name} {strings} Strings'] = []
            for column in dfnew.columns[1:]:# Iterate through each column except the first one
                    for cb_list in hickory_cbs:
                        for nam_cb in cb_list:
                            if column == nam_cb:
                                numeric_column = pd.to_numeric(dfnew[column], errors='coerce') #Sets to a number, correcting negative values
                                cbavg = numeric_column.mean(skipna=True)
                                if np.isnan(cbavg):
                                    cbavg = "No Comms or Device Offline, for a Month"  

                                if cb_list == hickory_20strs:
                                    strings = 20
                                elif cb_list == hickory_14strs:
                                    strings = 14
                                elif cb_list == hickory_12strs:
                                    strings = 12
                                elif cb_list == hickory_11strs:
                                    strings = 11
                                elif cb_list == hickory_15strs:
                                    strings = 15
                                elif cb_list == hickory_16strs:
                                    strings = 16
                                elif cb_list == hickory_17strs:
                                    strings = 17
                                elif cb_list == hickory_18strs:
                                    strings = 18
                                underperformers[f'{sheet_name} {strings} Strings'].append([column, cbavg])
            for key, list_avgs in underperformers.items(): #This is all CB's in lists sorted by Number of Strings
                if list_avgs:
                    results = []
                    if len(list_avgs) > 1:
                        new_list = [nameavg[1] for nameavg in list_avgs]
                        filtered_list = [cb for cb in new_list if cb != "No Comms or Device Offline, for a Month"]
                        siteMean = pd.Series(filtered_list).mean()
                        comparison_value = siteMean * 0.90
                        for cb in list_avgs:
                            if cb[1] == "No Comms or Device Offline, for a Month":
                                results.append([cb[0], cb[1], key])
                            elif cb[1] < comparison_value:  # Check if the value is less than 5% of siteMean
                                ratio = round((cb[1]/siteMean) * 100, 1)
                                results.append([cb[0], f"Average Performance: {round(cb[1], 2)} KWH.  {ratio}% of the Average", key + f" Average KWH: {round(siteMean, 2)}", ratio])        
                        if results:
                            if results[0]:
                                ic(results)
                                time.sleep(2.5)
                                write_results(sheet_name, results)
                    else:
                        ic("I am Unique", key, list_avgs[0][0], list_avgs[0][1])
                        if list_avgs[0][1] == "No Comms or Device Offline, for a Month":
                            results.append([list_avgs[0][0], list_avgs[0][1], key])
                        elif list_avgs[0][1]:
                            kwh = round(list_avgs[0][1], 1)
                            ratio = round((kwh/3) * 100, 1)
                            results.append([list_avgs[0][0], f"Last Month's Average KWH: {kwh} No other CB's with this # of Strings", key, ratio])
                        if results:
                            if results[0]:
                                ic(results)
                                time.sleep(2.5)
                                write_results(sheet_name, results)
        elif sheet_name == "Wellons Solar, LLC":
            df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=2)
            dfnew = df.iloc[1:] #Clear the first row from Data frame as it holds the Units for the Column
            # Convert the first column to datetime
            dfnew.iloc[:, 0] = pd.to_datetime(dfnew.iloc[:, 0], errors='coerce')

            # Filter rows based on time of day
            dfnew = dfnew[(dfnew.iloc[:, 0].dt.time >= pd.to_datetime('09:00 AM').time()) & 
                          (dfnew.iloc[:, 0].dt.time <= pd.to_datetime('05:00 PM').time())]
            underperformers = {}
            for column in dfnew.columns[1:]:# Iterate through each column except the first one
                    for cb_list in wellons_cbs:
                        for nam_cb in cb_list:
                            if column == nam_cb:
                                if cb_list == wellons_20strs:
                                    strings = 20
                                elif cb_list == wellons_28strs:
                                    strings = 28
                                elif cb_list == wellons_22strs:
                                    strings = 22
                                elif cb_list == wellons_18strs:
                                    strings = 18
                                underperformers[f'{sheet_name} {strings} Strings'] = []
            for column in dfnew.columns[1:]:# Iterate through each column except the first one
                    for cb_list in wellons_cbs:
                        for nam_cb in cb_list:
                            if column == nam_cb:
                                numeric_column = pd.to_numeric(dfnew[column], errors='coerce') #Sets to a number, correcting negative values
                                cbavg = numeric_column.mean(skipna=True)
                                if np.isnan(cbavg):
                                    cbavg = "No Comms or Device Offline, for a Month"  

                                if cb_list == wellons_20strs:
                                    strings = 20
                                elif cb_list == wellons_28strs:
                                    strings = 28
                                elif cb_list == wellons_22strs:
                                    strings = 22
                                elif cb_list == wellons_18strs:
                                    strings = 18
                                underperformers[f'{sheet_name} {strings} Strings'].append([column, cbavg])
            for key, list_avgs in underperformers.items(): #This is all CB's in lists sorted by Number of Strings
                if list_avgs:
                    results = []
                    if len(list_avgs) > 1:
                        new_list = [nameavg[1] for nameavg in list_avgs]
                        filtered_list = [cb for cb in new_list if cb != "No Comms or Device Offline, for a Month"]
                        siteMean = pd.Series(filtered_list).mean()
                        comparison_value = siteMean * 0.90
                        for cb in list_avgs:
                            if cb[1] == "No Comms or Device Offline, for a Month":
                                results.append([cb[0], cb[1], key])
                            elif cb[1] < comparison_value:  # Check if the value is less than 5% of siteMean
                                ratio = round((cb[1]/siteMean) * 100, 1)
                                results.append([cb[0], f"Average Performance: {round(cb[1], 2)} KWH  {ratio}% of the Average", key + f" Average KWH: {round(siteMean, 2)}", ratio])        
                        if results:
                            if results[0]:
                                ic(results)
                                time.sleep(2.5)
                                write_results(sheet_name, results)
                    else:
                        ic("I am Unique", key, list_avgs[0][0], list_avgs[0][1])
                        if list_avgs[0][1] == "No Comms or Device Offline, for a Month":
                            results.append([list_avgs[0][0], list_avgs[0][1], key])
                        elif list_avgs[0][1]:
                            kwh = round(list_avgs[0][1], 1)
                            ratio = round((kwh/3) * 100, 1)
                            results.append([list_avgs[0][0], f"Last Month's Average KWH: {kwh} No other CB's with this # of Strings", key, ratio])
                        if results:
                            if results[0]:
                                ic(results)
                                time.sleep(2.5)
                                write_results(sheet_name, results)

        elif sheet_name == "Cherry Blossom Solar, LLC":
            dfnew = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=2)
            dfnew = dfnew.iloc[1:] #Clear the first row from Data frame as it holds the Units for the Column
            dfnew.iloc[:, 0] = pd.to_datetime(dfnew.iloc[:, 0], errors='coerce') #Convert to DT

            # Filter rows based on time of day
            dfnew = dfnew[(dfnew.iloc[:, 0].dt.time >= pd.to_datetime('09:00 AM').time()) & 
                          (dfnew.iloc[:, 0].dt.time <= pd.to_datetime('05:00 PM').time())]
            cbAVGS = {}
            strings = r'ST(\d+)'
            for column in dfnew.columns[1:]:# Iterate through each column except the first one
                matchl = re.match(strings, column)
                #ic(matchl, column)
                if matchl:
                    identifier = matchl.group(1)
                    if identifier not in cbAVGS:
                        cbAVGS[f'{sheet_name} {identifier}'] = []
            for column in dfnew.columns[1:]:# Iterate through each column except the first one
                match = re.match(strings, column)
                numeric_column = pd.to_numeric(dfnew[column], errors='coerce') #Sets to a number, correcting negative values
                cbavg = numeric_column.mean(skipna=True) 
                if np.isnan(cbavg):
                    cbavg = "No Comms or Device Offline, for a Month"  
                #ic(match, column)
                if match:
                    identifier = match.group(1)
                    cbAVGS[f'{sheet_name} {identifier}'].append([column, cbavg])
            for key, list_avgs in cbAVGS.items():
                if list_avgs:
                    results = []
                    if len(list_avgs) > 1:
                        new_list = [nameavg[1] for nameavg in list_avgs]
                        filtered_list = [cb for cb in new_list if cb != "No Comms or Device Offline, for a Month"]
                        siteMean = pd.Series(filtered_list).mean()
                        comparison_value = siteMean * 0.90
                        for cb in list_avgs:
                            if cb[1] == "No Comms or Device Offline, for a Month":
                                results.append([cb[0], cb[1], key])
                            elif cb[1] < comparison_value:  # Check if the value is less than 5% of siteMean
                                ratio = round((cb[1]/siteMean) * 100, 1)
                                results.append([cb[0], f"Average Performance: {round(cb[1], 2)} KWH  {ratio}% of the Average", key + f" Average KWH: {round(siteMean, 2)}", ratio])        
                        if results:
                            if results[0]:
                                ic(results)
                                time.sleep(2.5)
                                write_results(sheet_name, results)
                    else:
                        ic("I am Unique", key, list_avgs[0][0], list_avgs[0][1])
                        if list_avgs[0][1] == "No Comms or Device Offline, for a Month":
                            results.append([list_avgs[0][0], list_avgs[0][1], key])
                        elif list_avgs[0][1]:
                            kwh = round(list_avgs[0][1], 1)
                            ratio = round((kwh/3) * 100, 1)
                            results.append([list_avgs[0][0], f"Last Month's Average KWH: {kwh} No other CB's with this # of Strings", key, ratio])
                        if results:
                            if results[0]:
                                ic(results)
                                time.sleep(2.5)
                                write_results(sheet_name, results)
    root.destroy() #Finished moving data to Sheet
            
def write_results(sheet_name, data):
    try:
        service = build('sheets', 'v4', credentials=credentials)
        sheet = service.spreadsheets()

        # Write new data to the sheet
        body = {"values": data}
        sheet.values().append(spreadsheetId=cbSheet, range=f"{sheet_name}!A1", valueInputOption="USER_ENTERED", body=body).execute()

    except HttpError as err:
        print(err)


def browse_file():
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")], initialdir=downloads_folder)
    process_file(file_path)

# Create the main window
root = tk.Tk()
root.title("Also Energy CB Check")
try:
    root.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\tracker_3KU_icon.ico")
except Exception as e:
    print(f"Error loading icon: {e}")
browse_button = tk.Button(root, text="Browse files", command=browse_file, width= 40, height= 8)
browse_button.pack(fill='both')

root.mainloop()