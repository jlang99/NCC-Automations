#Tracker Check 
import pandas as pd
import numpy as np
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

# --- Combiner Box Groupings ---
# This dictionary structure will help generalize the analysis process.
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
    with open("aeCB-token.json", "w") as token:
        token.write(credentials.to_json())

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

def update_cb_sheet(metrics):
    """Writes the collected metrics to a Google Sheet."""
    service = build('sheets', 'v4', credentials=credentials)

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
            natural_sort_key(x['cb_name'])[0],
            -max(
                (x['sum_loss'] if x['sum_loss'] is not None and not np.isnan(x['sum_loss']) else -1),
                (x['peak_loss'] if x['peak_loss'] is not None and not np.isnan(x['peak_loss']) else -1),
                (x['abs_peak_loss'] if x['abs_peak_loss'] is not None and not np.isnan(x['abs_peak_loss']) else -1)
            ),
            natural_sort_key(x['cb_name'])
        ))

        data_to_write = []
        for cb_entry in cb_data_for_sorting:
            # Helper to format for JSON
            def format_val(v):
                return None if v is None or (isinstance(v, float) and np.isnan(v)) else v

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
            spreadsheet = service.spreadsheets().get(spreadsheetId=cbSheet).execute()
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
                service.spreadsheets().batchUpdate(spreadsheetId=cbSheet, body={'requests': [append_rows_request]}).execute()
                print(f"Appended {rows_to_add} rows to sheet '{sheet_name}'.")
                time.sleep(1)

            # Clear the sheet before writing new data
            clear_range = f"{sheet_name}!A1:E" # Clear up to the required rows
            service.spreadsheets().values().clear(spreadsheetId=cbSheet, range=clear_range).execute()
            time.sleep(1) # API rate limit

            # Write all data at once
            range_to_update = f'{sheet_name}!A1'
            body = {'values': final_data}
            service.spreadsheets().values().update(
                spreadsheetId=cbSheet,
                range=range_to_update,
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            print(f"Successfully updated all ratios for {sheet_name}.")
            time.sleep(1)
            
            # Now, add the timestamps at the end of the respective columns
            avg_peak_timestamp_str = avg_peak_timestamp.strftime('%Y-%m-%d %H:%M') if pd.notna(avg_peak_timestamp) else ''
            if avg_peak_timestamp_str:
                timestamp_cell_d = f'{sheet_name}!D{len(final_data) + 1}'
                timestamp_body_d = {'values': [[avg_peak_timestamp_str]]}
                service.spreadsheets().values().update(
                    spreadsheetId=cbSheet,
                    range=timestamp_cell_d,
                    valueInputOption='USER_ENTERED',
                    body=timestamp_body_d
                ).execute()
                print(f"Successfully added average peak timestamp for {sheet_name} at {timestamp_cell_d}.")
                time.sleep(1)

            abs_peak_timestamp_str = abs_peak_timestamp.strftime('%Y-%m-%d %H:%M') if pd.notna(abs_peak_timestamp) else ''
            if abs_peak_timestamp_str:
                timestamp_cell_e = f'{sheet_name}!E{len(final_data) + 1}'
                timestamp_body_e = {'values': [[abs_peak_timestamp_str]]}
                service.spreadsheets().values().update(
                    spreadsheetId=cbSheet,
                    range=timestamp_cell_e,
                    valueInputOption='USER_ENTERED',
                    body=timestamp_body_e
                ).execute()
                print(f"Successfully added absolute peak timestamp for {sheet_name} at {timestamp_cell_e}.")
                time.sleep(1)

        except HttpError as err:
            print(f"An error occurred while writing to sheet {sheet_name}: {err}")

def process_file(file_path):
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
        peak_dataframe = df[(df[timestamp_col].dt.hour >= 9) & (df[timestamp_col].dt.hour <= 17)].copy()
        if peak_dataframe.empty:
            print(f"No data for peak hours in {sheet_name}. Skipping.")
            continue

        # Create sum_dataframe for the last 5 days
        max_date = peak_dataframe[timestamp_col].max()
        last_5_days_start_date = max_date - pd.Timedelta(days=5)
        sum_dataframe = peak_dataframe[peak_dataframe[timestamp_col] >= last_5_days_start_date].copy()

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
    update_cb_sheet(all_site_metrics)
    root.destroy()

def browse_file():
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")], initialdir=downloads_folder)
    if file_path:
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