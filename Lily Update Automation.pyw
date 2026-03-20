import pyodbc
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime, date, time
import time
from tkinter import *
from tkinter import simpledialog
import pandas as pd
from tkinter import messagebox
import ctypes
import cv2
import numpy as np
import pyautogui
import pytesseract
import re
import sys
import os
import requests
from icecream import ic


from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# Add the parent directory ('NCC Automations') to the Python path
# This allows us to import the 'PythonTools' package from there.
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(parent_dir)
from PythonTools import CREDS, EMAILS, PausableTimer, get_google_credentials

lily_update_file = r"G:\Shared drives\O&M\NCC Automations\Lily Tools\Lily Updates.txt"
path_AHK = r"C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe"
ahk_movedown_path = r"G:\Shared drives\O&M\NCC Automations\Lily Tools\Move Lily Down.ahk"
ahk_moveback_path = r"G:\Shared drives\O&M\NCC Automations\Lily Tools\Move Lily Back.ahk"
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
openEventsSheet = '1byNW58NbhuDotmdzJEFplczsXLAP0iMyS9YVU5FqQgk'
sheet_range= 'Open Events'
PERSONNEL_SHEET = '1Uq8LSya5w6xiAbFhqeCa1upT2Zs-ueQmYFP6cxNkOkI'
LILY_CB_SHEET = '1I4UfcXZYG87b4JmBahjgO4EADph-z4pCk6R9SS7vuW0'


#X-Elio Notification Functions
def send_lily_email(solar_production, irradiance, updates):
    # Email configuration
    sender_email = EMAILS['NCC Desk']
    test = EMAILS['Joseph Lang']
    recipients = EMAILS['Lily Update List']
    print(recipients)
    smtp_port = 587
    smtp_username = EMAILS['NCC Desk']
    smtp_password = CREDS['lilyEmail']

    credentials = get_google_credentials()

    try:
        service = build("sheets", "v4", credentials= credentials)
        sheets = service.spreadsheets()

        open_events = sheets.values().get(spreadsheetId=openEventsSheet, range= sheet_range).execute()
        values = open_events.get("values", [])

    except HttpError as error:
        print("Error accessing the sheet:", error)
    

    # Generate openEvents table
    openEvents_table = "<table style='border-collapse: collapse; border: 1px solid black;'>"
    for row in values:
        openEvents_table += "<tr>"
        for cell in row:
            openEvents_table += f"<td style='border: 1px solid black; padding: 8px; color: black;'>{cell}</td>"
        openEvents_table += "</tr>"
    openEvents_table += "</table>"


    # Compose email message
    msg = MIMEMultipart()
    msg['From'] = sender_email

    if test_var.get() == True:
        msg['To'] = test
    else:
        msg['To'] = ', '.join(recipients)
    ctime = datetime.now()
    print(ctime.minute)
    if 45 > ctime.minute > 15:
        hours = ctime.hour
        mins = '30'
    elif ctime.minute >= 45:
        hours = ctime.hour+1
        mins = '00'
    else:
        hours = ctime.hour
        mins = '00'
    msg['Subject'] = f'{hours}:{mins} Lily Solar Update'
    when = f'{hours}:{mins}'

    today = ctime.strftime("%m/%d/%y")

    # Email body with the provided data
    body = f'''<html>
            <body>
              <table>
                <tr>
                  <td>
                    <table border="1" style="border-collapse: collapse;">
                      <tr style="background-color: lightblue;">
                        <td><strong style="color: black;">Daily Update for Lily Solar - {today}</strong></td>
                      </tr>
                      <tr>
                        <td><strong style="color: black;">Site Performance Overview:</strong><br>
                        <span style="color: black;">As of {when}</span></td>
                      </tr>
                      <tr style="background-color: lightblue;">
                        <td><strong style="color: black;"> Current Solar Production:</strong><br>
                        <span style="color: black;">{solar_production} MW</span></td>
                      </tr>
                      <tr>
                        <td><strong style="color: black;">Irradiance:</strong><br>
                        <span style="color: black;">{irradiance} W/M²</span></td>
                      </tr>
                      <tr style="background-color: lightblue;">
                        <td><strong style="color: black;">Updates:</strong>
                          <ul>
                            <li style="color: black;">{updates}</li>
                          </ul>
                          </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr>
                  <td>
                    {openEvents_table} <!-- Existing table -->
                  </td>
                </tr>
              </table>
            </body>
            </html>'''

    msg.attach(MIMEText(body, 'html'))

    # Connect to Gmail SMTP server and send email
    with smtplib.SMTP("smtp.gmail.com", smtp_port) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.send_message(msg)

    

    with open(lily_update_file, "w") as lilyfile:
        lilyfile.write("No New Updates")



def lily_email_data():
    os.startfile(ahk_movedown_path)
    time.sleep(2) #Wait 2 seconds for window to move.

    myconfig = r"--psm 6 --oem 3"

    try:
        findme = r"G:\Shared drives\O&M\NCC Automations\Lily Tools\Don't Delete\to_find.png"
        location = pyautogui.locateOnScreen(findme)
        left, top, width, height = location
        new_left = int(left+102)
        loca = (new_left, int(top), int(width), int(height))
        
        sc = pyautogui.screenshot(region=loca)
        sc = cv2.cvtColor(np.array(sc), cv2.COLOR_RGB2GRAY)
        _, binary_image = cv2.threshold(sc, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        # Invert the binary image
        inverted_image = cv2.bitwise_not(binary_image)

        sc_file_path = fr"G:\Shared drives\O&M\NCC Automations\Lily Tools\Lily Update Data\Lily_data_{time.strftime('%Y-%m-%d_%H-%M-%S')}.png"
        cv2.imwrite(sc_file_path, inverted_image)
        testy = r"G:\Shared drives\O&M\NCC Automations\Lily Tools\Lily Update Data\Lily_data_2025-04-12_14-15-09.png"
        
        pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        text = pytesseract.image_to_string(inverted_image, config=myconfig)
        pattern = r"(\d+)\.?\,?(\d+)\.?(\d+)?"
        lilys = re.findall(pattern, text)
        print(lilys)
        
        irradiance = "N/A"

        for index, dig_set in enumerate(lilys, start=1):
            if len(dig_set) == 3 and dig_set[2]:
                join_digs = f"{dig_set[0]},{dig_set[1]}.{dig_set[2]}"
            else:
                join_digs = f"{dig_set[0]}.{dig_set[1]}"
            if index == 1:
                power = join_digs
            elif index == 2:
                irradiance = join_digs
        
        if lilys == []:
            power = simpledialog.askfloat("Power not found", "Input MW Value (Cancel for N/A): ", parent=root)
        if power is None: #Is only None if you click Cancel
            power = "N/A"

        if irradiance == "N/A":
            irradiance = simpledialog.askfloat("Irradiance not found", "Input Irradiance Value (Cancel for N/A): ", parent=root)
        if irradiance is None:  #Is only None if you click Cancel
            irradiance = "N/A"
        os.startfile(ahk_moveback_path)

        with open(lily_update_file, "r") as lilyfile:
            update_data = lilyfile.read()

        #Delete ScreenShot
        #os.remove(sc_file_path)
        
        if power and irradiance:
            send_lily_email(power, irradiance, update_data)
    except pytesseract.pytesseract.TesseractNotFoundError:
        print(f"Tesseract is not installed or not in {pytesseract.pytesseract.tesseract_cmd}.")

        
        power = simpledialog.askfloat("Power not found", "Input MW Value (Cancel for N/A): ", parent=root)
        if power is None: #Is only None if you click Cancel
            power = "N/A"
        
        irradiance = simpledialog.askfloat("Irradiance not found", "Input Irradiance Value (Cancel for N/A): ", parent=root)
        if irradiance is None:  #Is only None if you click Cancel
            irradiance = "N/A"

        with open(lily_update_file, "r") as lilyfile:
            update_data = lilyfile.read()
        
        send_lily_email(power, irradiance, update_data)


    except pyautogui.ImageNotFoundException:
        def send_email():
            power = power_var.get()
            irradiance = irradiance_var.get()
            if power and irradiance:
                with open(lily_update_file, "r") as lilyfile:
                    update_data = lilyfile.read()
                send_lily_email(power, irradiance, update_data)
            img_not_fnd_win.destroy()

        img_not_fnd_win = Toplevel()
        img_not_fnd_win.title("Manual Entry Lily Email")
        img_not_fnd_win.wm_attributes("-topmost", True) 
        
        power_var = StringVar(value="00.00")
        irradiance_var = StringVar(value="N/A")

        Label(img_not_fnd_win, text="Power (MW):").grid(row=0, column=0, padx=5, pady=5)
        Entry(img_not_fnd_win, textvariable=power_var, width=10).grid(row=0, column=1, padx=5, pady=5)
        Label(img_not_fnd_win, text="MW").grid(row=0, column=2, padx=5, pady=5)

        Label(img_not_fnd_win, text="Irradiance (W/m²):").grid(row=1, column=0, padx=5, pady=5)
        Entry(img_not_fnd_win, textvariable=irradiance_var, width=10).grid(row=1, column=1, padx=5, pady=5)
        Label(img_not_fnd_win, text="W/m²").grid(row=1, column=2, padx=5, pady=5)

        Button(img_not_fnd_win, text="Cancel", command=img_not_fnd_win.destroy).grid(row=2, column=0, padx=5, pady=5)
        Button(img_not_fnd_win, text="Send Email", command=send_email).grid(row=2, column=1, columnspan=2, padx=5, pady=5)





def lily_ask():
    response = messagebox.askyesno(parent= root, message="Are you sure?", title="Send Lily Email")
    if response == 1:
        lily_email_data()

def update_Personnel_Sheet():
    connect_db, c = dbcnxn()
    
    credentials = get_google_credentials()
    try:
        c.execute('SELECT * from [XELIOupdate]')
        data = c.fetchall()
    except pyodbc.ProgrammingError as e:
        print(f"Error executing query: {e}")
        connect_db.close()
        return
    finally:
        connect_db.close()
    #print(data)
    if data == []:
        messagebox.showinfo(title="Lily Update", message="No Entries found.")
        return
    # Function to format datetime to date or time
    def format_datetime(dt, fmt):
        if dt is None:
            return None
        if fmt == 'date':
            return dt.date()
        elif fmt == 'time':
            return dt.time()

    # Dictionary to hold processed data
    processed_data = {}

    for entry in data:
        if not entry[1] and not entry [2]:
            continue
        log_id = entry[0]
        description = entry[10]

        if entry[3] is None:
            messagebox.showwarning(title="Lily Update Error", message="No Start Date found.\nI will now proceed to the Lily Daily Email App, but neglect to update the Lily Personnel sheet due to the missing data.")
            return
        elif entry[4] is None:
            messagebox.showwarning(title="Lily Update Error", message="No Start Time found, but Start Date is present.\nI will now proceed to the Lily Daily Email App, but neglect to update the Lily Personnel sheet due to the missing data.")
            return

        if entry[5] is None:
            messagebox.showwarning(title="Lily Update Error", message="No End Date found.\nI will now proceed to the Lily Daily Email App, but neglect to update the Lily Personnel sheet due to the missing data.")
            return
        elif entry[6] is None:
            messagebox.showwarning(title="Lily Update Error", message="No End Time found, but End Date is present.\nI will now proceed to the Lily Daily Email App, but neglect to update the Lily Personnel sheet due to the missing data.")
            return
            
        # Format dates and times
        formatted_entry = (
            log_id,
            entry[1], #Company
            entry[2], #Name
            format_datetime(entry[3], 'date'),
            format_datetime(entry[4], 'time'),
            format_datetime(entry[6], 'date'),
            format_datetime(entry[7], 'time'),
            entry[9], #Total Work hours Float value
            description
        )
        dictkey = str(log_id) + formatted_entry[2]
        # Merge descriptions for the same log_id
        if dictkey in processed_data and processed_data[dictkey][2] == entry[2]:
            old_entry = processed_data[dictkey]
            new_description = old_entry[10] + '; ' + description
            new_entry = old_entry[:10] + (new_description,) 
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
            entry[7] if entry[7] else "",  # Column K
            '',  # Column L
            entry[8] if entry[8] else "",  # Column M
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
            spreadsheetId=PERSONNEL_SHEET, 
            range="X-Elio",  # Specify the sheet and range where data should be appended
            valueInputOption="USER_ENTERED", 
            body=body
        ).execute()
        
    except HttpError as error:
        print(error)
    messagebox.showinfo(message="Updated Lily Personnel Sheet!", title="Lily Updater")
    


def dbcnxn():
    #Connect to DB
    db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\Shared drives\O&M\NCC\NCC 039.accdb;'
    connect_db = pyodbc.connect(db)
    c = connect_db.cursor()
    return connect_db, c
#End of X-Elio Tasks


#####
# Internal CB spreadsheet email
def lilyCB_email():
    creds = get_google_credentials()
    try:
        service = build("sheets", "v4", credentials= creds)
        sheets = service.spreadsheets()
        
        # Get sheet ID for "OverView"
        spreadsheet_metadata = sheets.get(spreadsheetId=LILY_CB_SHEET).execute()
        sheet_id = None
        for s in spreadsheet_metadata.get('sheets', []):
            if s.get('properties', {}).get('title') == "OverView":
                sheet_id = s.get('properties', {}).get('sheetId')
                break
        
        if sheet_id is None:
            print("Sheet 'OverView' not found.")
            return

        # Download PDF
        headers = {'Authorization': 'Bearer ' + creds.token}
        pdf_url = f'https://docs.google.com/spreadsheets/d/{LILY_CB_SHEET}/export?format=pdf&gid={sheet_id}&range=A:K'
        
        response = requests.get(pdf_url, headers=headers)
        response.raise_for_status()
        
        # Temp file
        filename = f"Lily_CB_Overview_{datetime.now().strftime('%Y-%m-%d')}.pdf"
        temp_path = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Temp", filename)
        
        with open(temp_path, 'wb') as f:
            f.write(response.content)
            
        # Email configuration
        sender_email = EMAILS['NCC Desk']
        recipients = EMAILS['Lily CB Report List']
        smtp_port = 587
        smtp_username = EMAILS['NCC Desk']
        smtp_password = CREDS['lilyEmail']
        
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = ', '.join(recipients)
        msg['Subject'] = f"Lily Internal CB Report - {datetime.now().strftime('%Y-%m-%d')}"
        
        sheet_link = "https://docs.google.com/spreadsheets/d/1I4UfcXZYG87b4JmBahjgO4EADph-z4pCk6R9SS7vuW0/edit?gid=2071680363#gid=2071680363"
        
        body = f"""
        <html>
            <body>
                <p>Hello Team,</p>
                <p>Please find the Lily Internal Combiner Box Overview report attached.</p>
                <p>You can view the live sheet here: <a href="{sheet_link}">Lily CB Sheet</a></p>
                <p>Thank you,<br>NCC Automation</p>
            </body>
        </html>
        """
        msg.attach(MIMEText(body, 'html'))
        
        with open(temp_path, "rb") as f:
            attach = MIMEApplication(f.read(), _subtype="pdf")
            attach.add_header('Content-Disposition', 'attachment', filename=filename)
            msg.attach(attach)
            
        with smtplib.SMTP("smtp.gmail.com", smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.send_message(msg)
            
        messagebox.showinfo(message="Email sent successfully.", title="Lily CB Report")
        
        # Cleanup
        os.remove(temp_path)

    except HttpError as error:
        print(error)
    except Exception as e:
        print(f"An error occurred: {e}")





#Brandons Tools
def invoicing_noti():
    connect_dbn, c = dbcnxn()
    

    if datetime.today().weekday() == 4: # Friday
        table_name = "NARENCO_mf_Access_logs"
        subject_prefix = "Weekday"
    elif datetime.today().weekday() == 0: # Monday
        table_name = "NARENCO_wkEnd_Access_logs"
        subject_prefix = "Weekend"
    else:
        table_name = "NARENCO_mf_Access_logs"
        subject_prefix = "Weekday - Test"


    try:
        c.execute(f"SELECT * FROM [{table_name}]")
        entries = c.fetchall()
    except pyodbc.ProgrammingError as e:
        print(f"Error executing query for table {table_name}: {e}")
        connect_dbn.close()
        return
    
    connect_dbn.close() # Close the database connection

    if not entries:
        print(f"No entries found in {table_name} for today. Data: {entries}")
        return # No data to email
    processed_data_for_excel = []
    # Define column names based on the expected query result
    column_names = ["Unique ID", "WO No.", "Type", "Site (COUNT)",
                     "Invoice / Quote", "Description", "Last Name",
                       "Qty / Hours", "Rate", "Cost", "Markup %",
                         "Markup Total Cost", "Adddate", "Addtime",
                           "Date", "Site"]


    for entry in entries:
        start_date = entry[0]
        start_time = entry[1]
        end_date = entry[2]
        end_time = entry[3]
        location = entry[4]
        company = entry[5]
        employee = entry[6]
        peepscount = entry[7]
        notes = entry[8]

        # Initialize all 16 Excel columns, defaulting to None (which becomes blank in Excel).
        mileage_row = {col_name: None for col_name in column_names}

        mileage_row["Adddate"] = start_date.date()
        mileage_row["Date"] = start_date.date()
        mileage_row["Addtime"] = start_time.time()
        mileage_row["Type"] = 'Mileage'
        mileage_row["Site"] = location
        mileage_row["Invoice / Quote"] = 'Quote'
        mileage_row["Site (COUNT)"] = 1
        if " " in employee:
            mileage_row["Last Name"] = employee.split(" ")[-1]
        mileage_row["Description"] = notes

        processed_data_for_excel.append(mileage_row)

        # Initialize all 16 Excel columns, defaulting to None (which becomes blank in Excel).
        traveltime_row = {col_name: None for col_name in column_names}

        traveltime_row["Adddate"] = start_date.date()
        traveltime_row["Date"] = start_date.date()
        traveltime_row["Addtime"] = start_time.time()
        traveltime_row["Type"] = 'Travel Time'
        traveltime_row["Site"] = location
        traveltime_row["Invoice / Quote"] = 'Quote'
        traveltime_row["Site (COUNT)"] = 1
        if " " in employee:
            traveltime_row["Last Name"] = employee.split(" ")[-1]
        traveltime_row["Description"] = notes

        processed_data_for_excel.append(traveltime_row)

        # Initialize all 16 Excel columns, defaulting to None (which becomes blank in Excel).
        labor_row = {col_name: None for col_name in column_names}

        labor_row["Adddate"] = start_date.date()
        labor_row["Date"] = start_date.date()
        labor_row["Addtime"] = start_time.time()
        labor_row["Type"] = 'Labor'
        labor_row["Site"] = location
        labor_row["Invoice / Quote"] = 'Quote'
        labor_row["Site (COUNT)"] = 1
        if " " in employee:
            labor_row["Last Name"] = employee.split(" ")[-1]
        labor_row["Description"] = notes

                # Calculate duration for Labor row
        duration_str = "N/A"
        start_datetime_obj = None
        end_datetime_obj = None

        # Create start datetime object
        if start_date is not None and start_time is not None:
            actual_start_date = start_date.date()
            actual_start_time = start_time.time()
            start_datetime_obj = datetime.combine(actual_start_date, actual_start_time)

        # Create end datetime object
        if end_date is not None and end_time is not None:
            actual_end_date = end_date.date()
            actual_end_time = end_time.time()
            end_datetime_obj = datetime.combine(actual_end_date, actual_end_time)
        
        if start_datetime_obj and end_datetime_obj:
            if end_datetime_obj >= start_datetime_obj:
                time_difference = end_datetime_obj - start_datetime_obj
                total_seconds = time_difference.total_seconds()
                total_hours = total_seconds / 3600
                duration_str = f"{total_hours:.2f}"
        
        labor_row["Qty / Hours"] = duration_str # This now holds the calculated duration for Labor


        processed_data_for_excel.append(labor_row)

    # Create a Pandas DataFrame
    df = pd.DataFrame(processed_data_for_excel, columns=column_names)

    # Define Excel file path and name
    excel_filename = f"{subject_prefix}_Access_Logs_{datetime.today().strftime('%Y-%m-%d')}.xlsx"
    # Using a temporary directory is safer for temporary files
    excel_path = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Temp", excel_filename)
    # Ensure the directory exists (optional, temp dir usually exists)
    os.makedirs(os.path.dirname(excel_path), exist_ok=True)

    # Save DataFrame to Excel
    try:
        df.to_excel(excel_path, index=False)
        print(f"Data saved to {excel_path}")
    except Exception as e:
        print(f"Error saving data to Excel: {e}")
        return # Don't send email if file creation failed

    # Email configuration
    sender_email = EMAILS['NCC Desk']
    recipient_email = EMAILS['Joseph Lang']
    smtp_port = 587
    smtp_username = EMAILS['NCC Desk']
    smtp_password = CREDS['shiftsumEmail']

    # Compose email message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = f"Invoicing Automation {subject_prefix} Access Logs - {datetime.today().strftime('%m/%d/%Y')}"

    # Add body text
    body = f"Please find the {subject_prefix} access logs attached."
    msg.attach(MIMEText(body, 'plain'))

    # Attach the Excel file
    try:
        with open(excel_path, "rb") as attachment:
            part = MIMEApplication(attachment.read(), _subtype="xlsx")
            part.add_header('Content-Disposition', 'attachment', filename=excel_filename)
            msg.attach(part)
        print(f"Attached {excel_filename}")
    except Exception as e:
        body += f"\n\nError attaching Excel file: {e}"

    # Send the email
    try:
        with smtplib.SMTP("smtp.gmail.com", smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.send_message(msg)
        print(f"Email sent successfully to {recipient_email}")
    except Exception as e:
        print(f"Error sending email: {e}")
    finally:
        # Clean up the temporary Excel file
        try:
            os.remove(excel_path)
            print(f"Removed temporary file: {excel_path}")
        except OSError as e:
            print(f"Error removing temporary file {excel_path}: {e}")

#Brandon's Functions End



    
myappid = 'NCC.Desk.Functions'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

root = Tk()
root.title("Lily Update Tool")
#root.geometry("280x118")
root.wm_attributes("-topmost", True)
try:
    root.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\email-1.ico")
except Exception as e:
    print(f"Error loading icon: {e}")

cb_var = BooleanVar()
test_var = BooleanVar()

lilyCB_button = Button(root, command= lambda: lilyCB_email(), text= "Lily Internal CB Email", font=("Calibiri", 16), pady= 5, padx= 25, bg='lightblue')
lilyCB_button.pack(pady= 2)

lily_email_button = Button(root, command= lambda: lily_email_data(), text= "Lily Email", font=("Calibiri", 18), pady= 5, padx= 70, bg='yellow')
lily_email_button.pack(pady= 2)

test_frame = Frame(root)
test_frame.pack()
sched_frame = Frame(root)
sched_frame.pack()

test_CB = Checkbutton(test_frame, text= "To: Joseph Only = ✔️", variable=test_var)
test_CB.pack(side= LEFT)
auto_send_CB = Checkbutton(test_frame, text= "Lily Auto Email", variable=cb_var)
auto_send_CB.pack(side=RIGHT)

sched_label = Label(sched_frame, text="Schedule Lily Auto Send")
sched_label.grid(row=1, column= 0, columnspan= 4)

lb1var = IntVar()
ub1var = IntVar()
lb2var = IntVar()
ub2var = IntVar()
lb3var = IntVar()
ub3var = IntVar()
lb4var = IntVar()
ub4var = IntVar()

amlbl = Label(sched_frame, text="Morning Time")
amlbl.grid(row=2, column= 0, columnspan= 4)

lb1lbl = Label(sched_frame, text= "H:")
lb1lbl.grid(row=3, column=0, sticky= E)
lb1 = Entry(sched_frame, textvariable= lb1var, width= 3)
lb1.grid(row=3, column= 1, sticky= W)

lb2lbl = Label(sched_frame, text= "M:")
lb2lbl.grid(row=3, column=2, sticky= E)
lb2 = Entry(sched_frame, textvariable= lb2var, width= 3)
lb2.grid(row=3, column= 3, sticky= W)

ub1lbl = Label(sched_frame, text= "H:")
ub1lbl.grid(row=4, column=0, sticky= E)
ub1 = Entry(sched_frame, textvariable= ub1var, width= 3)
ub1.grid(row=4, column= 1, sticky= W)

ub2lbl = Label(sched_frame, text= "M:")
ub2lbl.grid(row=4, column=2, sticky= E)
ub2 = Entry(sched_frame, textvariable= ub2var, width= 3)
ub2.grid(row=4, column= 3, sticky= W)

pmlbl = Label(sched_frame, text="Afternoon Time")
pmlbl.grid(row=5, column= 0, columnspan= 4)

lb3lbl = Label(sched_frame, text= "H:")
lb3lbl.grid(row=6, column=0, sticky= E)
lb3 = Entry(sched_frame, textvariable= lb3var, width= 3)
lb3.grid(row=6, column= 1, sticky= W)

lb4lbl = Label(sched_frame, text= "M:")
lb4lbl.grid(row=6, column=2, sticky= E)
lb4 = Entry(sched_frame, textvariable= lb4var, width= 3)
lb4.grid(row=6, column= 3, sticky= W)

ub3lbl = Label(sched_frame, text= "H:")
ub3lbl.grid(row=7, column=0, sticky= E)
ub3 = Entry(sched_frame, textvariable= ub3var, width= 3)
ub3.grid(row=7, column= 1, sticky= W)

ub4lbl = Label(sched_frame, text= "M:")
ub4lbl.grid(row=7, column=2, sticky= E)
ub4 = Entry(sched_frame, textvariable= ub4var, width= 3)
ub4.grid(row=7, column= 3, sticky= W)



def check_to_send():
    check = cb_var.get()
    amhl = lb1var.get()
    amml = lb2var.get()
    amhu = ub1var.get()
    ammu = ub2var.get()

    pmhl = lb3var.get()
    pmml = lb4var.get()
    pmhu = ub3var.get()
    pmmu = ub4var.get()


    #Morning Time
    lilyamUtarget_time = datetime.now().replace(hour= amhu, minute= ammu, second=0, microsecond= 0)
    lilyamLtarget_time = datetime.now().replace(hour= amhl, minute= amml, second=0, microsecond= 0)
    #Afternoon Time
    lilypmLtarget_time = datetime.now().replace(hour= pmhl, minute= pmml, second=0, microsecond= 0)
    lilypmUtarget_time = datetime.now().replace(hour= pmhu, minute= pmmu, second=0, microsecond= 0)
    #Now
    currenttime = datetime.now().time()


    if ((lilyamLtarget_time.time() < currenttime < lilyamUtarget_time.time()) or (lilypmLtarget_time.time() < currenttime < lilypmUtarget_time.time())) and check:
        lily_email_data()
    else:
        ic("Criteria not met", check, pmhl, pmml, lilypmLtarget_time.time(), pmhu, pmmu, lilypmUtarget_time.time())
    
    root.after(56000, lambda: check_to_send())



    
#Turn Autosend ON
auto_send_CB.select()
#Set AM Lily Email
lb1var.set(10)
lb2var.set(0)
ub1var.set(10)
ub2var.set(1)

#Set PM Lily Email
open_day = datetime.now().weekday()
if open_day > 4:
    lb3var.set(14)
    lb4var.set(0)
    ub3var.set(14)
    ub4var.set(1)
elif datetime.now().month in [10, 11, 12, 1, 2, 3]:
    lb3var.set(15)
    lb4var.set(0)
    ub3var.set(15)
    ub4var.set(1)
else:
    lb3var.set(17)
    lb4var.set(0)
    ub3var.set(17)
    ub4var.set(1)


if datetime.now().weekday() in [0, 4]:
    invoicing_noti()

response = messagebox.askyesno("Confirmation", "Do you want to update X-Elio's Lily Sign in Spreadsheet?")
if response:
    update_Personnel_Sheet()

check_to_send()
root.mainloop()




########
# Add a Function to send the CB Sheet to Newman and Tom and Zack and Jacob and I only a weekly basis, every Monday. Or Via a Button on the Main GUI???
